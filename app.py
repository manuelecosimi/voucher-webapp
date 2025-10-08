from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime
import os, time
import re  # >>> FIX: ci serve per tenere solo le cifre

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

# >>> OneDrive / MS Graph (rimane importabile; lo useremo solo se USE_CLOUD=1)
import requests
import msal
from dotenv import load_dotenv

# -------------------- ENV & MODALITÀ DEV --------------------
# Carica .env.dev se esiste, altrimenti .env (prod)
env_file = ".env.dev" if os.path.exists(".env.dev") else ".env"
load_dotenv(env_file)

DEV_MODE  = os.getenv("DEV_MODE")  == "1"   # in locale: 1
USE_CLOUD = os.getenv("USE_CLOUD") == "1"   # in locale: 0
SKIP_CLOUD = os.getenv("SKIP_CLOUD") == "1" # per disabilitare le sync se serve
# ------------------------------------------------------------

app = Flask(__name__)

# ---------- HEALTH CHECK PER RENDER ----------
@app.get("/healthz")
def healthz():
    return "ok", 200
# --------------------------------------------

# --- Path robusto all'Excel ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# In dev locale (USE_CLOUD=0) usa sempre il file locale
if DEV_MODE and not USE_CLOUD:
    EXCEL_PATH = os.path.join(BASE_DIR, os.getenv("LOCAL_EXCEL_PATH", "voucher-clienti.xlsx"))
else:
    EXCEL_PATH = os.path.join(BASE_DIR, "voucher-clienti.xlsx")

# >>> Client OneDrive usato solo se USE_CLOUD=1
graph = None
ONEDRIVE_PATH = None
if USE_CLOUD:
    from graph_client import GraphClient
    ONEDRIVE_PATH = os.getenv("ONEDRIVE_EXCEL_PATH", "/voucher-clienti.xlsx")
    graph = GraphClient()

def sync_from_cloud():
    """Scarica l'Excel da OneDrive prima di ogni lettura (no-op in locale o se SKIP_CLOUD=1)."""
    if SKIP_CLOUD or not USE_CLOUD or graph is None:
        return
    try:
        graph.download_excel(EXCEL_PATH, ONEDRIVE_PATH)
    except Exception as e:
        print(f"[SYNC] download da OneDrive saltato: {e}")

def sync_to_cloud() -> bool:
    """
    Carica l'Excel su OneDrive dopo ogni salvataggio.
    Ritorna True se ok, False se fallito (es. 423 Locked).
    """
    if SKIP_CLOUD or not USE_CLOUD or graph is None:
        return True
    try:
        graph.upload_excel(EXCEL_PATH, ONEDRIVE_PATH)
        print("[SYNC] upload completato su OneDrive")
        return True
    except Exception as e:
        print(f"[SYNC] upload verso OneDrive fallito/ritardato: {e}")
        return False
# <<< END NEW


# --- Utility ---

def _parse_money(val):
    """Converte stringhe tipo '€ 1.234,56' in float 1234.56. Ritorna None se vuoto/non valido."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (int, float)):
        try:
            return float(val)
        except Exception:
            return None
    txt = str(val)
    txt = txt.replace('€', '').replace('\xa0', '').replace('\u202f', '').strip()
    if ',' in txt and '.' in txt:
        txt = txt.replace('.', '')
    txt = txt.replace(',', '.')
    try:
        return float(txt) if txt else None
    except Exception:
        return None

def format_valore(valore):
    num = _parse_money(valore)
    if num is None:
        return ""
    return f"€{int(num)}" if float(num).is_integer() else f"€{num:.2f}"

def _norm_card(v):
    """Normalizza il valore della colonna 'N° CARD' a stringa senza decimali."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

# >>> FIX: helper per tenere solo le cifre da input e da Excel
def _digits(s) -> str:
    return "".join(re.findall(r"\d+", str(s))) if s is not None else ""

# >>> FIX: trova la colonna "SERVIZIO" (prefix match, robusto)
def find_service_col(columns):
    for c in columns:
        if c is None:
            continue
        if str(c).strip().upper().startswith("SERVIZIO"):
            return c
    return None

# >>> FIX: trova la colonna NOTE anche se si chiama "SERVIZIO / NOTE"
def find_note_col(columns):
    for c in columns:
        if c is None:
            continue
        s = str(c).strip().upper()
        if "NOTE" in s:
            return c
    return None


# --- Ricerca voucher ---

def cerca_voucher(numero, force_local: bool = False):
    """
    Se force_local=True NON scarica dal cloud prima di leggere (utile subito dopo un salvataggio
    in cui l'upload è fallito: evitiamo di sovrascrivere con la versione vecchia).
    """
    if not force_local:
        sync_from_cloud()

    wb = load_workbook(EXCEL_PATH, data_only=True)
    ws = wb.active

    headers = [
        cell.value.strip() if isinstance(cell.value, str) else cell.value
        for cell in next(ws.iter_rows(min_row=1, max_row=1))
    ]
    ordine_col = headers.index("ORDINE") + 1

    servizio_col = find_service_col(headers)         # >>> FIX
    note_col = find_note_col(headers)                # >>> FIX

    target = _digits(numero)

    found_row = None
    for row in ws.iter_rows(min_row=2, values_only=False):
        ordine_val = row[ordine_col - 1].value
        if _digits(ordine_val) == target:
            found_row = row
            break

    if not found_row:
        wb.close()
        return None

    values = dict(zip(headers, [cell.value for cell in found_row]))
    wb.close()

    # Storico/scalature
    storico = []
    somma_scalature = 0.0
    for col_name in ['1', '2', '3', '4', '5']:
        val = values.get(col_name)
        imp = _parse_money(val)
        if imp is not None:
            storico.append(format_valore(imp))
            somma_scalature += imp

    # Valore totale e residuo
    valore_raw = _parse_money(values.get('VALORE')) or 0.0
    residuo_raw = max(0.0, valore_raw - somma_scalature)

    # Non utilizzabile se tutte le 5 scalature sono valorizzate
    def _filled(x):
        if x is None:
            return False
        if isinstance(x, str) and x.strip() == "":
            return False
        num = _parse_money(x)
        return num is not None and num != 0
    non_utilizzabile = all(_filled(values.get(c)) for c in ['1', '2', '3', '4', '5'])

    # servizio sicuro
    servizio_val = values.get(servizio_col) if servizio_col else ""

    # NOTE: prendi dalla colonna giusta (anche "SERVIZIO / NOTE")
    note_raw = values.get(note_col) or ""

    # data sicura (datetime o stringa)  >>> FIX
    _data_cell = values.get('DATA')
    if isinstance(_data_cell, datetime):
        _data_str = _data_cell.strftime("%d/%m/%Y")
    elif isinstance(_data_cell, str):
        _data_str = _data_cell.strip()
    else:
        _data_str = ""

    # mappa note per appuntamento (in NOTE salviamo "[N] testo")
    notes_map = {}
    if isinstance(note_raw, str):
        for m in re.finditer(r"\[(\d)\]\s*(.+?)(?=(?:\s*\[\d\])|$)", note_raw, flags=re.S):
            idx = int(m.group(1))
            txt = m.group(2).strip().replace("\r", " ").replace("\n", " ")
            notes_map[idx] = txt

    return {
        'numero': target,
        'ordine': values.get('ORDINE'),
        'status': "scaduta" if residuo_raw == 0 else "attiva",
        'valore': format_valore(valore_raw),
        'residuo': format_valore(residuo_raw),
        'servizio': servizio_val or "",
        'card_fisica': "✅" if values.get('N° CARD') else "",
        'box': "✅" if values.get('BOX') else "",
        'card': values.get('CARD') or "",
        'email': values.get('CLIENTE \\ MAIL ORDINE'),
        'data': _data_str,                           # >>> FIX
        'storico': storico,
        'note': note_raw or "",
        'storico_note_map': notes_map,
        'non_utilizzabile': non_utilizzabile
    }


# --- ROUTES ---

@app.route('/', methods=['GET', 'POST'])
def index():
    errore = None

    if request.method == 'POST':
        query = (request.form.get('numero') or '').strip()
        query_digits = _digits(query)

        if query_digits.isdigit() and len(query_digits) == 5:
            risultato = cerca_voucher(query_digits)
            if risultato:
                return render_template('voucher.html', voucher=risultato, by_gift=False)
            else:
                errore = "Voucher non trovato. Controlla il numero inserito."

        elif query_digits.isdigit() and len(query_digits) == 4:
            try:
                sync_from_cloud()
                df = pd.read_excel(EXCEL_PATH)
            except Exception as e:
                return f"Errore lettura Excel: {e}"

            df.columns = df.columns.str.strip()

            if 'N° CARD' not in df.columns or 'ORDINE' not in df.columns:
                return "Colonne 'N° CARD' o 'ORDINE' non trovate nel file."

            q = query_digits.zfill(4)

            def _norm4(v):
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return ""
                s = str(v).strip()
                if s.endswith('.0'):
                    s = s[:-2]
                if s.isdigit():
                    s = s.zfill(4)
                return s

            cards = df['N° CARD'].apply(_norm4)
            sel = cards == q

            if sel.any():
                ordine_val = str(df.loc[sel, 'ORDINE'].iloc[0]).strip()
                ordine_digits = _digits(ordine_val)
                risultato = cerca_voucher(ordine_digits)
                if risultato:
                    return render_template('voucher.html', voucher=risultato, by_gift=True)
                else:
                    errore = "Si è verificato un problema nel recupero della gift."
            else:
                errore = "Gift non ancora assegnata."

        else:
            errore = "Inserisci 5 cifre (voucher) oppure 4 cifre (gift)."

    return render_template('index.html', errore=errore)


@app.route('/gestisci', methods=['GET', 'POST'])
def gestisci():
    numero = (request.args.get('numero') or '').strip()
    if not numero:
        return "Numero voucher mancante"

    numero_digits = _digits(numero)

    try:
        sync_from_cloud()
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        return f"Errore lettura Excel: {e}"

    df.columns = df.columns.str.strip()
    sel = df['ORDINE'].astype(str).apply(_digits) == numero_digits
    if not sel.any():
        return "Voucher non trovato"

    index = df[sel].index[0]
    r = df.loc[index]

    serv_col = find_service_col(df.columns)

    # prima colonna scalatura libera
    prossimo = None
    for i, col in enumerate(['1', '2', '3', '4', '5'], start=1):
        val = r[col]
        if pd.isna(val) or (isinstance(val, str) and val.strip() == ""):
            prossimo = i
            break

    non_utilizzabile = (prossimo is None)
    label_appuntamento = f"Appuntamento {prossimo}" if prossimo else ""

    if request.method == 'POST':
        if non_utilizzabile:
            return "Voucher non più utilizzabile (scalature esaurite)"

        wb = None
        try:
            sync_from_cloud()
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
            excel_row = index + 2  # +1 header

            # --- mappa header -> indice colonna (dinamico) ---
            header_cells = next(ws.iter_rows(min_row=1, max_row=1))
            header_idx = {}
            for i, cell in enumerate(header_cells, start=1):
                key = (cell.value.strip() if isinstance(cell.value, str) else str(cell.value)) if cell.value is not None else ""
                header_idx[key] = i

            # colonne per scalature "1".."5" e NOTE ricavate dagli header
            idx_scal = {str(i): header_idx.get(str(i)) for i in range(1, 6)}
            # >>> FIX: riconosci anche "SERVIZIO / NOTE"
            note_header = find_note_col(header_idx.keys())
            idx_note = header_idx.get(note_header) if note_header else None
            idx_card = header_idx.get('N° CARD')

            # aggiorna card (se inviata)
            if idx_card:
                ws[f'{get_column_letter(idx_card)}{excel_row}'] = request.form.get('card', '')

            nota_form = (request.form.get('note') or '').strip()

            # colonna target della scalatura corrente
            col_idx = idx_scal.get(str(prossimo))
            if not col_idx:
                if wb: wb.close()
                return "Struttura file non valida: colonne 1-5 non trovate."

            col_letter = get_column_letter(col_idx)
            target_addr = f'{col_letter}{excel_row}'

            # indirizzo colonna NOTE (se presente)
            note_addr = f'{get_column_letter(idx_note)}{excel_row}' if idx_note else None

            # lettura sicura del servizio
            serv_val = r.get(serv_col) if serv_col else None

            if serv_val is not None and str(serv_val).strip() != "":
                # checkbox "Servizio effettuato": copia il VALORE intero nella 1a libera
                if request.form.get('servizio_effettuato'):
                    valore = _parse_money(r['VALORE']) or 0.0
                    ws[target_addr] = valore

                    # commento Excel sulla cella giusta
                    if nota_form:
                        cell = ws[target_addr]
                        prev = cell.comment.text if cell.comment else ""
                        txt = f"Servizio effettuato: {nota_form}"
                        cell.comment = Comment(((prev + "\n") if prev else "") + txt, "WebApp")

                # nota generica in NOTE (se esiste)
                if nota_form and note_addr:
                    existing = ws[note_addr].value or ''
                    sep = '\n' if existing else ''
                    ws[note_addr].value = f"{existing}{sep}{nota_form}"

            else:
                # scalatura manuale: importo nel primo slot libero + commento nella stessa cella
                importo_txt = (request.form.get('scalatura') or '').strip()
                # >>> FIX: normalizza input tipo "20,00" o "€ 20"
                importo_txt_norm = importo_txt.replace('€', '').replace('\xa0', ' ').strip().replace(',', '.')
                imp = _parse_money(importo_txt_norm)
                if imp is None or imp <= 0:
                    if wb: wb.close()
                    return "Importo non valido"

                ws[target_addr] = imp

                if nota_form:
                    # nota taggata anche in NOTE (se presente)
                    if note_addr:
                        existing = ws[note_addr].value or ''
                        sep = '\n' if existing else ''
                        ws[note_addr].value = f"{existing}{sep}[{prossimo}] {nota_form}"

                    # commento sulla cella della scalatura
                    cell = ws[target_addr]
                    prev = cell.comment.text if cell.comment else ""
                    txt = f"Appuntamento {prossimo}: {nota_form}"
                    cell.comment = Comment(((prev + "\n") if prev else "") + txt, "WebApp")

            wb.save(EXCEL_PATH)
            wb.close()

            ok = sync_to_cloud()
            return render_template('voucher.html', voucher=cerca_voucher(numero_digits, force_local=not ok), by_gift=False)

        except Exception as e:
            try:
                if wb:
                    wb.close()
            except:
                pass
            return f"Errore scrittura Excel: {e}"

    # GET: prepara dati per la pagina (textarea note vuota)
    return render_template(
        'gestisci.html',
        label_appuntamento=label_appuntamento,
        numero=numero_digits,
        voucher={
            'numero': numero_digits,
            'card': (r['N° CARD'] if not pd.isna(r['N° CARD']) else ''),
            'note': '',
            'servizio': (r.get(serv_col) if (serv_col and not pd.isna(r.get(serv_col))) else ''),
            'valore': format_valore(r['VALORE']),
            'non_utilizzabile': non_utilizzabile,
            'prossimo_appuntamento': prossimo,
            'colonna_attiva': prossimo
        }
    )


# --- NUOVA ROTTA: assegna card ---
@app.route('/assegna-card', methods=['GET', 'POST'])
def assegna_card():
    numero = (request.args.get('numero') or '').strip()
    if not numero:
        return "Numero voucher mancante"

    numero_digits = _digits(numero)

    try:
        sync_from_cloud()
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        return f"Errore lettura Excel: {e}"

    df.columns = df.columns.str.strip()
    sel = df['ORDINE'].astype(str).apply(_digits) == numero_digits
    if not sel.any():
        return "Voucher non trovato"

    index = df[sel].index[0]
    r = df.loc[index]

    if request.method == 'POST':
        card_val = (request.form.get('card') or '').strip()
        if not card_val:
            return render_template(
                'assegna_card.html',
                numero=numero_digits,
                valore_card='',
                errore="Inserisci un numero card."
            )

        wb = None
        try:
            sync_from_cloud()
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
            excel_row = index + 2  # +1 header
            # Colonna B = "N° CARD"
            ws[f'B{excel_row}'] = card_val
            wb.save(EXCEL_PATH)
            wb.close()

            ok = sync_to_cloud()
            return render_template('voucher.html', voucher=cerca_voucher(numero_digits, force_local=not ok), by_gift=False)

        except Exception as e:
            try:
                if wb:
                    wb.close()
            except:
                pass
            return f"Errore scrittura Excel: {e}"

    val_esistente = '' if pd.isna(r['N° CARD']) else str(r['N° CARD'])
    return render_template('assegna_card.html', numero=numero_digits, valore_card=val_esistente, errore=None)

if __name__ == '__main__':
    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "5000"))
    app.run(debug=DEV_MODE, host=host, port=port)
