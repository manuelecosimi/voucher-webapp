import os
import json
import requests
from urllib.parse import quote
import msal

# ============== Config comuni =================
ONEDRIVE_EXCEL_PATH = os.getenv("ONEDRIVE_EXCEL_PATH", "/voucher-clienti.xlsx")
MS_DRIVE_USER = os.getenv("MS_DRIVE_USER")  # per account business (UPN), opzionale

# ============== Modalità A: Business (client credentials) ============
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
SCOPES_BUSINESS = ["https://graph.microsoft.com/.default"]

# ============== Modalità B: Personale (Public client + cache MSAL) ===
MSAL_CACHE_BLOB = os.getenv("MSAL_CACHE_BLOB")        # JSON serializzato della cache
MS_ACCOUNT = os.getenv("MS_ACCOUNT")                  # UPN/email dell’account personale (consigliato)
AUTHORITY_PERSONAL = "https://login.microsoftonline.com/consumers"
SCOPES_PERSONAL = ["Files.ReadWrite"]

# piccola utility per richieste http
_DEFAULT_TIMEOUT = 30


class GraphClient:
    """
    Un wrapper minimo per OneDrive via Microsoft Graph.
    Sceglie automaticamente:
      - 'business' se è presente MS_CLIENT_SECRET (client credentials)
      - 'personal' altrimenti (device code+cache salvata in MSAL_CACHE_BLOB)
    """
    def __init__(self):
        self.mode = "business" if MS_CLIENT_SECRET else "personal"
        print(f"[GRAPH] init mode={self.mode}")

        if self.mode == "business":
            # ---- BUSINESS: Confidential client con tenant e client secret
            if not (MS_CLIENT_ID and MS_TENANT_ID and MS_CLIENT_SECRET):
                raise RuntimeError("Config business: servono MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET.")
            self.authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
            self._app = msal.ConfidentialClientApplication(
                client_id=MS_CLIENT_ID,
                authority=self.authority,
                client_credential=MS_CLIENT_SECRET,
            )
            # se MS_DRIVE_USER è impostato, usa il suo drive, altrimenti /me
            if MS_DRIVE_USER:
                self.base_drive = f"https://graph.microsoft.com/v1.0/users/{MS_DRIVE_USER}/drive"
            else:
                self.base_drive = "https://graph.microsoft.com/v1.0/me/drive"

        else:
            # ---- PERSONALE: Public client + cache su MSAL_CACHE_BLOB
            if not MS_CLIENT_ID:
                raise RuntimeError("Config personale: manca MS_CLIENT_ID.")
            if not MSAL_CACHE_BLOB:
                raise RuntimeError("MSAL_CACHE_BLOB assente: genera la cache in locale e incollala su Render.")

            self.authority = AUTHORITY_PERSONAL  # <<< fondamentale!
            self.cache = msal.SerializableTokenCache()
            try:
                self.cache.deserialize(MSAL_CACHE_BLOB)
            except Exception as e:
                raise RuntimeError(f"MSAL_CACHE_BLOB non valido: {e}")

            self._app = msal.PublicClientApplication(
                client_id=MS_CLIENT_ID,
                authority=self.authority,
                token_cache=self.cache,
            )
            # Con account personale usiamo sempre /me
            self.base_drive = "https://graph.microsoft.com/v1.0/me/drive"

    # ------------------------ Token handling ------------------------

    def _token(self):
        if self.mode == "business":
            # prova silent, poi client credentials
            result = self._app.acquire_token_silent(SCOPES_BUSINESS, account=None)
            if not result:
                result = self._app.acquire_token_for_client(scopes=SCOPES_BUSINESS)
        else:
            # PERSONAL: cerca l'account nella cache (se MS_ACCOUNT è dato, usalo)
            accounts = (self._app.get_accounts(username=MS_ACCOUNT)
                        if MS_ACCOUNT else self._app.get_accounts())
            account = accounts[0] if accounts else None
            result = self._app.acquire_token_silent(SCOPES_PERSONAL, account=account)

            # se non c'è token in cache, va rigenerato localmente
            if not result:
                raise RuntimeError("Token non disponibile: rigenera MSAL_CACHE_BLOB (login locale).")

        if "access_token" not in result:
            raise RuntimeError(f"Errore token MSAL: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._token()}"}

    # ------------------------ OneDrive helpers ----------------------

    @staticmethod
    def _norm_path(item_path: str) -> str:
        # garantisce lo slash iniziale e URL-encode (tenendo gli slash)
        if not item_path.startswith("/"):
            item_path = "/" + item_path
        return quote(item_path, safe="/")

    def _get_item_by_path(self, item_path: str):
        path = self._norm_path(item_path)
        url = f"{self.base_drive}/root:{path}"
        r = requests.get(url, headers=self._headers(), timeout=_DEFAULT_TIMEOUT)

        if r.status_code == 401:
            # tipico se la cache non combacia con authority/scope o è scaduta
            raise RuntimeError("401 Unauthorized da Graph. "
                               "Probabile cache MSAL scaduta: rigenera MSAL_CACHE_BLOB.")
        if r.status_code == 404:
            raise FileNotFoundError(f"File non trovato su OneDrive: {item_path}")

        r.raise_for_status()
        return r.json()

    def download_excel(self, local_path: str, item_path: str):
        info = self._get_item_by_path(item_path)
        download_url = info["@microsoft.graph.downloadUrl"]

        with requests.get(download_url, stream=True, timeout=_DEFAULT_TIMEOUT) as r:
            r.raise_for_status()
            with open(local_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 512):
                    if chunk:
                        f.write(chunk)

    def upload_excel(self, local_path: str, item_path: str):
        path = self._norm_path(item_path)
        size = os.path.getsize(local_path)

        if size <= 4 * 1024 * 1024:
            # upload semplice
            url = f"{self.base_drive}/root:{path}:/content"
            with open(local_path, "rb") as f:
                r = requests.put(url, headers=self._headers(), data=f, timeout=_DEFAULT_TIMEOUT)
            if r.status_code == 401:
                raise RuntimeError("401 durante upload: rigenera MSAL_CACHE_BLOB.")
            r.raise_for_status()
            return

        # upload a sessione (chunked)
        url = f"{self.base_drive}/root:{path}:/createUploadSession"
        r = requests.post(url, headers=self._headers(), json={}, timeout=_DEFAULT_TIMEOUT)
        if r.status_code == 401:
            raise RuntimeError("401 durante creazione upload session: rigenera MSAL_CACHE_BLOB.")
        r.raise_for_status()
        upload_url = r.json()["uploadUrl"]

        chunk = 1024 * 1024 * 4  # 4MB
        total = size
        sent = 0
        with open(local_path, "rb") as f:
            while sent < total:
                data = f.read(chunk)
                start = sent
                end = sent + len(data) - 1
                headers = {
                    "Content-Length": str(len(data)),
                    "Content-Range": f"bytes {start}-{end}/{total}",
                }
                rr = requests.put(upload_url, headers=headers, data=data, timeout=_DEFAULT_TIMEOUT)
                if rr.status_code not in (200, 201, 202):
                    raise RuntimeError(f"Upload chunk failed: {rr.status_code} {rr.text}")
                sent = end + 1
