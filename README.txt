=========================================
🎫 Voucher Web App – John Barber
=========================================

📝 DESCRIZIONE
-------------
Applicazione web per la consultazione dei voucher acquistati dai clienti.  
Basata su Flask e un file Excel (`voucher-clienti.xlsx`) salvato localmente o su OneDrive.  
L'interfaccia è responsive, elegante e ottimizzata per mobile.

📁 STRUTTURA PROGETTO
---------------------
/app.py                     → file principale Flask  
/templates/index.html       → form per inserimento numero voucher  
/templates/voucher.html     → visualizzazione risultato voucher  
/static/style.css           → stile grafico dark elegante  
/static/logo.png            → logo John Barber  
/voucher-clienti.xlsx       → database Excel dei voucher (da aggiornare)

🚀 COME AVVIARE L'APP
---------------------
1. Assicurati di avere Python 3 installato.
2. Installa i pacchetti richiesti (se non l’hai già fatto):

   pip install flask pandas openpyxl

3. Avvia il server Flask:

   python app.py

4. Apri il browser e vai su:

   http://127.0.0.1:5000/

📌 FUNZIONAMENTO
---------------
- Inserisci il numero del voucher (es. 123 → verrà trasformato in #123)
- Se trovato, viene mostrata la pagina con tutti i dettagli
- Se non trovato, viene mostrato un messaggio di errore
- Il bottone "Cerca un altro voucher" riporta alla homepage

✅ NOTE TECNICHE
---------------
- I dati sono letti da `voucher-clienti.xlsx`
- Il file Excel deve contenere una colonna "ORDINE" con valori tipo "#123"
- Il campo "STATUS" determina il colore verde/rosso
- Lo stile è personalizzabile nel file `style.css`

📧 CONTATTI
-----------
Creato da John Barber Srls – www.johnbarber.it  
Supporto tecnico o modifiche → rivolgersi al team sviluppo




