import os, json, msal
from dotenv import load_dotenv
load_dotenv()  # legge MS_CLIENT_ID dal tuo .env se presente

CLIENT_ID = os.getenv("MS_CLIENT_ID") or input("MS_CLIENT_ID: ").strip()
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite"]

cache = msal.SerializableTokenCache()
app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY, token_cache=cache)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise RuntimeError("Impossibile avviare device flow: " + json.dumps(flow, indent=2))

print("\n=== Autenticazione necessaria ===")
print("Vai su:", flow["verification_uri"])
print("Inserisci questo codice:", flow["user_code"])
print("(Dopo il login torna qui: lo script continuerà da solo...)\n")

result = app.acquire_token_by_device_flow(flow)  # attende finché completi il login
if "access_token" not in result:
    raise RuntimeError("Errore auth: " + json.dumps(result, indent=2))

blob = cache.serialize()
# salva anche su file per sicurezza
with open("msal_cache.json", "w", encoding="utf-8") as f:
    f.write(blob)
print("\n=== FATTO ===")
print("Cache salvata in msal_cache.json")
