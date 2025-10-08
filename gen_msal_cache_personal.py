# gen_msal_cache_personal.py
import os, json, msal
from dotenv import load_dotenv
load_dotenv()

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
print("Codice:", flow["user_code"])
print("(Dopo il login torna qui: lo script continua da solo...)\n")

result = app.acquire_token_by_device_flow(flow)
if "access_token" not in result:
    raise RuntimeError("Errore auth: " + json.dumps(result, indent=2))

blob = cache.serialize()
safe_blob = json.dumps(json.loads(blob), separators=(",",":"))  # mono-riga
print("\nLUNGHEZZA:", len(safe_blob))
print("\n--- COPIA TUTTO DA QUI ---\n")
print(safe_blob)
print("\n--- FINO A QUI ---")
print("Account MS:", app.get_accounts()[0]["username"])
