import os
import json
import requests
from urllib.parse import quote
import msal

# ---- Config comuni ----
ONEDRIVE_EXCEL_PATH = os.getenv("ONEDRIVE_EXCEL_PATH", "/voucher-clienti.xlsx")
MS_DRIVE_USER = os.getenv("MS_DRIVE_USER")  # per account business (UPN), opzionale

# --- Modalità A: Business (Con confidential client) ---
MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")

# --- Modalità B: Personale (senza secret) con cache MSAL ---
MSAL_CACHE_BLOB = os.getenv("MSAL_CACHE_BLOB")  # JSON serializzato
MS_ACCOUNT = os.getenv("MS_ACCOUNT")  # email dell'account personale (consigliato)

SCOPES_BUSINESS = ["https://graph.microsoft.com/.default"]
SCOPES_PERSONAL = ["Files.ReadWrite"]

class GraphClient:
    def __init__(self):
        # Se c'è il secret -> Business; altrimenti Personale con token cache
        self.mode = "business" if MS_CLIENT_SECRET else "personal"

        if self.mode == "business":
            if not (MS_CLIENT_ID and MS_TENANT_ID and MS_CLIENT_SECRET):
                raise RuntimeError("Config business: servono MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET.")
            self.authority = f"https://login.microsoftonline.com/{MS_TENANT_ID}"
            self._app = msal.ConfidentialClientApplication(
                client_id=MS_CLIENT_ID,
                authority=self.authority,
                client_credential=MS_CLIENT_SECRET,
            )
            if MS_DRIVE_USER:
                self.base_drive = f"https://graph.microsoft.com/v1.0/users/{MS_DRIVE_USER}/drive"
            else:
                self.base_drive = "https://graph.microsoft.com/v1.0/me/drive"

        else:
            if not MS_CLIENT_ID:
                raise RuntimeError("Config personale: manca MS_CLIENT_ID.")
            self.authority = "https://login.microsoftonline.com/consumers"
            self.cache = msal.SerializableTokenCache()
            if not MSAL_CACHE_BLOB:
                raise RuntimeError("MSAL_CACHE_BLOB assente: genera la cache in locale e incollala su Render.")
            try:
                self.cache.deserialize(MSAL_CACHE_BLOB)
            except Exception as e:
                raise RuntimeError(f"MSAL_CACHE_BLOB non valido: {e}")
            self._app = msal.PublicClientApplication(
                client_id=MS_CLIENT_ID,
                authority=self.authority,
                token_cache=self.cache,
            )
            self.base_drive = "https://graph.microsoft.com/v1.0/me/drive"

    def _token(self):
        if self.mode == "business":
            result = self._app.acquire_token_silent(SCOPES_BUSINESS, account=None)
            if not result:
                result = self._app.acquire_token_for_client(scopes=SCOPES_BUSINESS)
        else:
            accounts = self._app.get_accounts(username=MS_ACCOUNT) if MS_ACCOUNT else self._app.get_accounts()
            account = accounts[0] if accounts else None
            result = self._app.acquire_token_silent(SCOPES_PERSONAL, account=account)
            if not result:
                raise RuntimeError("Token non disponibile: rigenera MSAL_CACHE_BLOB (login locale).")
        if "access_token" not in result:
            raise RuntimeError(f"Errore token MSAL: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._token()}"}

    @staticmethod
    def _norm_path(item_path: str) -> str:
        if not item_path.startswith("/"):
            item_path = "/" + item_path
        return quote(item_path, safe="/")

    def _get_item_by_path(self, item_path: str):
        path = self._norm_path(item_path)
        url = f"{self.base_drive}/root:{path}"
        r = requests.get(url, headers=self._headers())
        if r.status_code == 404:
            raise FileNotFoundError(f"File non trovato su OneDrive: {item_path}")
        r.raise_for_status()
        return r.json()

    def download_excel(self, local_path: str, item_path: str):
        info = self._get_item_by_path(item_path)
        download_url = info["@microsoft.graph.downloadUrl"]
        r = requests.get(download_url, stream=True)
        r.raise_for_status()
        with open(local_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024*512):
                if chunk:
                    f.write(chunk)

    def upload_excel(self, local_path: str, item_path: str):
        path = self._norm_path(item_path)
        size = os.path.getsize(local_path)
        if size <= 4 * 1024 * 1024:
            url = f"{self.base_drive}/root:{path}:/content"
            with open(local_path, "rb") as f:
                r = requests.put(url, headers=self._headers(), data=f)
            r.raise_for_status()
        else:
            url = f"{self.base_drive}/root:{path}:/createUploadSession"
            r = requests.post(url, headers=self._headers(), json={})
            r.raise_for_status()
            upload_url = r.json()["uploadUrl"]
            chunk = 1024 * 1024 * 4
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
                    rr = requests.put(upload_url, headers=headers, data=data)
                    if rr.status_code not in (200, 201, 202):
                        raise RuntimeError(f"Upload chunk failed: {rr.status_code} {rr.text}")
                    sent = end + 1
