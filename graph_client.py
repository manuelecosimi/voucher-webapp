import os
import msal
import requests
from urllib.parse import quote

MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
MS_DRIVE_USER = os.getenv("MS_DRIVE_USER")  # es. nome.cognome@dominio.it

# Scope per client credentials (sempre .default)
SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"

# Drive base: se c'è MS_DRIVE_USER usiamo il suo OneDrive, altrimenti /me/drive
if MS_DRIVE_USER:
    BASE_DRIVE = f"https://graph.microsoft.com/v1.0/users/{MS_DRIVE_USER}/drive"
else:
    BASE_DRIVE = "https://graph.microsoft.com/v1.0/me/drive"


class GraphClient:
    def __init__(self):
        if not (MS_CLIENT_ID and MS_TENANT_ID and MS_CLIENT_SECRET):
            raise RuntimeError("Variabili MS_CLIENT_ID, MS_TENANT_ID, MS_CLIENT_SECRET mancanti.")
        self._app = msal.ConfidentialClientApplication(
            client_id=MS_CLIENT_ID,
            authority=AUTHORITY,
            client_credential=MS_CLIENT_SECRET,
        )

    def _token(self):
        # Prima cache, poi client credentials
        result = self._app.acquire_token_silent(SCOPES, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" not in result:
            raise RuntimeError(f"MSAL token error: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._token()}"}

    def _norm_path(self, item_path: str) -> str:
        # Garantisce lo slash iniziale e fa URL-encode (spazi, caratteri speciali)
        if not item_path:
            raise ValueError("item_path vuoto")
        if not item_path.startswith("/"):
            item_path = "/" + item_path
        return quote(item_path, safe="/")

    def _get_item_by_path(self, item_path: str):
        path = self._norm_path(item_path)
        url = f"{BASE_DRIVE}/root:{path}"
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
            for chunk in r.iter_content(chunk_size=1024 * 512):
                if chunk:
                    f.write(chunk)

    def upload_excel(self, local_path: str, item_path: str):
        size = os.path.getsize(local_path)
        path = self._norm_path(item_path)

        if size <= 4 * 1024 * 1024:
            # Upload semplice
            url = f"{BASE_DRIVE}/root:{path}:/content"
            with open(local_path, "rb") as f:
                r = requests.put(url, headers=self._headers(), data=f)
            r.raise_for_status()
        else:
            # Upload a sessione (resumable)
            url = f"{BASE_DRIVE}/root:{path}:/createUploadSession"
            r = requests.post(url, headers=self._headers(), json={})
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
                    rr = requests.put(upload_url, headers=headers, data=data)
                    if rr.status_code not in (200, 201, 202):
                        raise RuntimeError(f"Upload chunk failed: {rr.status_code} {rr.text}")
                    sent = end + 1
