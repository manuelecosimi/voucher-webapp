import os
import msal
import requests

MS_CLIENT_ID = os.getenv("MS_CLIENT_ID")
MS_TENANT_ID = os.getenv("MS_TENANT_ID")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")

# Scope per client credentials (sempre .default)
SCOPES = ["https://graph.microsoft.com/.default"]
AUTHORITY = f"https://login.microsoftonline.com/{MS_TENANT_ID}"

# OneDrive percorso (es. "/VoucherApp/voucher-clienti.xlsx") viene passato da app.py
# Qui usiamo Graph API con /me/drive se l'app è single-tenant con un user associato.
# Per account “app-only” su drive di un utente specifico, conviene usare /users/{UPN}/drive.
# Manteniamo /me/drive: con application permission punta al drive “predefinito” dell’app owner.
ME_DRIVE = "https://graph.microsoft.com/v1.0/me/drive"

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
        # Prima la cache (nulla), poi client credentials
        result = self._app.acquire_token_silent(SCOPES, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" not in result:
            raise RuntimeError(f"MSAL token error: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._token()}"}

    def _get_item_by_path(self, item_path: str):
        # item_path come "/Cartella/file.xlsx"
        url = f"{ME_DRIVE}/root:{item_path}"
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
        # Usa Upload semplice (<= 4MB) o crea una sessione se il file è più grande
        size = os.path.getsize(local_path)
        if size <= 4 * 1024 * 1024:
            url = f"{ME_DRIVE}/root:{item_path}:/content"
            with open(local_path, "rb") as f:
                r = requests.put(url, headers=self._headers(), data=f)
            r.raise_for_status()
        else:
            # Upload a sessione (resumable)
            url = f"{ME_DRIVE}/root:{item_path}:/createUploadSession"
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
