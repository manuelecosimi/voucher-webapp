# graph_client.py
import os
import json
import pathlib
import requests
import msal
from msal import SerializableTokenCache

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def _env(name, default=None):
    v = os.getenv(name)
    return v if (v is not None and str(v).strip() != "") else default

class GraphClient:
    """
    Client minimale per autenticarsi via Device Code + MSAL
    e leggere/scrivere un file su OneDrive (path basato su root).
    """
    def __init__(self):
        # Default "consumers" funziona meglio per OneDrive personale.
        # Per OneDrive for Business / SharePoint usare il GUID del tenant in .env
        self.tenant_id = _env("MS_TENANT_ID", "consumers")

        self.client_id = _env("MS_CLIENT_ID")  # obbligatorio
        if not self.client_id:
            raise RuntimeError("MS_CLIENT_ID non impostato in .env")

        # Scopes minimi: niente openid/profile/offline_access (MSAL li gestisce da sé)
        # Se in .env hai messo valori multipli, separali con spazio (es. "Files.ReadWrite User.Read")
        scopes_str = _env("MS_SCOPES", "Files.ReadWrite")
        # pulizia e guard-rail contro scope riservati
        reserved = {"openid", "profile", "offline_access"}
        self.scopes = [s for s in scopes_str.split() if s and s not in reserved]

        self.cache_path = _env("MSAL_CACHE_PATH", ".msal_cache.bin")

        # Token cache persistente (file)
        self.cache = SerializableTokenCache()
        if os.path.exists(self.cache_path):
            try:
                self.cache.deserialize(pathlib.Path(self.cache_path).read_text())
            except Exception:
                pass

        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self.cache,
        )
        self.account = None
        self._ensure_account_loaded()

    def _ensure_account_loaded(self):
        accts = self.app.get_accounts()
        self.account = accts[0] if accts else None

    def _save_cache(self):
        if self.cache.has_state_changed:
            pathlib.Path(self.cache_path).write_text(self.cache.serialize())

    # ---------------------------- AUTH ---------------------------- #
    def ensure_token(self):
        """Ritorna un access token valido; se serve, avvia il device code flow."""
        # Prova silent
        if self.account:
            result = self.app.acquire_token_silent(self.scopes, account=self.account)
            if result and "access_token" in result:
                return result["access_token"]

        # Device code flow
        flow = self.app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Impossibile iniziare il device flow: {json.dumps(flow, indent=2)}"
            )

        print("\n=== Autenticazione necessaria ===")
        print("Vai a:", flow.get("verification_uri") or flow.get("verification_uri_complete"))
        print("Inserisci questo codice:", flow["user_code"])
        print("(Dopo l'accesso torna qui, l'operazione continuerà da sola...)\n")

        result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Errore nell'autenticazione: {json.dumps(result, indent=2)}")

        self._ensure_account_loaded()
        self._save_cache()
        return result["access_token"]

    # ---------------------------- FILES --------------------------- #
    @staticmethod
    def _normalize_path(drive_path: str) -> str:
        # Deve iniziare con "/"
        p = drive_path.strip()
        if not p.startswith("/"):
            p = "/" + p
        return p

    def download_excel(self, local_path: str, drive_path: str):
        """Scarica da OneDrive -> file locale (sovrascrive se esiste)."""
        path = self._normalize_path(drive_path)
        token = self.ensure_token()
        url = f"{GRAPH_BASE}/me/drive/root:{path}:/content"

        resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
        if resp.status_code == 200:
            with open(local_path, "wb") as f:
                f.write(resp.content)
        elif resp.status_code == 404:
            # File non esiste ancora in cloud: non fare nulla (lascia il locale)
            return
        else:
            raise RuntimeError(f"Download fallito ({resp.status_code}): {resp.text}")

    def upload_excel(self, local_path: str, drive_path: str):
        """Carica file locale -> OneDrive (PUT su /content)."""
        if not os.path.exists(local_path):
            raise FileNotFoundError(local_path)

        path = self._normalize_path(drive_path)
        token = self.ensure_token()
        url = f"{GRAPH_BASE}/me/drive/root:{path}:/content"

        with open(local_path, "rb") as f:
            data = f.read()

        resp = requests.put(
            url,
            headers={"Authorization": f"Bearer {token}"},
            data=data,
        )
        if resp.status_code not in (200, 201):
            raise RuntimeError(f"Upload fallito ({resp.status_code}): {resp.text}")
