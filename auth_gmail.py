"""
Modulo de autenticacion OAuth2 con Gmail y Drive API.
Maneja el flujo de autorizacion y almacenamiento de tokens.
(Adaptado de gmail-comparativos-agent)

Soporta ejecucion en:
- Local (abre navegador para autorizar)
- GitHub Actions (usa token.json desde secrets)
"""
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from config import CREDENTIALS_FILE, TOKEN_FILE, SCOPES


_creds = None


def _obtener_credenciales():
    """Obtiene credenciales OAuth2, reutilizando si ya existen."""
    global _creds

    if _creds and _creds.valid:
        return _creds

    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("[AUTH] Refrescando token expirado...")
            creds.refresh(Request())
        else:
            # En GitHub Actions no hay navegador
            if os.environ.get("GITHUB_ACTIONS"):
                raise RuntimeError(
                    "[AUTH] Token expirado o invalido en GitHub Actions. "
                    "Regenera token.json localmente y actualiza el Secret GOOGLE_TOKEN."
                )

            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    f"Archivo de credenciales no encontrado: {CREDENTIALS_FILE}\n"
                    "Descargalo desde Google Cloud Console > APIs & Services > Credentials\n"
                    "O copia el credentials.json del proyecto gmail-comparativos-agent."
                )

            print("[AUTH] Iniciando flujo de autorizacion OAuth2...")
            print("[AUTH] Se abrira el navegador para autorizar acceso a Gmail + Drive.")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as token:
            token.write(creds.to_json())
        print("[AUTH] Token guardado exitosamente.")

    _creds = creds
    return creds


def autenticar_gmail():
    """Retorna el servicio de Gmail API."""
    creds = _obtener_credenciales()
    service = build("gmail", "v1", credentials=creds)
    print("[AUTH] Conectado a Gmail API correctamente.")
    return service


def autenticar_drive():
    """Retorna el servicio de Google Drive API."""
    creds = _obtener_credenciales()
    service = build("drive", "v3", credentials=creds)
    print("[AUTH] Conectado a Drive API correctamente.")
    return service


def obtener_perfil(service):
    """Obtiene el perfil del usuario autenticado."""
    profile = service.users().getProfile(userId="me").execute()
    return profile.get("emailAddress", "desconocido")
