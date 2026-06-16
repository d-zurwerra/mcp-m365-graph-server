"""
github_auth.py – GitHub App Authentication für den Oskar MCP Server

Ersetzt den bisherigen PAT-Ansatz. Generiert kurzlebige Installation Access Tokens
(max. 1h) via GitHub App JWT. Kein manuelles Token-Erneuern nötig.

Umgebungsvariablen (als Azure Container App Secrets):
  GITHUB_APP_ID           – Numerische App ID (z.B. "123456")
  GITHUB_INSTALLATION_ID  – Numerische Installation ID (z.B. "78901234")
  GITHUB_PRIVATE_KEY      – Inhalt der .pem Datei (inkl. BEGIN/END Zeilen)
"""

import os
import time
import logging
import jwt
import httpx

logger = logging.getLogger("oskar-mcp-server.github_auth")


def _generate_jwt() -> str:
    """Generiert ein kurzlebiges JWT (9 Min.) zum Authentifizieren als GitHub App."""
    app_id = os.environ["GITHUB_APP_ID"]
    private_key = os.environ["GITHUB_PRIVATE_KEY"]

    now = int(time.time())
    payload = {
        "iat": now - 60,   # 60s Toleranz für Clock Skew
        "exp": now + 540,  # 9 Minuten gültig (Max: 10 Min.)
        "iss": app_id,
    }

    return jwt.encode(payload, private_key, algorithm="RS256")


def get_installation_token() -> str:
    """
    Holt ein kurzlebiges Installation Access Token von GitHub.
    Dieses Token hat die Berechtigungen aus der GitHub App Installation
    (Issues R/W, Contents R, Metadata R).
    """
    installation_id = os.environ["GITHUB_INSTALLATION_ID"]
    app_jwt = _generate_jwt()

    url = f"https://api.github.com/app/installations/{installation_id}/access_tokens"

    try:
        with httpx.Client(timeout=15) as client:
            response = client.post(
                url,
                headers={
                    "Authorization": f"Bearer {app_jwt}",
                    "Accept": "application/vnd.github+json",
                    "X-GitHub-Api-Version": "2022-11-28",
                },
            )
        response.raise_for_status()
        return response.json()["token"]
    except httpx.HTTPStatusError as e:
        logger.error(f"GitHub Installation Token Fehler: {e.response.status_code}")
        raise RuntimeError(f"GitHub Token Exchange fehlgeschlagen: {e}") from e
    except httpx.RequestError as e:
        logger.error(f"GitHub API nicht erreichbar: {e}")
        raise RuntimeError(f"GitHub API Verbindung fehlgeschlagen: {e}") from e


def get_github_headers() -> dict:
    """
    Gibt fertige HTTP-Headers für GitHub API Calls zurück.
    Bei jedem API Call frisch aufrufen – nicht cachen.
    Die Funktion holt automatisch ein frisches Token.
    """
    token = get_installation_token()
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
