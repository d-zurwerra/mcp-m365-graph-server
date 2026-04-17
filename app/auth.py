"""
auth.py – Client Secret Auth (temporär für POC)

⚠️ TODO Produktion: Auf Workload Identity Federation umstellen
sobald App Registration im in2success Tenant angelegt ist.
Dann: ClientSecretCredential → ClientAssertionCredential + ManagedIdentityCredential

Aktueller Flow:
  Client Credentials (App Registration DEV-Tenant + Secret) → Graph Token
"""

import os
import logging
from azure.identity import ClientSecretCredential
from azure.core.exceptions import ClientAuthenticationError

logger = logging.getLogger("oskar-mcp-server.auth")

DEV_TENANT_ID   = os.environ["DEV_TENANT_ID"]    # DEV-Tenant
APP_CLIENT_ID   = os.environ["APP_CLIENT_ID"]     # App Registration Client ID
CLIENT_SECRET   = os.environ["CLIENT_SECRET"]     # App Registration Client Secret

GRAPH_SCOPE = "https://graph.microsoft.com/.default"


async def get_graph_token() -> str:
    """Holt einen Graph Access Token via Client Secret Credentials."""
    try:
        credential = ClientSecretCredential(
            tenant_id=DEV_TENANT_ID,
            client_id=APP_CLIENT_ID,
            client_secret=CLIENT_SECRET,
        )
        token = credential.get_token(GRAPH_SCOPE)
        logger.info("Graph Token erfolgreich geholt")
        return token.token

    except ClientAuthenticationError as e:
        logger.error(f"Auth Fehler: {e}")
        raise RuntimeError(f"Token Exchange fehlgeschlagen: {e}") from e


async def get_graph_headers() -> dict:
    """Gibt fertige Authorization Headers für Graph API Calls zurück."""
    token = await get_graph_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
