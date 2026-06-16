"""
auth.py – Managed Identity Auth für Microsoft Graph

Der MCP Server läuft in Azure Container Apps mit System-assigned Managed Identity.
Die Managed Identity hat die benötigten Graph API Permissions (Phase 5 des Setups).

Kein Client Secret, kein Tenant ID, keine Umgebungsvariablen nötig –
Azure verwaltet die Identität automatisch.
"""

import logging
from azure.identity import ManagedIdentityCredential
from azure.core.exceptions import ClientAuthenticationError

logger = logging.getLogger("oskar-mcp-server.auth")

GRAPH_SCOPE = "https://graph.microsoft.com/.default"

# Credential-Objekt einmal erstellen – azure-identity cached Tokens intern automatisch
_credential = ManagedIdentityCredential()


async def get_graph_token() -> str:
    """Holt einen Graph Access Token via Managed Identity."""
    try:
        token = _credential.get_token(GRAPH_SCOPE)
        logger.debug("Graph Token erfolgreich geholt")
        return token.token
    except ClientAuthenticationError as e:
        logger.error(f"Managed Identity Auth Fehler: {e}")
        raise RuntimeError(f"Token Exchange fehlgeschlagen: {e}") from e


async def get_graph_headers() -> dict:
    """Gibt fertige Authorization Headers für Graph API Calls zurück."""
    token = await get_graph_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
