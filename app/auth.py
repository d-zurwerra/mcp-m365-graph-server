"""
auth.py – Federated Identity Token Flow

Flow:
  1. User-Assigned Managed Identity (ACA) holt OIDC Token (Audience: api://AzureADTokenExchange)
  2. ClientAssertionCredential tauscht diesen Token gegen Graph Token im DEV-Tenant
  3. Resultat: Graph Access Token für DEV-Tenant B

Kein Secret. Kein Zertifikat.
Requires: User-Assigned MI (nicht System-Assigned!)
"""

import os
import logging
from azure.identity import ManagedIdentityCredential, ClientAssertionCredential
from azure.core.exceptions import ClientAuthenticationError

logger = logging.getLogger("oskar-mcp-server.auth")

DEV_TENANT_ID = os.environ["DEV_TENANT_ID"]    # DEV-Tenant B
APP_CLIENT_ID = os.environ["APP_CLIENT_ID"]     # App Registration Client ID (DEV-Tenant B)
MI_CLIENT_ID  = os.environ["MI_CLIENT_ID"]      # User-Assigned MI Client ID

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
AUDIENCE    = "api://AzureADTokenExchange"


def _get_mi_token() -> str:
    """Holt einen OIDC Token von der User-Assigned MI."""
    credential = ManagedIdentityCredential(client_id=MI_CLIENT_ID)
    token = credential.get_token(f"{AUDIENCE}/.default")
    logger.info("MI Token erfolgreich geholt")
    return token.token


async def get_graph_token() -> str:
    """
    Holt einen Graph Access Token via Workload Identity Federation.

    Verwendet ClientAssertionCredential – der offizielle Microsoft-Weg
    für MI as Federated Identity Credential.
    """
    try:
        credential = ClientAssertionCredential(
            tenant_id=DEV_TENANT_ID,
            client_id=APP_CLIENT_ID,
            func=_get_mi_token,  # Callback der den MI Token zurückgibt
        )

        token = credential.get_token(GRAPH_SCOPE)
        logger.info("Graph Token erfolgreich geholt")
        return token.token

    except ClientAuthenticationError as e:
        logger.error(f"Auth Fehler: {e}")
        raise RuntimeError(f"Token Exchange fehlgeschlagen: {e}") from e
    except Exception as e:
        logger.error(f"Unbekannter Fehler: {e}")
        raise


async def get_graph_headers() -> dict:
    """Gibt fertige Authorization Headers für Graph API Calls zurück."""
    token = await get_graph_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
