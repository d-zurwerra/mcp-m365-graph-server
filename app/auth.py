"""
auth.py – Federated Identity Token Flow
"""

import os
import logging
import httpx
from azure.identity import ManagedIdentityCredential
from azure.core.exceptions import ClientAuthenticationError

logger = logging.getLogger("oskar-mcp-server.auth")

DEV_TENANT_ID = os.environ["DEV_TENANT_ID"]
APP_CLIENT_ID = os.environ["APP_CLIENT_ID"]

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
TOKEN_URL = f"https://login.microsoftonline.com/{DEV_TENANT_ID}/oauth2/v2.0/token"


async def get_graph_token() -> str:
    """
    Holt einen Graph Access Token via Workload Identity Federation.

    Korrekte Methode:
    1. MI Token mit Graph Scope holen (nicht AzureADTokenExchange)
    2. Diesen Token als client_assertion verwenden
    """
    try:
        # Schritt 1: MI Token holen – Scope muss auf den Ziel-Tenant zeigen
        # Für Federated Identity: der MI Token selbst ist die Assertion
        credential = ManagedIdentityCredential()

        # Der Scope für den MI Token muss die App Registration im DEV-Tenant sein
        mi_scope = f"{APP_CLIENT_ID}/.default"
        logger.info(f"Hole MI Token mit Scope: {mi_scope}")

        mi_token = credential.get_token(mi_scope)
        assertion = mi_token.token
        logger.info("MI Token erfolgreich geholt")

        # Schritt 2: Client Credentials mit Federated Assertion
        async with httpx.AsyncClient() as client:
            response = await client.post(
                TOKEN_URL,
                data={
                    "grant_type": "client_credentials",
                    "client_id": APP_CLIENT_ID,
                    "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                    "client_assertion": assertion,
                    "scope": GRAPH_SCOPE,
                },
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                timeout=30,
            )

            logger.info(f"Token Exchange Response: {response.status_code}")
            if response.status_code != 200:
                logger.error(f"Token Exchange Fehler: {response.text}")
                raise RuntimeError(
                    f"Token Exchange fehlgeschlagen: {response.status_code} – {response.text}"
                )

            return response.json()["access_token"]

    except ClientAuthenticationError as e:
        raise RuntimeError(
            f"Managed Identity nicht verfügbar: {e}"
        ) from e


async def get_graph_headers() -> dict:
    """Gibt fertige Authorization Headers für Graph API Calls zurück."""
    token = await get_graph_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
