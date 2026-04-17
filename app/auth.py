"""
auth.py – Federated Identity Token Flow

Flow:
  1. Managed Identity (ACA) holt einen OIDC-Token von Entra Tenant A (in2success)
  2. Dieser Token wird gegen Entra Tenant B (DEV-Tenant) getauscht
  3. Resultat: Graph Access Token für Tenant B

Kein Secret. Kein Zertifikat.
"""

import os
import httpx
from azure.identity import ManagedIdentityCredential
from azure.core.exceptions import ClientAuthenticationError

# Umgebungsvariablen (werden in ACA als Environment Variables gesetzt)
DEV_TENANT_ID = os.environ["DEV_TENANT_ID"]       # DEV-Tenant B
APP_CLIENT_ID = os.environ["APP_CLIENT_ID"]         # App Registration Client ID (DEV-Tenant B)

GRAPH_SCOPE = "https://graph.microsoft.com/.default"
TOKEN_EXCHANGE_SCOPE = f"api://AzureADTokenExchange"
TOKEN_URL = f"https://login.microsoftonline.com/{DEV_TENANT_ID}/oauth2/v2.0/token"


async def get_graph_token() -> str:
    """
    Holt einen Graph Access Token via Workload Identity Federation.

    1. MI-Token von Tenant A holen (als OIDC assertion)
    2. Token Exchange gegen Tenant B durchführen
    3. Graph Token zurückgeben
    """
    try:
        # Schritt 1: MI-Token als OIDC Assertion holen
        credential = ManagedIdentityCredential()
        mi_token = credential.get_token(TOKEN_EXCHANGE_SCOPE)
        assertion = mi_token.token

        # Schritt 2: Token Exchange gegen DEV-Tenant B
        async with httpx.AsyncClient() as client:
            response = await client.post(
                TOKEN_URL,
                data={
                    "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                    "client_id": APP_CLIENT_ID,
                    "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                    "client_assertion": assertion,
                    "scope": GRAPH_SCOPE,
                    "requested_token_use": "on_behalf_of",
                },
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                timeout=30,
            )

            if response.status_code != 200:
                raise RuntimeError(
                    f"Token Exchange fehlgeschlagen: {response.status_code} – {response.text}"
                )

            return response.json()["access_token"]

    except ClientAuthenticationError as e:
        raise RuntimeError(
            f"Managed Identity nicht verfügbar (läuft der Server lokal?): {e}"
        ) from e


async def get_graph_headers() -> dict:
    """Gibt fertige Authorization Headers für Graph API Calls zurück."""
    token = await get_graph_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
