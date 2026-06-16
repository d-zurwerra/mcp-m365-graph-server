"""
oauth_middleware.py – OAuth2 / Entra ID Token-Validierung

Validiert eingehende Bearer Tokens gegen Microsoft Entra ID.
Nur Requests mit gültigem Token kommen durch — alle anderen erhalten 401.

Umgebungsvariablen:
  ENTRA_TENANT_ID  – Tenant ID des M365 Tenants
  ENTRA_CLIENT_ID  – Application (client) ID der App Registration `oskar-mcp-server-api`
"""

import asyncio
import logging
import jwt
from jwt import PyJWKClient
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse

logger = logging.getLogger("oskar-mcp-server.oauth")


class OAuth2Middleware(BaseHTTPMiddleware):
    def __init__(self, app, tenant_id: str, client_id: str):
        super().__init__(app)
        self.tenant_id = tenant_id
        self.client_id = client_id

        # Öffentliche Schlüssel von Entra ID – wird gecacht
        self.jwks_client = PyJWKClient(
            f"https://login.microsoftonline.com/{tenant_id}/discovery/v2.0/keys",
            cache_keys=True,
        )

        # Gültige Token-Aussteller für diesen Tenant
        self.valid_issuers = [
            f"https://login.microsoftonline.com/{tenant_id}/v2.0",
            f"https://sts.windows.net/{tenant_id}/",
        ]

        # Gültige Audiences: GUID und Application ID URI
        self.valid_audiences = [
            client_id,
            f"api://{client_id}",
        ]

    def _validate_token_sync(self, token: str) -> dict:
        """Validiert Token synchron — wird im Thread-Pool ausgeführt."""
        signing_key = self.jwks_client.get_signing_key_from_jwt(token)

        last_error = None
        for audience in self.valid_audiences:
            try:
                return jwt.decode(
                    token,
                    signing_key.key,
                    algorithms=["RS256"],
                    audience=audience,
                )
            except jwt.InvalidAudienceError as e:
                last_error = e
                continue

        raise last_error

    async def dispatch(self, request, call_next):
        auth_header = request.headers.get("Authorization", "")

        if not auth_header.startswith("Bearer "):
            logger.warning(f"Request ohne Bearer Token abgelehnt: {request.url.path}")
            return JSONResponse(
                {"jsonrpc": "2.0", "id": "auth-error", "error": {"code": -32600, "message": "Unauthorized: Bearer token required"}},
                status_code=401,
            )

        token = auth_header[7:]

        try:
            loop = asyncio.get_event_loop()
            payload = await loop.run_in_executor(None, self._validate_token_sync, token)

            if payload.get("iss") not in self.valid_issuers:
                logger.warning(f"Token mit ungültigem Issuer abgelehnt: {payload.get('iss')}")
                return JSONResponse(
                    {"jsonrpc": "2.0", "id": "auth-error", "error": {"code": -32600, "message": "Unauthorized: Invalid token issuer"}},
                    status_code=401,
                )

            caller = payload.get("app_displayname") or payload.get("appid") or payload.get("sub", "unknown")
            logger.info(f"Token validiert – Aufrufer: {caller}")

        except jwt.ExpiredSignatureError:
            logger.warning("Abgelaufener Token abgelehnt")
            return JSONResponse(
                {"jsonrpc": "2.0", "id": "auth-error", "error": {"code": -32600, "message": "Unauthorized: Token expired"}},
                status_code=401,
            )
        except jwt.InvalidTokenError as e:
            logger.warning(f"Ungültiger Token abgelehnt: {type(e).__name__}")
            return JSONResponse(
                {"jsonrpc": "2.0", "id": "auth-error", "error": {"code": -32600, "message": "Unauthorized: Invalid token"}},
                status_code=401,
            )
        except Exception as e:
            logger.error(f"Unerwarteter Fehler bei Token-Validierung: {e}")
            return JSONResponse(
                {"jsonrpc": "2.0", "id": "auth-error", "error": {"code": -32603, "message": "Internal error during token validation"}},
                status_code=500,
            )

        return await call_next(request)
