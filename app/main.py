"""
main.py – Oskar MCP Server

Transport: Streamable HTTP
Framework: mcp Python SDK + uvicorn
"""

import os
import logging
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import JSONResponse
from starlette.routing import Route, Mount
from app.tools.planner import planner_get_plans, planner_create_plan, planner_get_tasks, planner_create_task, planner_update_task, planner_create_bucket
from app.tools.teams import teams_get_list, teams_get_channels, teams_create_channel, teams_get_members, teams_add_member, teams_create_chat, teams_send_chat_message
from app.tools.sharepoint import sharepoint_get_sites, sharepoint_get_lists, sharepoint_create_list, sharepoint_get_list_items, sharepoint_create_list_item
from app.tools.groups import find_user as _find_user, create_m365_group as _create_m365_group, upgrade_to_team as _upgrade_to_team, get_m365_groups as _get_m365_groups, get_group_owners as _get_group_owners, add_group_owner as _add_group_owner, get_group_site as _get_group_site, add_group_member as _add_group_member
from app.oauth_middleware import OAuth2Middleware

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("oskar-mcp-server")

mcp = FastMCP(
    name="Oskar M365 MCP Server",
    instructions=(
        "Dieser Server stellt Tools für Microsoft 365 bereit für das interne Kunden-Onboarding: "
        "Teams (auflisten, Channels, Mitglieder), Planner (Plans und Tasks), "
        "SharePoint (Sites, Listen, Items) "
        "und Gruppen (M365 Gruppe + Team erstellen, User suchen)."
    ),
    transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False),
)


# ─── GROUPS & ONBOARDING TOOLS ────────────────────────────────────────────────

@mcp.tool()
async def get_m365_groups(search: str = None) -> dict:
    """
    Listet alle Microsoft 365 Gruppen auf – auch solche die noch kein Team sind.
    Verwende dies IMMER zuerst um zu prüfen ob eine Gruppe bereits existiert,
    bevor du create_m365_group aufrufst.

    Args:
        search: Optional – Suchbegriff für Gruppenname
    """
    logger.info(f"Tool aufgerufen: get_m365_groups (search={search})")
    return await _get_m365_groups(search)


@mcp.tool()
async def add_group_member(group_id: str, user_id: str) -> dict:
    """
    Fügt einen User als Member zu einer M365 Gruppe hinzu.
    Verwende dieses Tool um sicherzustellen dass User den Planner Plan sehen können.
    Wichtig: Owner sind nicht automatisch Members – beide Rollen müssen separat gesetzt werden.

    Args:
        group_id: Die ID der M365 Gruppe
        user_id:  Die Entra User ID
    """
    logger.info(f"Tool aufgerufen: add_group_member (group={group_id}, user={user_id})")
    return await _add_group_member(group_id, user_id)


@mcp.tool()
async def get_group_site(group_id: str) -> dict:
    """
    Holt die SharePoint Site einer M365 Gruppe.
    Jede M365 Gruppe hat automatisch eine zugehörige SharePoint Site.
    Verwende die zurückgegebene siteId für alle SharePoint Operationen.

    Args:
        group_id: Die ID der M365 Gruppe
    """
    logger.info(f"Tool aufgerufen: get_group_site (group_id={group_id})")
    return await _get_group_site(group_id)


@mcp.tool()
async def get_group_owners(group_id: str) -> dict:
    """
    Listet alle Owner einer M365 Gruppe auf.
    Verwende dies um zu prüfen ob ein User bereits Owner ist.

    Args:
        group_id: Die ID der M365 Gruppe
    """
    logger.info(f"Tool aufgerufen: get_group_owners (group_id={group_id})")
    return await _get_group_owners(group_id)


@mcp.tool()
async def add_group_owner(group_id: str, user_id: str) -> dict:
    """
    Fügt einen Owner zu einer M365 Gruppe hinzu.
    Prüfe vorher mit get_group_owners ob der User bereits Owner ist.

    Args:
        group_id: Die ID der M365 Gruppe
        user_id:  Die Entra User ID des neuen Owners
    """
    logger.info(f"Tool aufgerufen: add_group_owner (group={group_id}, user={user_id})")
    return await _add_group_owner(group_id, user_id)


@mcp.tool()
async def find_user(search: str) -> dict:
    """
    Sucht einen User per Name oder Email im Tenant.
    Verwende dies um die User ID für andere Tools zu ermitteln.

    Args:
        search: Name oder Email Adresse des Users
    """
    logger.info(f"Tool aufgerufen: find_user (search={search})")
    return await _find_user(search)


@mcp.tool()
async def create_m365_group(
    display_name: str,
    description: str = None,
    owner_user_ids: list[str] = None,
    member_user_ids: list[str] = None,
    visibility: str = "Private",
) -> dict:
    """
    Erstellt eine neue Microsoft 365 Gruppe für einen Kunden.
    Erstellt automatisch: SharePoint Site, Planner, Postfach.
    Prüfe vorher mit get_m365_groups ob die Gruppe bereits existiert!

    Args:
        display_name:    Name der Gruppe (z.B. "Kunde XY")
        description:     Optional – Beschreibung
        owner_user_ids:  Optional – Liste von Entra User IDs als Owners
        member_user_ids: Optional – Liste von Entra User IDs als Members
        visibility:      "Private" (Standard) oder "Public"
    """
    logger.info(f"Tool aufgerufen: create_m365_group (name={display_name}, visibility={visibility})")
    return await _create_m365_group(display_name, description, owner_user_ids, member_user_ids, visibility)


@mcp.tool()
async def upgrade_to_team(group_id: str) -> dict:
    """
    Upgraded eine M365 Gruppe zu einem Microsoft Team.
    Warte nach create_m365_group ca. 30 Sekunden bevor du dies aufrufst.

    Args:
        group_id: Die ID der M365 Gruppe (aus create_m365_group)
    """
    logger.info(f"Tool aufgerufen: upgrade_to_team (group_id={group_id})")
    return await _upgrade_to_team(group_id)


# ─── TEAMS TOOLS ──────────────────────────────────────────────────────────────

@mcp.tool()
async def get_teams_list() -> dict:
    """Listet alle Microsoft Teams im Tenant auf. Verwende dies um zu prüfen ob ein Team bereits existiert."""
    logger.info("Tool aufgerufen: get_teams_list")
    return await teams_get_list()


@mcp.tool()
async def get_team_channels(team_id: str) -> dict:
    """
    Listet alle Channels eines Microsoft Teams auf.

    Args:
        team_id: Die ID des Teams
    """
    logger.info(f"Tool aufgerufen: get_team_channels (team_id={team_id})")
    return await teams_get_channels(team_id)


@mcp.tool()
async def create_team_channel(team_id: str, display_name: str, description: str = None) -> dict:
    """
    Erstellt einen neuen Channel in einem Microsoft Team.

    Args:
        team_id:      Die ID des Teams
        display_name: Name des Channels
        description:  Optional – Beschreibung
    """
    logger.info(f"Tool aufgerufen: create_team_channel (team={team_id}, name={display_name})")
    return await teams_create_channel(team_id, display_name, description)


@mcp.tool()
async def get_team_members(team_id: str) -> dict:
    """
    Listet alle Mitglieder eines Microsoft Teams auf.

    Args:
        team_id: Die ID des Teams
    """
    logger.info(f"Tool aufgerufen: get_team_members (team_id={team_id})")
    return await teams_get_members(team_id)


@mcp.tool()
async def add_team_member(team_id: str, user_id: str, role: str = "member") -> dict:
    """
    Fügt ein Mitglied zu einem Microsoft Team hinzu.

    Args:
        team_id: Die ID des Teams
        user_id: Die Entra User ID des neuen Mitglieds
        role:    Rolle: 'member' oder 'owner'
    """
    logger.info(f"Tool aufgerufen: add_team_member (team={team_id}, user={user_id})")
    return await teams_add_member(team_id, user_id, role)


@mcp.tool()
async def create_teams_chat(member_user_ids: list[str], chat_topic: str = None) -> dict:
    """
    Erstellt einen neuen Microsoft Teams Chat (1:1 oder Gruppe).

    Args:
        member_user_ids: Liste von Entra User IDs (mind. 2)
        chat_topic:      Optional – Thema (nur für Gruppen)
    """
    logger.info(f"Tool aufgerufen: create_teams_chat ({len(member_user_ids)} Mitglieder)")
    return await teams_create_chat(member_user_ids, chat_topic)


@mcp.tool()
async def send_teams_chat_message(chat_id: str, message: str) -> dict:
    """
    Sendet eine Nachricht in einen bestehenden Microsoft Teams Chat.

    Args:
        chat_id: Die ID des Chats
        message: Der Nachrichtentext (HTML erlaubt)
    """
    logger.info(f"Tool aufgerufen: send_teams_chat_message (chat_id={chat_id})")
    return await teams_send_chat_message(chat_id, message)


# ─── PLANNER TOOLS ────────────────────────────────────────────────────────────

@mcp.tool()
async def get_planner_plans(group_id: str) -> dict:
    """
    Listet alle Planner Plans einer Microsoft 365 Gruppe auf.

    Args:
        group_id: Die ID der M365 Gruppe (= Team ID)
    """
    logger.info(f"Tool aufgerufen: get_planner_plans (group_id={group_id})")
    return await planner_get_plans(group_id)


@mcp.tool()
async def create_planner_plan(group_id: str, title: str) -> dict:
    """
    Erstellt einen neuen Planner Plan in einer Microsoft 365 Gruppe.

    Args:
        group_id: Die ID der M365 Gruppe (= Team ID)
        title:    Titel des Plans
    """
    logger.info(f"Tool aufgerufen: create_planner_plan (group={group_id}, title={title})")
    return await planner_create_plan(group_id, title)


@mcp.tool()
async def create_planner_bucket(plan_id: str, name: str) -> dict:
    """
    Erstellt einen neuen Bucket in einem Planner Plan.
    Buckets sind Kategorien/Spalten in Planner.

    Args:
        plan_id: Die ID des Planner Plans
        name:    Name des Buckets (z.B. "Onboarding")
    """
    logger.info(f"Tool aufgerufen: create_planner_bucket (plan={plan_id}, name={name})")
    return await planner_create_bucket(plan_id, name)


@mcp.tool()
async def get_planner_tasks(plan_id: str) -> dict:
    """
    Liest alle Aufgaben eines Planner Plans.

    Args:
        plan_id: Die ID des Plans
    """
    logger.info(f"Tool aufgerufen: get_planner_tasks (plan_id={plan_id})")
    return await planner_get_tasks(plan_id)


@mcp.tool()
async def create_planner_task(plan_id: str, title: str, bucket_id: str = None, due_date: str = None, assigned_user_id: str = None) -> dict:
    """
    Erstellt eine neue Aufgabe in einem Planner Plan.

    Args:
        plan_id:          Die ID des Plans
        title:            Titel der Aufgabe
        bucket_id:        Optional – ID des Buckets
        due_date:         Optional – Fälligkeitsdatum (ISO 8601)
        assigned_user_id: Optional – Entra User ID
    """
    logger.info(f"Tool aufgerufen: create_planner_task (title={title})")
    return await planner_create_task(plan_id, title, bucket_id, due_date, assigned_user_id)


@mcp.tool()
async def update_planner_task(task_id: str, title: str = None, percent_complete: int = None, due_date: str = None) -> dict:
    """
    Aktualisiert eine bestehende Planner Aufgabe.

    Args:
        task_id:          Die ID der Aufgabe
        title:            Optional – Neuer Titel
        percent_complete: Optional – Fortschritt (0, 50 oder 100)
        due_date:         Optional – Neues Fälligkeitsdatum (ISO 8601)
    """
    logger.info(f"Tool aufgerufen: update_planner_task (task_id={task_id})")
    return await planner_update_task(task_id, title, percent_complete, due_date)


# ─── SHAREPOINT TOOLS ─────────────────────────────────────────────────────────

@mcp.tool()
async def get_sharepoint_sites(search_term: str = None) -> dict:
    """
    Listet verfügbare SharePoint Sites auf.

    Args:
        search_term: Optional – Suchbegriff
    """
    logger.info(f"Tool aufgerufen: get_sharepoint_sites (search={search_term})")
    return await sharepoint_get_sites(search_term)


@mcp.tool()
async def get_sharepoint_lists(site_id: str) -> dict:
    """
    Listet alle Listen einer SharePoint Site auf.

    Args:
        site_id: Die ID der SharePoint Site
    """
    logger.info(f"Tool aufgerufen: get_sharepoint_lists (site_id={site_id})")
    return await sharepoint_get_lists(site_id)


@mcp.tool()
async def create_sharepoint_list(site_id: str, display_name: str, description: str = None, columns: list[dict] = None) -> dict:
    """
    Erstellt eine neue Liste in einer SharePoint Site.

    Args:
        site_id:      Die ID der SharePoint Site
        display_name: Name der Liste
        description:  Optional – Beschreibung
        columns:      Optional – Spalten z.B. [{"name": "Status", "text": {}}]
    """
    logger.info(f"Tool aufgerufen: create_sharepoint_list (site={site_id}, name={display_name})")
    return await sharepoint_create_list(site_id, display_name, description, columns)


@mcp.tool()
async def get_sharepoint_list_items(site_id: str, list_id: str, top: int = 50) -> dict:
    """
    Liest Items aus einer SharePoint Liste.

    Args:
        site_id: Die ID der Site
        list_id: Die ID der Liste
        top:     Max. Anzahl Items (Standard: 50)
    """
    logger.info(f"Tool aufgerufen: get_sharepoint_list_items (site={site_id}, list={list_id})")
    return await sharepoint_get_list_items(site_id, list_id, top)


@mcp.tool()
async def create_sharepoint_list_item(site_id: str, list_id: str, fields: dict) -> dict:
    """
    Erstellt ein neues Item in einer SharePoint Liste.

    Args:
        site_id:  Die ID der Site
        list_id:  Die ID der Liste
        fields:   Felder z.B. {"Title": "Neues Item", "Status": "Offen"}
    """
    logger.info(f"Tool aufgerufen: create_sharepoint_list_item (site={site_id}, list={list_id})")
    return await sharepoint_create_list_item(site_id, list_id, fields)


# ─── SERVER START ─────────────────────────────────────────────────────────────

_tenant_id = os.environ.get("ENTRA_TENANT_ID")
_client_id = os.environ.get("ENTRA_CLIENT_ID")


async def oauth_register(request: Request) -> JSONResponse:
    """
    RFC 7591 Dynamic Client Registration Endpoint.
    Copilot Studio ruft diesen Endpoint auf wenn Auth-Typ 'Dynamisch' gewählt wird.
    Antwortet mit der Entra ID App Registration Client ID — kein Secret nötig (PKCE).
    """
    try:
        body = await request.json()
    except Exception:
        body = {}

    redirect_uris = body.get("redirect_uris", [])
    logger.info(f"OAuth Dynamic Registration angefragt (redirect_uris={redirect_uris})")

    if not _client_id:
        return JSONResponse({"error": "ENTRA_CLIENT_ID nicht konfiguriert"}, status_code=500)

    return JSONResponse({
        "client_id": _client_id,
        "client_secret_expires_at": 0,
        "redirect_uris": redirect_uris,
        "grant_types": ["authorization_code"],
        "response_types": ["code"],
        "token_endpoint_auth_method": "none",  # Public client – kein Secret, PKCE wird verwendet
        "scope": f"api://{_client_id}/access_as_agent",
    }, status_code=201)


# MCP App aufbauen und /register Endpoint ergänzen
_mcp_app = mcp.streamable_http_app()

if _tenant_id and _client_id:
    _mcp_app.add_middleware(OAuth2Middleware, tenant_id=_tenant_id, client_id=_client_id)
    logger.info("OAuth2 Middleware aktiviert (Entra ID Token-Validierung)")

    app = Starlette(routes=[
        Route("/register", oauth_register, methods=["POST"]),
        Mount("/", app=_mcp_app),
    ])
    logger.info("Dynamic Client Registration Endpoint aktiviert (/register)")
else:
    app = _mcp_app
    logger.warning("ENTRA_TENANT_ID oder ENTRA_CLIENT_ID nicht gesetzt – OAuth deaktiviert")

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    logger.info(f"Oskar MCP Server startet auf Port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
