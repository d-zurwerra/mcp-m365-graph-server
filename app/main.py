"""
main.py – Oskar MCP Server

Transport: Streamable HTTP (kein SSE – deprecated)
Framework: mcp Python SDK + uvicorn
"""

import os
import logging
from mcp.server.fastmcp import FastMCP
from app.tools.planner import (
    planner_get_tasks,
    planner_create_task,
    planner_update_task,
)
from app.tools.teams import (
    teams_send_channel_message,
    teams_create_chat,
    teams_send_chat_message,
)
from app.tools.sharepoint import (
    sharepoint_get_sites,
    sharepoint_get_list_items,
    sharepoint_create_list_item,
)

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("oskar-mcp-server")

# MCP Server initialisieren
mcp = FastMCP(
    name="Oskar M365 MCP Server",
    instructions=(
        "Dieser Server stellt Tools für Microsoft 365 bereit: "
        "Planner (Aufgaben lesen/erstellen/aktualisieren), "
        "Teams (Nachrichten senden, Chats erstellen), "
        "SharePoint (Sites und Listen lesen/schreiben)."
    ),
)


# ─── PLANNER TOOLS ────────────────────────────────────────────────────────────

@mcp.tool()
async def get_planner_tasks(plan_id: str) -> dict:
    """
    Liest alle Aufgaben eines Microsoft Planner Plans.

    Args:
        plan_id: Die ID des Planner Plans
    """
    logger.info(f"Tool aufgerufen: get_planner_tasks (plan_id={plan_id})")
    return await planner_get_tasks(plan_id)


@mcp.tool()
async def create_planner_task(
    plan_id: str,
    title: str,
    bucket_id: str = None,
    due_date: str = None,
    assigned_user_id: str = None,
) -> dict:
    """
    Erstellt eine neue Aufgabe in einem Microsoft Planner Plan.

    Args:
        plan_id:          Die ID des Planner Plans
        title:            Titel der Aufgabe
        bucket_id:        Optional – ID des Buckets (Spalte im Board)
        due_date:         Optional – Fälligkeitsdatum (ISO 8601, z.B. 2026-05-01T00:00:00Z)
        assigned_user_id: Optional – Entra User ID des Zugewiesenen
    """
    logger.info(f"Tool aufgerufen: create_planner_task (title={title})")
    return await planner_create_task(plan_id, title, bucket_id, due_date, assigned_user_id)


@mcp.tool()
async def update_planner_task(
    task_id: str,
    title: str = None,
    percent_complete: int = None,
    due_date: str = None,
) -> dict:
    """
    Aktualisiert eine bestehende Planner Aufgabe.

    Args:
        task_id:          Die ID der Aufgabe
        title:            Optional – Neuer Titel
        percent_complete: Optional – Fortschritt in Prozent (0, 50 oder 100)
        due_date:         Optional – Neues Fälligkeitsdatum (ISO 8601)
    """
    logger.info(f"Tool aufgerufen: update_planner_task (task_id={task_id})")
    return await planner_update_task(task_id, title, percent_complete, due_date)


# ─── TEAMS TOOLS ──────────────────────────────────────────────────────────────

@mcp.tool()
async def send_teams_channel_message(
    team_id: str,
    channel_id: str,
    message: str,
) -> dict:
    """
    Sendet eine Nachricht in einen Microsoft Teams Channel.

    Args:
        team_id:    Die ID des Teams
        channel_id: Die ID des Channels
        message:    Der Nachrichtentext (HTML erlaubt)
    """
    logger.info(f"Tool aufgerufen: send_teams_channel_message (team_id={team_id})")
    return await teams_send_channel_message(team_id, channel_id, message)


@mcp.tool()
async def create_teams_chat(
    member_user_ids: list[str],
    chat_topic: str = None,
) -> dict:
    """
    Erstellt einen neuen Microsoft Teams Chat (1:1 oder Gruppe).

    Args:
        member_user_ids: Liste von Entra User IDs der Chat-Mitglieder (mind. 2)
        chat_topic:      Optional – Thema des Chats (nur für Gruppen mit 3+ Mitgliedern)
    """
    logger.info(f"Tool aufgerufen: create_teams_chat ({len(member_user_ids)} Mitglieder)")
    return await teams_create_chat(member_user_ids, chat_topic)


@mcp.tool()
async def send_teams_chat_message(
    chat_id: str,
    message: str,
) -> dict:
    """
    Sendet eine Nachricht in einen bestehenden Microsoft Teams Chat.

    Args:
        chat_id: Die ID des Chats
        message: Der Nachrichtentext (HTML erlaubt)
    """
    logger.info(f"Tool aufgerufen: send_teams_chat_message (chat_id={chat_id})")
    return await teams_send_chat_message(chat_id, message)


# ─── SHAREPOINT TOOLS ─────────────────────────────────────────────────────────

@mcp.tool()
async def get_sharepoint_sites(search_term: str = None) -> dict:
    """
    Listet verfügbare SharePoint Sites auf.

    Args:
        search_term: Optional – Suchbegriff für Site-Name
    """
    logger.info(f"Tool aufgerufen: get_sharepoint_sites (search={search_term})")
    return await sharepoint_get_sites(search_term)


@mcp.tool()
async def get_sharepoint_list_items(
    site_id: str,
    list_id: str,
    top: int = 50,
) -> dict:
    """
    Liest Items aus einer SharePoint Liste.

    Args:
        site_id: Die ID der SharePoint Site
        list_id: Die ID der Liste (Name oder GUID)
        top:     Maximale Anzahl Items (Standard: 50, Maximum: 999)
    """
    logger.info(f"Tool aufgerufen: get_sharepoint_list_items (site={site_id}, list={list_id})")
    return await sharepoint_get_list_items(site_id, list_id, top)


@mcp.tool()
async def create_sharepoint_list_item(
    site_id: str,
    list_id: str,
    fields: dict,
) -> dict:
    """
    Erstellt ein neues Item in einer SharePoint Liste.

    Args:
        site_id:  Die ID der SharePoint Site
        list_id:  Die ID der Liste
        fields:   Dictionary mit Feldnamen und Werten, z.B. {"Title": "Neues Item", "Status": "Offen"}
    """
    logger.info(f"Tool aufgerufen: create_sharepoint_list_item (site={site_id}, list={list_id})")
    return await sharepoint_create_list_item(site_id, list_id, fields)


# ─── SERVER START ─────────────────────────────────────────────────────────────

# ASGI App auf Modul-Ebene – wird von uvicorn direkt geladen
_mcp_app = mcp.streamable_http_app()

# Allowed Hosts Middleware – erlaubt alle Hosts (ACA Proxy)
from starlette.middleware.trustedhost import TrustedHostMiddleware
app = TrustedHostMiddleware(_mcp_app, allowed_hosts=["*"])

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    logger.info(f"Oskar MCP Server startet auf Port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
