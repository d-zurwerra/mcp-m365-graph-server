"""
tools/teams.py – Microsoft Teams Tools

Tools:
  - get_teams_list:          Alle Teams auflisten
  - get_team_channels:       Channels eines Teams auflisten
  - create_team_channel:     Neuen Channel erstellen
  - get_team_members:        Mitglieder eines Teams auflisten
  - add_team_member:         Mitglied zu Team hinzufügen
  - create_teams_chat:       Neuen Chat erstellen
  - send_teams_chat_message: Nachricht in Chat senden
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def teams_get_list() -> dict:
    """Listet alle Teams im Tenant auf."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        teams = [{"id": t.get("id"), "displayName": t.get("displayName"), "description": t.get("description")} for t in data.get("value", [])]
        return {"teams": teams, "count": len(teams)}


async def teams_get_channels(team_id: str) -> dict:
    """Listet alle Channels eines Teams auf."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/channels",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        channels = [{"id": c.get("id"), "displayName": c.get("displayName"), "description": c.get("description")} for c in data.get("value", [])]
        return {"channels": channels, "count": len(channels)}


async def teams_create_channel(team_id: str, display_name: str, description: str = None) -> dict:
    """Erstellt einen neuen Channel in einem Team."""
    headers = await get_graph_headers()
    body = {"displayName": display_name, "membershipType": "standard"}
    if description:
        body["description"] = description
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        channel = response.json()
        return {"success": True, "channelId": channel.get("id"), "displayName": channel.get("displayName")}


async def teams_get_members(team_id: str) -> dict:
    """Listet alle Mitglieder eines Teams auf."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/teams/{team_id}/members",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        members = [{"id": m.get("id"), "displayName": m.get("displayName"), "email": m.get("email"), "roles": m.get("roles")} for m in data.get("value", [])]
        return {"members": members, "count": len(members)}


async def teams_add_member(team_id: str, user_id: str, role: str = "member") -> dict:
    """Fügt ein Mitglied zu einem Team hinzu."""
    headers = await get_graph_headers()
    body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": ["owner"] if role == "owner" else [],
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')",
    }
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/members",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        return {"success": True, "userId": user_id, "role": role}


async def teams_create_chat(member_user_ids: list[str], chat_topic: str = None) -> dict:
    """Erstellt einen neuen Teams Chat (1:1 oder Gruppe)."""
    headers = await get_graph_headers()
    chat_type = "oneOnOne" if len(member_user_ids) == 2 else "group"
    members = [
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{uid}')",
        }
        for uid in member_user_ids
    ]
    body = {"chatType": chat_type, "members": members}
    if chat_topic and chat_type == "group":
        body["topic"] = chat_topic
    async with httpx.AsyncClient() as client:
        response = await client.post(f"{GRAPH_BASE}/chats", headers=headers, json=body, timeout=30)
        response.raise_for_status()
        chat = response.json()
        return {"success": True, "chatId": chat.get("id"), "chatType": chat.get("chatType")}


async def teams_send_chat_message(chat_id: str, message: str) -> dict:
    """Sendet eine Nachricht in einen bestehenden Teams Chat."""
    headers = await get_graph_headers()
    body = {"body": {"contentType": "html", "content": message}}
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/chats/{chat_id}/messages",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        msg = response.json()
        return {"success": True, "messageId": msg.get("id")}
