"""
tools/teams.py – Microsoft Teams Tools

Tools:
  - teams_send_channel_message: Nachricht in einen Channel senden
  - teams_create_chat:          Neuen 1:1 oder Gruppen-Chat erstellen
  - teams_send_chat_message:    Nachricht in bestehenden Chat senden
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def teams_send_channel_message(
    team_id: str,
    channel_id: str,
    message: str,
) -> dict:
    """
    Sendet eine Nachricht in einen Teams Channel.

    Args:
        team_id:    Die ID des Teams
        channel_id: Die ID des Channels
        message:    Der Nachrichtentext (HTML erlaubt)
    """
    headers = await get_graph_headers()

    body = {
        "body": {
            "contentType": "html",
            "content": message,
        }
    }

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        msg = response.json()

        return {
            "success": True,
            "messageId": msg.get("id"),
            "webUrl": msg.get("webUrl"),
        }


async def teams_create_chat(
    member_user_ids: list[str],
    chat_topic: str = None,
) -> dict:
    """
    Erstellt einen neuen Teams Chat (1:1 oder Gruppe).

    Args:
        member_user_ids: Liste von Entra User IDs der Chat-Mitglieder
        chat_topic:      Optional – Thema (nur für Gruppen-Chats)
    """
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

    body = {
        "chatType": chat_type,
        "members": members,
    }

    if chat_topic and chat_type == "group":
        body["topic"] = chat_topic

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/chats",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        chat = response.json()

        return {
            "success": True,
            "chatId": chat.get("id"),
            "chatType": chat.get("chatType"),
            "webUrl": chat.get("webUrl"),
        }


async def teams_send_chat_message(
    chat_id: str,
    message: str,
) -> dict:
    """
    Sendet eine Nachricht in einen bestehenden Teams Chat.

    Args:
        chat_id: Die ID des Chats
        message: Der Nachrichtentext (HTML erlaubt)
    """
    headers = await get_graph_headers()

    body = {
        "body": {
            "contentType": "html",
            "content": message,
        }
    }

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/chats/{chat_id}/messages",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        msg = response.json()

        return {
            "success": True,
            "messageId": msg.get("id"),
        }
