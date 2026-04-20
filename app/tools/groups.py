"""
tools/groups.py – Microsoft 365 Gruppen & Teams erstellen

Tools:
  - find_user:          User per Name oder Email suchen
  - create_m365_group:  Neue M365 Gruppe anlegen
  - upgrade_to_team:    M365 Gruppe zu Team upgraden
"""

import httpx
import asyncio
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def find_user(search: str) -> dict:
    """
    Sucht einen User per Name oder Email.

    Args:
        search: Name oder Email Adresse des Users
    """
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/users?$filter=startswith(displayName,'{search}') or startswith(mail,'{search}')&$select=id,displayName,mail,userPrincipalName",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        users = [
            {
                "id": u.get("id"),
                "displayName": u.get("displayName"),
                "mail": u.get("mail"),
                "userPrincipalName": u.get("userPrincipalName"),
            }
            for u in data.get("value", [])
        ]
        return {"users": users, "count": len(users)}


async def create_m365_group(
    display_name: str,
    description: str = None,
    owner_user_ids: list[str] = None,
    member_user_ids: list[str] = None,
) -> dict:
    """
    Erstellt eine neue Microsoft 365 Gruppe.
    Erstellt automatisch: SharePoint Site, Planner, Postfach.

    Args:
        display_name:    Name der Gruppe
        description:     Optional – Beschreibung
        owner_user_ids:  Optional – Liste von Entra User IDs als Owners
        member_user_ids: Optional – Liste von Entra User IDs als Members
    """
    headers = await get_graph_headers()

    # Mail nickname aus display_name ableiten (keine Leerzeichen/Sonderzeichen)
    mail_nickname = "".join(c for c in display_name if c.isalnum())[:20]

    body = {
        "displayName": display_name,
        "mailNickname": mail_nickname,
        "mailEnabled": True,
        "securityEnabled": False,
        "groupTypes": ["Unified"],  # Unified = M365 Gruppe
    }

    if description:
        body["description"] = description

    if owner_user_ids:
        body["owners@odata.bind"] = [
            f"https://graph.microsoft.com/v1.0/users/{uid}"
            for uid in owner_user_ids
        ]

    if member_user_ids:
        body["members@odata.bind"] = [
            f"https://graph.microsoft.com/v1.0/users/{uid}"
            for uid in member_user_ids
        ]

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/groups",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        group = response.json()

        return {
            "success": True,
            "groupId": group.get("id"),
            "displayName": group.get("displayName"),
            "mailNickname": group.get("mailNickname"),
            "description": group.get("description"),
        }


async def upgrade_to_team(group_id: str) -> dict:
    """
    Upgraded eine bestehende M365 Gruppe zu einem Microsoft Team.
    Wartet automatisch bis die Gruppe bereit ist (kann 10-30 Sekunden dauern).

    Args:
        group_id: Die ID der M365 Gruppe
    """
    headers = await get_graph_headers()

    # Retry-Loop – Gruppe braucht nach Erstellung etwas Zeit
    max_retries = 6
    retry_delay = 10  # Sekunden

    async with httpx.AsyncClient() as client:
        for attempt in range(max_retries):
            response = await client.put(
                f"{GRAPH_BASE}/groups/{group_id}/team",
                headers=headers,
                json={
                    "memberSettings": {"allowCreatePrivateChannels": True, "allowCreateUpdateChannels": True},
                    "messagingSettings": {"allowUserEditMessages": True, "allowUserDeleteMessages": True},
                    "funSettings": {"allowGiphy": True, "giphyContentRating": "moderate"},
                },
                timeout=30,
            )

            if response.status_code == 201:
                team = response.json()
                return {
                    "success": True,
                    "teamId": team.get("id"),
                    "displayName": team.get("displayName"),
                    "webUrl": team.get("webUrl"),
                }
            elif response.status_code == 409:
                # Gruppe existiert bereits als Team
                return {"success": True, "teamId": group_id, "note": "Gruppe war bereits ein Team"}
            elif response.status_code in [404, 400] and attempt < max_retries - 1:
                # Gruppe noch nicht bereit – warten und retry
                await asyncio.sleep(retry_delay)
                continue
            else:
                raise RuntimeError(
                    f"Team-Upgrade fehlgeschlagen: {response.status_code} – {response.text}"
                )

    raise RuntimeError("Team-Upgrade fehlgeschlagen: Maximale Anzahl Versuche erreicht")
