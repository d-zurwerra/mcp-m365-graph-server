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
    visibility: str = "Private",
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
        "groupTypes": ["Unified"],
        "visibility": visibility,  # "Private" oder "Public"
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
        if response.status_code != 201:
            raise RuntimeError(
                f"Gruppe konnte nicht erstellt werden: {response.status_code} – {response.text}"
            )
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


async def get_m365_groups(search: str = None) -> dict:
    """
    Listet alle M365 Gruppen auf – auch solche die noch kein Team sind.

    Args:
        search: Optional – Suchbegriff für Gruppenname
    """
    headers = await get_graph_headers()
    if search:
        url = f"{GRAPH_BASE}/groups?$filter=groupTypes/any(c:c eq 'Unified') and startswith(displayName,'{search}')&$select=id,displayName,description"
    else:
        url = f"{GRAPH_BASE}/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName,description"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        groups = [
            {
                "id": g.get("id"),
                "displayName": g.get("displayName"),
                "description": g.get("description"),
            }
            for g in data.get("value", [])
        ]
        return {"groups": groups, "count": len(groups)}


async def get_group_owners(group_id: str) -> dict:
    """Listet alle Owner einer M365 Gruppe auf."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/groups/{group_id}/owners?$select=id,displayName,mail",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        owners = [{"id": o.get("id"), "displayName": o.get("displayName"), "mail": o.get("mail")} for o in data.get("value", [])]
        return {"owners": owners, "count": len(owners)}


async def add_group_owner(group_id: str, user_id: str) -> dict:
    """Fügt einen Owner zu einer M365 Gruppe hinzu."""
    headers = await get_graph_headers()
    body = {"@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"}
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/groups/{group_id}/owners/$ref",
            headers=headers,
            json=body,
            timeout=30,
        )
        if response.status_code == 204:
            return {"success": True, "userId": user_id, "note": "Owner erfolgreich hinzugefügt"}
        elif response.status_code == 400 and "already exist" in response.text.lower():
            return {"success": True, "userId": user_id, "note": "User ist bereits Owner"}
        response.raise_for_status()
        return {"success": True, "userId": user_id}


async def get_group_site(group_id: str) -> dict:
    """
    Holt die SharePoint Site einer M365 Gruppe.
    Jede M365 Gruppe hat automatisch eine zugehörige SharePoint Site.

    Args:
        group_id: Die ID der M365 Gruppe
    """
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/groups/{group_id}/sites/root",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        site = response.json()
        return {
            "siteId": site.get("id"),
            "displayName": site.get("displayName"),
            "webUrl": site.get("webUrl"),
            "name": site.get("name"),
        }


async def add_group_member(group_id: str, user_id: str) -> dict:
    """
    Fügt einen User als Member zu einer M365 Gruppe hinzu.
    Verwendet /groups/{id}/members/$ref – funktioniert unabhängig vom Team-Status.

    Args:
        group_id: Die ID der M365 Gruppe
        user_id:  Die Entra User ID
    """
    import logging
    logger = logging.getLogger("oskar-mcp-server.groups")
    headers = await get_graph_headers()
    body = {"@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"}
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/groups/{group_id}/members/$ref",
            headers=headers,
            json=body,
            timeout=30,
        )
        logger.info(f"add_group_member response: {response.status_code}")
        if response.status_code == 204:
            return {"success": True, "userId": user_id, "note": "Member erfolgreich hinzugefügt"}
        elif response.status_code == 400 and "already exist" in response.text.lower():
            return {"success": True, "userId": user_id, "note": "User ist bereits Member"}
        elif response.status_code not in [200, 201, 204]:
            raise RuntimeError(f"Member hinzufügen fehlgeschlagen: {response.status_code} – {response.text}")
        return {"success": True, "userId": user_id}
