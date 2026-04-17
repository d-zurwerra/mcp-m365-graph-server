"""
tools/sharepoint.py – Microsoft SharePoint Tools

Tools:
  - sharepoint_get_sites:       Verfügbare Sites auflisten
  - sharepoint_get_list_items:  Items einer SharePoint Liste lesen
  - sharepoint_create_list_item: Neues Item in einer Liste erstellen
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def sharepoint_get_sites(search_term: str = None) -> dict:
    """
    Listet SharePoint Sites auf, optional gefiltert nach Suchbegriff.

    Args:
        search_term: Optional – Suchbegriff für Site-Name
    """
    headers = await get_graph_headers()

    url = f"{GRAPH_BASE}/sites"
    if search_term:
        url = f"{GRAPH_BASE}/sites?search={search_term}"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()

        sites = []
        for site in data.get("value", []):
            sites.append({
                "id": site.get("id"),
                "name": site.get("name"),
                "displayName": site.get("displayName"),
                "webUrl": site.get("webUrl"),
            })

        return {"sites": sites, "count": len(sites)}


async def sharepoint_get_list_items(
    site_id: str,
    list_id: str,
    top: int = 50,
) -> dict:
    """
    Liest Items aus einer SharePoint Liste.

    Args:
        site_id: Die ID der SharePoint Site
        list_id: Die ID der Liste
        top:     Maximale Anzahl Items (Standard: 50)
    """
    headers = await get_graph_headers()

    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields&$top={top}",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()

        items = []
        for item in data.get("value", []):
            items.append({
                "id": item.get("id"),
                "fields": item.get("fields", {}),
                "createdDateTime": item.get("createdDateTime"),
                "lastModifiedDateTime": item.get("lastModifiedDateTime"),
            })

        return {"items": items, "count": len(items)}


async def sharepoint_create_list_item(
    site_id: str,
    list_id: str,
    fields: dict,
) -> dict:
    """
    Erstellt ein neues Item in einer SharePoint Liste.

    Args:
        site_id:  Die ID der SharePoint Site
        list_id:  Die ID der Liste
        fields:   Dictionary mit Feldnamen und Werten
    """
    headers = await get_graph_headers()

    body = {"fields": fields}

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        item = response.json()

        return {
            "success": True,
            "itemId": item.get("id"),
            "fields": item.get("fields", {}),
        }
