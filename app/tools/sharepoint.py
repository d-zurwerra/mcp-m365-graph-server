"""
tools/sharepoint.py – Microsoft SharePoint Tools

Tools:
  - get_sharepoint_sites:        Sites auflisten
  - get_sharepoint_lists:        Listen einer Site auflisten
  - create_sharepoint_list:      Neue Liste anlegen
  - get_sharepoint_list_items:   Items einer Liste lesen
  - create_sharepoint_list_item: Neues Item erstellen
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def sharepoint_get_sites(search_term: str = None) -> dict:
    """Listet SharePoint Sites auf, optional gefiltert nach Suchbegriff."""
    headers = await get_graph_headers()
    url = f"{GRAPH_BASE}/sites?search={search_term}" if search_term else f"{GRAPH_BASE}/sites?search=*"
    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        sites = [{"id": s.get("id"), "name": s.get("name"), "displayName": s.get("displayName"), "webUrl": s.get("webUrl")} for s in data.get("value", [])]
        return {"sites": sites, "count": len(sites)}


async def sharepoint_get_lists(site_id: str) -> dict:
    """Listet alle Listen einer SharePoint Site auf."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,displayName,description,createdDateTime",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        lists = [{"id": l.get("id"), "displayName": l.get("displayName"), "description": l.get("description")} for l in data.get("value", [])]
        return {"lists": lists, "count": len(lists)}


async def sharepoint_create_list(site_id: str, display_name: str, description: str = None, columns: list[dict] = None) -> dict:
    """
    Erstellt eine neue Liste in einer SharePoint Site.

    Args:
        site_id:      Die ID der SharePoint Site
        display_name: Name der Liste
        description:  Optional – Beschreibung
        columns:      Optional – Spalten z.B. [{"name": "Status", "choice": {"choices": ["Offen", "Erledigt"]}}, {"name": "Verantwortlich", "text": {}}]
    """
    headers = await get_graph_headers()
    body = {
        "displayName": display_name,
        "list": {"template": "genericList"},
    }
    if description:
        body["description"] = description
    if columns:
        # Korrektes Format für SharePoint Graph API Columns
        formatted_columns = []
        for col in columns:
            formatted_col = {"name": col["name"]}
            if "choice" in col:
                formatted_col["choice"] = {
                    "allowTextEntry": False,
                    "choices": col["choice"].get("choices", []),
                    "displayAs": col["choice"].get("displayAs", "dropDownMenu"),
                }
            elif "text" in col:
                formatted_col["text"] = {}
            elif "number" in col:
                formatted_col["number"] = {}
            elif "dateTime" in col:
                formatted_col["dateTime"] = {"format": "dateOnly"}
            formatted_columns.append(formatted_col)
        body["columns"] = formatted_columns

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/sites/{site_id}/lists",
            headers=headers,
            json=body,
            timeout=30,
        )
        if response.status_code not in [200, 201]:
            raise RuntimeError(
                f"Liste konnte nicht erstellt werden: {response.status_code} – {response.text}"
            )
        lst = response.json()
        return {"success": True, "listId": lst.get("id"), "displayName": lst.get("displayName"), "webUrl": lst.get("webUrl")}


async def sharepoint_get_list_items(site_id: str, list_id: str, top: int = 50) -> dict:
    """Liest Items aus einer SharePoint Liste."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items?expand=fields&$top={top}",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        items = [{"id": i.get("id"), "fields": i.get("fields", {}), "createdDateTime": i.get("createdDateTime")} for i in data.get("value", [])]
        return {"items": items, "count": len(items)}


async def sharepoint_create_list_item(site_id: str, list_id: str, fields: dict) -> dict:
    """Erstellt ein neues Item in einer SharePoint Liste."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}/items",
            headers=headers,
            json={"fields": fields},
            timeout=30,
        )
        response.raise_for_status()
        item = response.json()
        return {"success": True, "itemId": item.get("id"), "fields": item.get("fields", {})}
