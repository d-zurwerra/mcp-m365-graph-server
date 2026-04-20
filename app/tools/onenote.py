"""
tools/onenote.py – Microsoft OneNote Tools

Tools:
  - get_onenote_notebooks:    Notebooks auflisten
  - create_onenote_notebook:  Neues Notebook erstellen
  - create_onenote_page:      Neue Seite in Notebook erstellen
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def onenote_get_notebooks(group_id: str = None) -> dict:
    """
    Listet OneNote Notebooks auf.

    Args:
        group_id: Optional – ID einer M365 Gruppe / Teams für Gruppen-Notebooks
    """
    headers = await get_graph_headers()
    if group_id:
        url = f"{GRAPH_BASE}/groups/{group_id}/onenote/notebooks"
    else:
        url = f"{GRAPH_BASE}/me/onenote/notebooks"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        notebooks = [{"id": n.get("id"), "displayName": n.get("displayName"), "webUrl": n.get("links", {}).get("oneNoteWebUrl", {}).get("href")} for n in data.get("value", [])]
        return {"notebooks": notebooks, "count": len(notebooks)}


async def onenote_create_notebook(display_name: str, group_id: str = None) -> dict:
    """
    Erstellt ein neues OneNote Notebook.

    Args:
        display_name: Name des Notebooks
        group_id:     Optional – ID einer M365 Gruppe / Teams
    """
    headers = await get_graph_headers()
    if group_id:
        url = f"{GRAPH_BASE}/groups/{group_id}/onenote/notebooks"
    else:
        url = f"{GRAPH_BASE}/me/onenote/notebooks"

    async with httpx.AsyncClient() as client:
        response = await client.post(
            url,
            headers=headers,
            json={"displayName": display_name},
            timeout=30,
        )
        response.raise_for_status()
        notebook = response.json()
        return {
            "success": True,
            "notebookId": notebook.get("id"),
            "displayName": notebook.get("displayName"),
            "webUrl": notebook.get("links", {}).get("oneNoteWebUrl", {}).get("href"),
        }


async def onenote_create_page(notebook_id: str, title: str, content: str, group_id: str = None) -> dict:
    """
    Erstellt eine neue Seite in einem OneNote Notebook.

    Args:
        notebook_id: Die ID des Notebooks
        title:       Titel der Seite
        content:     HTML-Inhalt der Seite
        group_id:    Optional – ID einer M365 Gruppe / Teams
    """
    if group_id:
        url = f"{GRAPH_BASE}/groups/{group_id}/onenote/notebooks/{notebook_id}/sections"
    else:
        url = f"{GRAPH_BASE}/me/onenote/notebooks/{notebook_id}/sections"

    # Erst Section holen oder erstellen
    graph_headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        sections_response = await client.get(url, headers=graph_headers, timeout=30)
        sections_response.raise_for_status()
        sections = sections_response.json().get("value", [])

        if sections:
            section_id = sections[0].get("id")
        else:
            # Default Section erstellen
            if group_id:
                section_url = f"{GRAPH_BASE}/groups/{group_id}/onenote/notebooks/{notebook_id}/sections"
            else:
                section_url = f"{GRAPH_BASE}/me/onenote/notebooks/{notebook_id}/sections"
            section_response = await client.post(
                section_url,
                headers=graph_headers,
                json={"displayName": "Allgemein"},
                timeout=30,
            )
            section_response.raise_for_status()
            section_id = section_response.json().get("id")

        # Seite erstellen
        html_content = f"""<!DOCTYPE html>
<html>
<head><title>{title}</title></head>
<body>{content}</body>
</html>"""

        page_headers = {**graph_headers, "Content-Type": "text/html"}
        if group_id:
            page_url = f"{GRAPH_BASE}/groups/{group_id}/onenote/sections/{section_id}/pages"
        else:
            page_url = f"{GRAPH_BASE}/me/onenote/sections/{section_id}/pages"

        page_response = await client.post(
            page_url,
            headers=page_headers,
            content=html_content.encode("utf-8"),
            timeout=30,
        )
        page_response.raise_for_status()
        page = page_response.json()
        return {
            "success": True,
            "pageId": page.get("id"),
            "title": page.get("title"),
            "webUrl": page.get("links", {}).get("oneNoteWebUrl", {}).get("href"),
        }
