"""
tools/planner.py – Microsoft Planner Tools

Tools:
  - planner_get_tasks:    Aufgaben eines Plans lesen
  - planner_create_task:  Neue Aufgabe erstellen
  - planner_update_task:  Aufgabe aktualisieren (Status, Titel, etc.)
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def planner_get_tasks(plan_id: str) -> dict:
    """
    Liest alle Aufgaben eines Planner Plans.

    Args:
        plan_id: Die ID des Planner Plans
    """
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/planner/plans/{plan_id}/tasks",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()

        tasks = []
        for task in data.get("value", []):
            tasks.append({
                "id": task.get("id"),
                "title": task.get("title"),
                "percentComplete": task.get("percentComplete"),
                "dueDateTime": task.get("dueDateTime"),
                "assignedTo": list(task.get("assignments", {}).keys()),
                "bucketId": task.get("bucketId"),
            })

        return {"tasks": tasks, "count": len(tasks)}


async def planner_create_task(
    plan_id: str,
    title: str,
    bucket_id: str = None,
    due_date: str = None,
    assigned_user_id: str = None,
) -> dict:
    """
    Erstellt eine neue Aufgabe in einem Planner Plan.

    Args:
        plan_id:          Die ID des Planner Plans
        title:            Titel der Aufgabe
        bucket_id:        Optional – ID des Buckets (Spalte)
        due_date:         Optional – Fälligkeitsdatum (ISO 8601, z.B. 2026-05-01T00:00:00Z)
        assigned_user_id: Optional – Entra User ID des Zugewiesenen
    """
    headers = await get_graph_headers()

    body = {
        "planId": plan_id,
        "title": title,
    }

    if bucket_id:
        body["bucketId"] = bucket_id
    if due_date:
        body["dueDateTime"] = due_date
    if assigned_user_id:
        body["assignments"] = {
            assigned_user_id: {
                "@odata.type": "#microsoft.graph.plannerAssignment",
                "orderHint": " !",
            }
        }

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/planner/tasks",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        task = response.json()

        return {
            "success": True,
            "taskId": task.get("id"),
            "title": task.get("title"),
            "planId": task.get("planId"),
        }


async def planner_update_task(
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
        percent_complete: Optional – Fortschritt (0, 50 oder 100)
        due_date:         Optional – Neues Fälligkeitsdatum (ISO 8601)
    """
    headers = await get_graph_headers()

    # Zuerst ETag holen (Planner braucht das für Updates)
    async with httpx.AsyncClient() as client:
        get_response = await client.get(
            f"{GRAPH_BASE}/planner/tasks/{task_id}",
            headers=headers,
            timeout=30,
        )
        get_response.raise_for_status()
        etag = get_response.headers.get("ETag")

        # Update Body bauen
        body = {}
        if title:
            body["title"] = title
        if percent_complete is not None:
            body["percentComplete"] = percent_complete
        if due_date:
            body["dueDateTime"] = due_date

        if not body:
            return {"success": False, "error": "Keine Änderungen angegeben"}

        patch_headers = {**headers, "If-Match": etag}
        patch_response = await client.patch(
            f"{GRAPH_BASE}/planner/tasks/{task_id}",
            headers=patch_headers,
            json=body,
            timeout=30,
        )
        patch_response.raise_for_status()

        return {"success": True, "taskId": task_id, "updated": body}
