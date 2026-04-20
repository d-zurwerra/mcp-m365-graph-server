"""
tools/planner.py – Microsoft Planner Tools

Tools:
  - get_planner_plans:   Plans eines Teams auflisten
  - create_planner_plan: Neuen Plan erstellen
  - get_planner_tasks:   Tasks eines Plans lesen
  - create_planner_task: Neue Task erstellen
  - update_planner_task: Task aktualisieren
"""

import httpx
from app.auth import get_graph_headers

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def planner_get_plans(group_id: str) -> dict:
    """Listet alle Planner Plans einer Microsoft 365 Gruppe / Teams auf."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/groups/{group_id}/planner/plans",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        plans = [{"id": p.get("id"), "title": p.get("title"), "createdDateTime": p.get("createdDateTime")} for p in data.get("value", [])]
        return {"plans": plans, "count": len(plans)}


async def planner_create_plan(group_id: str, title: str) -> dict:
    """Erstellt einen neuen Planner Plan in einer Microsoft 365 Gruppe."""
    headers = await get_graph_headers()
    body = {"owner": group_id, "title": title}
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/planner/plans",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        plan = response.json()
        return {"success": True, "planId": plan.get("id"), "title": plan.get("title")}


async def planner_get_tasks(plan_id: str) -> dict:
    """Liest alle Aufgaben eines Planner Plans."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GRAPH_BASE}/planner/plans/{plan_id}/tasks",
            headers=headers,
            timeout=30,
        )
        response.raise_for_status()
        data = response.json()
        tasks = [{"id": t.get("id"), "title": t.get("title"), "percentComplete": t.get("percentComplete"), "dueDateTime": t.get("dueDateTime"), "assignedTo": list(t.get("assignments", {}).keys())} for t in data.get("value", [])]
        return {"tasks": tasks, "count": len(tasks)}


async def planner_create_bucket(plan_id: str, name: str) -> dict:
    """Erstellt einen neuen Bucket in einem Planner Plan."""
    headers = await get_graph_headers()
    body = {"planId": plan_id, "name": name, "orderHint": " !"}
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GRAPH_BASE}/planner/buckets",
            headers=headers,
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        bucket = response.json()
        return {"success": True, "bucketId": bucket.get("id"), "name": bucket.get("name")}


async def planner_create_task(plan_id: str, title: str, bucket_id: str = None, due_date: str = None, assigned_user_id: str = None) -> dict:
    """Erstellt eine neue Aufgabe in einem Planner Plan."""
    headers = await get_graph_headers()
    body = {"planId": plan_id, "title": title}
    if bucket_id:
        body["bucketId"] = bucket_id
    if due_date:
        body["dueDateTime"] = due_date
    if assigned_user_id:
        body["assignments"] = {assigned_user_id: {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}}
    async with httpx.AsyncClient() as client:
        response = await client.post(f"{GRAPH_BASE}/planner/tasks", headers=headers, json=body, timeout=30)
        response.raise_for_status()
        task = response.json()
        return {"success": True, "taskId": task.get("id"), "title": task.get("title")}


async def planner_update_task(task_id: str, title: str = None, percent_complete: int = None, due_date: str = None) -> dict:
    """Aktualisiert eine bestehende Planner Aufgabe."""
    headers = await get_graph_headers()
    async with httpx.AsyncClient() as client:
        get_response = await client.get(f"{GRAPH_BASE}/planner/tasks/{task_id}", headers=headers, timeout=30)
        get_response.raise_for_status()
        etag = get_response.headers.get("ETag")
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
        patch_response = await client.patch(f"{GRAPH_BASE}/planner/tasks/{task_id}", headers=patch_headers, json=body, timeout=30)
        patch_response.raise_for_status()
        return {"success": True, "taskId": task_id, "updated": body}
