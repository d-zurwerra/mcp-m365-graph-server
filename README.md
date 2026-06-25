# mcp-m365-graph-server

Ein MCP Server für Microsoft 365 — läuft in Azure Container Apps, authentifiziert sich per Managed Identity gegen die Graph API und spricht per OAuth mit Copilot Studio.

29 Tools für Teams, SharePoint, Planner, Outlook, Kalender und Dateien. Kein Secret für die Graph API. Kein manuelles Token-Rotieren. Fertig schlägt perfekt.

> Dieser Server ist Teil von **Oskar**, dem KI-Agenten hinter [The Raccoon Way](https://the-raccoon-way.de). Wie er entstanden ist und warum bestimmte Entscheidungen so getroffen wurden — das steht im dazugehörigen [Blogartikel](https://the-raccoon-way.de/posts/m365-mcp-server/).

Dieser Code ist im Rahmen von [The Raccoon Way](https://the-raccoon-way.de) entstanden — teils vibe-coded mit KI-Unterstützung, vollständig getestet und in Produktion im Einsatz.

---

## Was dieser Server macht

```
Copilot Studio Agent
    │  OAuth mit Client Secret
    ▼
Azure Container App (dieser Server)
    │  OAuth2 Middleware prüft eingehende Tokens
    │  System-assigned Managed Identity
    ▼
Microsoft Graph API
    │  Teams · SharePoint · Planner · Outlook · Kalender · Dateien
    ▼
Euer M365 Tenant
```

| Strecke | Methode | Läuft ab? |
|---|---|---|
| Container App → Graph API | Managed Identity | ❌ nie |
| Copilot Studio → Container App | OAuth / Client Secret | ✅ alle 24 Monate |
| GHCR → Container App (Image Pull) | Classic PAT | ✅ je nach Laufzeit |

---

## Voraussetzungen

| Was | Wo |
|---|---|
| Microsoft 365 Tenant mit Global Admin | entra.microsoft.com |
| Azure Subscription mit Contributor-Rechten | portal.azure.com |
| Copilot Studio Lizenz (M365 Copilot oder Standalone) | admin.microsoft.com |
| GitHub Account | github.com |
| PowerShell mit Microsoft.Graph Modul | lokal oder Azure Cloud Shell |

---

## Nachbauen

**Wichtig: Erst forken, dann loslegen.**

Das GHCR Package unter meinem Account ist privat — ihr könnt es nicht direkt ziehen. Stattdessen: Repo forken, einmal auf `main` pushen, GitHub Actions baut euer eigenes Image automatisch. Vorher unter **Settings → Actions → General → Workflow permissions** auf **Read and write permissions** stellen.

Die vollständige Schritt-für-Schritt Anleitung gibt es hier:

📄 **[Setup Guide](docs/oskar-m365-mcp-server.md)**

Der Guide erklärt die richtige Reihenfolge, warum Dynamic Client Registration (PKCE) nicht funktioniert, was es mit dem Classic PAT auf sich hat — und welche Stolpersteine auf dem Weg lauern.

---

## Tools

Der Server stellt 29 Tools bereit:

**Teams & Gruppen**
`create_m365_group` · `upgrade_to_team` · `add_team_member` · `create_channel` · `list_teams` · `get_team` · `list_team_members` · `list_channels`

**SharePoint**
`create_sharepoint_site` · `get_sharepoint_site` · `list_sharepoint_sites` · `create_sharepoint_list` · `create_sharepoint_page`

**Planner**
`create_planner_plan` · `create_planner_bucket` · `create_planner_task` · `list_planner_plans` · `list_planner_tasks`

**Outlook & Kalender**
`send_email` · `list_emails` · `get_email` · `create_calendar_event` · `list_calendar_events`

**Dateien**
`list_files` · `get_file` · `create_folder` · `upload_file`

**Benutzer**
`list_users` · `get_user`

> OneNote ist bewusst nicht enthalten — die Graph API Endpoints für OneNote funktionieren im Application-Context (Managed Identity) nicht.

---

## Struktur

```
app/
├── main.py              # Server-Einstiegspunkt, OAuth Middleware, MCP Setup
├── auth.py              # Managed Identity Token für Graph API
├── oauth_middleware.py  # Entra ID Token-Validierung eingehender Requests
└── tools/
    ├── groups.py        # M365 Gruppen & Teams
    ├── sharepoint.py    # SharePoint Sites & Listen
    ├── planner.py       # Planner Plans, Buckets, Tasks
    ├── mail.py          # Outlook Mail
    ├── calendar.py      # Kalender
    ├── files.py         # OneDrive / SharePoint Dateien
    └── users.py         # Benutzer
```

---

## Umgebungsvariablen

| Variable | Pflicht | Beschreibung |
|---|---|---|
| `ENTRA_TENANT_ID` | ✅ | Directory (tenant) ID aus der Entra App Registration |
| `ENTRA_CLIENT_ID` | ✅ | Application (client) ID aus der Entra App Registration |
| `PORT` | optional | Standard: `8000` |

---

## Laufender Betrieb

**Neues Image deployen:**
Code auf `main` pushen → GitHub Actions abwarten → in Azure unter `Revisions and replicas → + Create new revision` eine neue Revision erstellen. Name/Suffix eintragen — auch wenn es optional wirkt, der Button bleibt sonst ausgegraut.

**Client Secret erneuern (alle 24 Monate):**
Entra → App Registration → neues Secret → in Copilot Studio unter der MCP Tool-Konfiguration aktualisieren → altes Secret löschen.

**Classic PAT erneuern:**
GitHub → Classic PAT mit `read:packages` neu generieren → in der Container App unter `Containers → Edit and deploy` aktualisieren.

---

## The Raccoon Way

Dieser Server ist entstanden im Rahmen von [The Raccoon Way](https://the-raccoon-way.de) — Inhalte rund um Microsoft Power Platform, Copilot Studio und Azure für alle, die Dinge einfach zum Laufen bringen wollen. From No-Code to No-Limits.

Fragen, Feedback, eigene Erfahrungen beim Nachbauen: [LinkedIn](https://www.linkedin.com/in/danielazurwerra/) — ich beisse auch nicht.

---

## Lizenz

MIT
