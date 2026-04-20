# Oskar MCP Server – Setup Guide

Vollständige Schritt-für-Schritt Anleitung zur Einrichtung des Oskar MCP Servers.  
Stand: April 2026 | Version: 1.0 (POC)

> **Wichtige Reihenfolge:** Erst Entra App Registration, dann GitHub/Code fertigstellen und Image bauen, DANN erst Azure Container App anlegen. So startet der Container beim ersten Deployment direkt sauber – ohne Image-Fehler und ohne leere Container-Slots.

---

## Inhaltsverzeichnis

1. [Voraussetzungen](#1-voraussetzungen)
2. [Phase 1 – Entra App Registration (DEV-Tenant)](#2-phase-1--entra-app-registration-dev-tenant)
3. [Phase 2 – Admin Consent erteilen](#3-phase-2--admin-consent-erteilen)
4. [Phase 3 – Python MCP Server Code](#4-phase-3--python-mcp-server-code)
5. [Phase 4 – GitHub Actions CI/CD einrichten](#5-phase-4--github-actions-cicd-einrichten)
6. [Phase 5 – Erstes Docker Image bauen](#6-phase-5--erstes-docker-image-bauen)
7. [Phase 6 – Azure Infrastruktur anlegen](#7-phase-6--azure-infrastruktur-anlegen)
8. [Phase 7 – Managed Identity & Federated Credential](#8-phase-7--managed-identity--federated-credential)
9. [Phase 8 – Copilot Studio Connector](#9-phase-8--copilot-studio-connector)
10. [Oskar Instruktionen](#10-oskar-instruktionen)
11. [Laufender Betrieb](#11-laufender-betrieb)
12. [Offene TODOs](#12-offene-todos)

---

## 1. Voraussetzungen

| Was | Wo | Wer |
|---|---|---|
| Global Admin im DEV-Tenant (r6ly) | entra.microsoft.com | Daniela |
| Azure Subscription (DZurwerra = in2success) | portal.azure.com | Michael |
| GitHub Repo (privat) angelegt | github.com/d-zurwerra/mcp-m365-graph-server | Daniela |
| GitHub PAT mit `read:packages` + `write:packages` + `repo` | github.com → Settings → Developer Settings | Daniela |
| Copilot Studio Lizenz | copilotstudio.microsoft.com | in2success |
| Resource Group `rg-oskar-mcp-env` | portal.azure.com | Michael/Daniela |

---

## 2. Phase 1 – Entra App Registration (DEV-Tenant)

**Wo:** `https://entra.microsoft.com` → DEV-Tenant (r6ly) → Identity → Applications → App registrations

### 2.1 App Registration anlegen

1. **+ New registration**
2. Werte:
   - **Name:** `oskar-mcp-graph-connector`
   - **Supported account types:** `Accounts in any organizational directory (Multitenant)`
3. **Register**
4. Notieren:
   - `Application (client) ID` → wird später `APP_CLIENT_ID`
   - `Directory (tenant) ID` → wird später `DEV_TENANT_ID`

### 2.2 Client Secret anlegen (temporär für POC)

1. **Certificates & secrets → Client secrets → + New client secret**
2. Description: `poc-temp`, Expires: `90 days`
3. **Add** → Value **sofort kopieren** → wird `CLIENT_SECRET`

> ⚠️ Secret läuft in 90 Tagen ab! Vor Ablauf auf Workload Identity Federation umstellen (siehe TODOs).

### 2.3 API Permissions setzen

1. **API permissions → + Add a permission → Microsoft Graph → Application permissions**
2. Folgende Permissions hinzufügen:

| Permission | Zweck |
|---|---|
| `Channel.Create` | Teams Channels erstellen |
| `Chat.ReadWrite.All` | Teams Chat |
| `Group.ReadWrite.All` | M365 Gruppen verwalten |
| `Notes.ReadWrite.All` | OneNote (⚠️ App-only nicht supported) |
| `Sites.Manage.All` | SharePoint Listen/Struktur anlegen |
| `Sites.ReadWrite.All` | SharePoint Items lesen/schreiben |
| `Tasks.ReadWrite.All` | Planner |
| `Team.Create` | Teams erstellen |
| `TeamMember.ReadWrite.All` | Team-Mitglieder verwalten |
| `TeamSettings.ReadWrite.All` | Team-Einstellungen |
| `User.Read.All` | User suchen |

---

## 3. Phase 2 – Admin Consent erteilen

**Wo:** DEV-Tenant → App Registration `oskar-mcp-graph-connector` → **API permissions**

1. **Grant admin consent for [DEV-Tenant]** klicken
2. **Yes** bestätigen
3. Alle Permissions zeigen ✅ **Granted for [DEV-Tenant]**

> Prüfen unter **Enterprise applications → oskar-mcp-graph-connector** ob Service Principal angelegt wurde.

---

## 4. Phase 3 – Python MCP Server Code

### 4.1 Repository-Struktur

Alle Dateien über die GitHub Web UI anlegen (Pfade mit `/` direkt im Dateinamen eingeben – GitHub erstellt Ordner automatisch):

```
mcp-m365-graph-server/
├── app/
│   ├── __init__.py
│   ├── main.py              # MCP Server Entry Point
│   ├── auth.py              # Graph API Auth
│   └── tools/
│       ├── __init__.py
│       ├── groups.py        # M365 Gruppen (Kern-Onboarding)
│       ├── teams.py         # Teams Tools
│       ├── planner.py       # Planner Tools
│       ├── sharepoint.py    # SharePoint Tools
│       └── onenote.py       # OneNote Tools (eingeschränkt)
├── Dockerfile
├── requirements.txt
├── .env.example
├── .gitignore
└── .github/
    └── workflows/
        └── deploy.yml
```

### 4.2 Wichtige Code-Details

**main.py – DNS Rebinding Protection deaktivieren (notwendig für ACA Proxy):**
```python
mcp = FastMCP(
    name="Oskar M365 MCP Server",
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=False
    ),
)
# WICHTIG: app muss auf Modul-Ebene stehen, nicht in if __name__ == "__main__"
app = mcp.streamable_http_app()
```

**Dockerfile CMD:**
```dockerfile
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--forwarded-allow-ips", "*"]
```

**auth.py – Client Secret (temporär):**
```python
credential = ClientSecretCredential(
    tenant_id=DEV_TENANT_ID,
    client_id=APP_CLIENT_ID,
    client_secret=CLIENT_SECRET,
)
```

---

## 5. Phase 4 – GitHub Actions CI/CD einrichten

### 5.1 App Registration für GitHub Actions (DEV-Tenant)

**Wo:** `https://entra.microsoft.com` → DEV-Tenant

> ⚠️ Für den POC im DEV-Tenant, da dort Global Admin Rechte vorhanden sind. TODO: Nach in2success umziehen.

1. **App registrations → + New registration**
2. Werte:
   - **Name:** `oskar-github-actions-deploy`
   - **Supported account types:** Single tenant
3. **Register** → Notieren: `CLIENT_ID`, `TENANT_ID`

### 5.2 Federated Credential für GitHub Branch main

1. **Certificates & secrets → Federated credentials → + Add credential**
2. **Scenario:** `GitHub Actions deploying Azure resources`
3. Werte:
   - **Organization:** `d-zurwerra`
   - **Repository:** `mcp-m365-graph-server`
   - **Entity type:** Verzweigung (Branch)
   - **Branch:** `main`
   - **Name:** `github-actions-main`
4. **Add**

### 5.3 GitHub Secrets setzen

**Wo:** `github.com/d-zurwerra/mcp-m365-graph-server → Settings → Secrets and variables → Actions`

| Secret | Wert |
|---|---|
| `AZURE_CLIENT_ID` | Client ID von `oskar-github-actions-deploy` |
| `AZURE_TENANT_ID` | Tenant ID (DEV-Tenant, später in2success) |
| `AZURE_SUBSCRIPTION_ID` | Subscription ID von DZurwerra |

### 5.4 RBAC auf Resource Group (Michael)

> ⚠️ Erfordert Owner-Rechte auf der Subscription – muss Michael durchführen. Bis dahin bleibt der Deploy-Job in `deploy.yml` auskommentiert.

**Wo:** Azure Portal → Resource Group `rg-oskar-mcp-env` → **Access control (IAM)**

1. **+ Add → Add role assignment**
2. Role: `Contributor`
3. Members: `oskar-github-actions-deploy`
4. **Review + assign**

---

## 6. Phase 5 – Erstes Docker Image bauen

> ✅ Dieser Schritt muss abgeschlossen sein bevor Azure angelegt wird!

1. Code auf `main` pushen → GitHub Actions startet automatisch
2. Build prüfen: `github.com/d-zurwerra/mcp-m365-graph-server/actions`
3. Build muss ✅ grün sein
4. Image prüfen: `github.com/d-zurwerra?tab=packages` → `mcp-m365-graph-server:latest` muss sichtbar sein

**Erst wenn das Image in `ghcr.io` vorhanden ist → weiter mit Phase 6!**

---

## 7. Phase 6 – Azure Infrastruktur anlegen

**Wo:** `https://portal.azure.com` → DZurwerra Subscription

### 7.1 User-Assigned Managed Identity anlegen

1. Suche: **Managed Identities → + Create**
2. Werte:
   - **Subscription:** DZurwerra
   - **Resource Group:** `rg-oskar-mcp-env`
   - **Name:** `mi-oskar-mcp-dev`
   - **Region:** West Europe
3. **Review + Create → Create**
4. Notieren: **Object (principal) ID** → `MI_PRINCIPAL_ID`

### 7.2 Container Apps Environment anlegen

1. Suche: **Container Apps Environments → + Create**
2. Werte:
   - **Subscription:** DZurwerra
   - **Resource Group:** `rg-oskar-mcp-env`
   - **Name:** `cae-oskar-mcp-dev`
   - **Region:** West Europe
   - **Workload Profiles:** Consumption only
   - **Networking:** Public, kein Custom VNet
3. **Review + Create → Create**

### 7.3 Container App anlegen

1. Suche: **Container Apps → + Create**
2. **Basics Tab:**
   - **Subscription:** DZurwerra
   - **Resource Group:** `rg-oskar-mcp-env`
   - **Name:** `cap-oskar-mcp-dev`
   - **Region:** West Europe
   - **Container Apps Environment:** `cae-oskar-mcp-dev`
3. **Container Tab:**
   - Image source: `Docker Hub or other registries`
   - Image type: `Private`
   - Registry login server: `ghcr.io`
   - Username: `d-zurwerra`
   - Password: GitHub PAT
   - Image and tag: `d-zurwerra/mcp-m365-graph-server:latest`
     > ⚠️ `ghcr.io` NUR im Registry-Feld – NICHT nochmal im Image-Pfad wiederholen!
   - CPU: `0.5`, Memory: `1Gi`
   - Environment Variables:

     | Name | Value |
     |---|---|
     | `DEV_TENANT_ID` | DEV-Tenant ID (r6ly) |
     | `APP_CLIENT_ID` | Client ID von `oskar-mcp-graph-connector` |
     | `CLIENT_SECRET` | Secret aus Phase 1 |

4. **Ingress Tab:**
   - Ingress: ✅ Enabled
   - Traffic: `Accepting traffic from anywhere`
   - Type: `HTTP`, Transport: `Auto`
   - Insecure connections: ❌ Nicht erlaubt
   - Target port: `8000`
5. **Review + Create → Create**

> ⚠️ Bekannter Bug: Bei "Create new revision" ist der Create-Button ausgegraut bis das **Name/Suffix** Feld ausgefüllt ist – auch wenn es optional wirkt.

### 7.4 Managed Identity zuweisen

1. `cap-oskar-mcp-dev` → **Settings → Identity → User assigned → + Add**
2. `mi-oskar-mcp-dev` auswählen → **Add**

### 7.5 Deployment testen

Öffne im Browser:
```
https://cap-oskar-mcp-dev.mangoisland-de1ecd4a.westeurope.azurecontainerapps.io/mcp
```

Erwartete Antwort (korrekt – Server läuft):
```json
{"jsonrpc":"2.0","id":"server-error","error":{"code":-32600,"message":"Not Acceptable: Client must accept text/event-stream"}}
```

---

## 8. Phase 7 – Managed Identity & Federated Credential

**Wo:** `https://entra.microsoft.com` → DEV-Tenant → App Registration `oskar-mcp-graph-connector`

1. **Certificates & secrets → Federated credentials → + Add credential**
2. **Scenario:** `Other issuer`
3. Werte:
   - **Issuer:** `https://login.microsoftonline.com/55d7a86a-b2ce-413d-b267-92be62ff9c9e/v2.0`
     *(= in2success Tenant ID – **NICHT** der DEV-Tenant!)*
   - **Subject identifier:** `MI_PRINCIPAL_ID` (Object ID der `mi-oskar-mcp-dev`)
   - **Name:** `oskar-mcp-aca-identity`
   - **Audience:** `api://AzureADTokenExchange`
4. **Add**

> ⚠️ Aktuell schlägt WIF mit Fehler AADSTS700236 fehl weil die App Registration im DEV-Tenant liegt, die MI aber im in2success Tenant. Lösung: App Registration nach in2success umziehen (Michael). Bis dahin läuft alles über Client Secret.

---

## 9. Phase 8 – Copilot Studio Connector

**Wo:** `https://copilotstudio.microsoft.com` → Oskar 4 Onboarding → **Tools → + Add a tool**

1. **Model Context Protocol** auswählen
2. Werte:
   - **Server name:** `M365 Tools R6LY`
   - **Server description:** `Microsoft 365 Tools für Teams, SharePoint und Planner`
   - **Server URL:** `https://cap-oskar-mcp-dev.mangoisland-de1ecd4a.westeurope.azurecontainerapps.io/mcp`
   - **Authentication:** `None`
3. **Create**
4. Alle Tools aktivieren (Master-Toggle oben)
5. **Save → Publish**

> ℹ️ Copilot Studio MCP ist ein Trusted Runtime – kein OAuth nötig und aktuell nicht möglich.

---

## 10. Oskar Instruktionen

Vollständige Instruktionen für **Oskar 4 Onboarding** – direkt einfügen unter **Settings → Instructions**:

```
Du bist Oskar, der interne Onboarding-Assistent von in2success.
Du hilfst dabei, neue Kunden in der Microsoft 365 Umgebung einzurichten.
Antworte immer auf Deutsch.

Der Tenant-Admin muss immer als Owner und Member hinzugefügt werden:
- Admin User ID: 8f7e5b0e-9c50-4e87-93c0-8a526bb1ab4e

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHASE 1: STRUKTUR ANLEGEN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SCHRITT 1: Informationen sammeln
Frage den Nutzer BEVOR du irgendetwas anlegst:
- "Wie heißt der neue Kunde?"
- "Wer soll Owner des Teams werden? (Name oder Email-Adresse)"

SCHRITT 2: Owner suchen
→ find_user (search = angegebener Owner-Name/Email)
→ Zeige den gefundenen User und bitte um Bestätigung
→ Merke dir die User ID als OWNER_ID
→ Nach Bestätigung: Verwende OWNER_ID direkt in allen weiteren Schritten
→ Rufe find_user NIEMALS nochmal auf
→ Falls du OWNER_ID nicht mehr hast: frage den Nutzer erneut nach
  der Email-Adresse und rufe find_user einmalig auf

SCHRITT 3: Gruppe prüfen
→ get_m365_groups (search = Kundenname)
→ Falls NICHT vorhanden: weiter mit Schritt 4
→ Falls vorhanden: GROUP_ID merken, weiter mit Schritt 5

SCHRITT 4: M365 Gruppe anlegen
→ Verwende OWNER_ID aus Schritt 2 direkt – suche den User NICHT nochmal
→ create_m365_group (
    display_name = Kundenname,
    visibility = "Private",
    owner_user_ids = [OWNER_ID, "8f7e5b0e-9c50-4e87-93c0-8a526bb1ab4e"]
  )
→ Merke dir die group_id als GROUP_ID
→ Berichte: "✅ Gruppe angelegt: [GROUP_ID]"

SCHRITT 5: Owners prüfen
→ get_group_owners (GROUP_ID)
→ Falls OWNER_ID fehlt: add_group_owner (GROUP_ID, OWNER_ID)
→ Falls Admin fehlt: add_group_owner (GROUP_ID, "8f7e5b0e-9c50-4e87-93c0-8a526bb1ab4e")
→ Berichte: "✅ Owners gesetzt"

SCHRITT 6: Team upgraden
→ upgrade_to_team (GROUP_ID)
→ Berichte: "✅ Team erstellt"

SCHRITT 7: Members hinzufügen
→ Falls GROUP_ID nicht mehr bekannt: get_m365_groups (search = Kundenname)
→ add_group_member (GROUP_ID, OWNER_ID)
→ Berichte Ergebnis des Tool-Aufrufs
→ add_group_member (GROUP_ID, "8f7e5b0e-9c50-4e87-93c0-8a526bb1ab4e")
→ Berichte Ergebnis des Tool-Aufrufs
→ Berichte: "✅ Members hinzugefügt"

Nach Schritt 7 – STOPPE und gib aus:
"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ PHASE 1 ABGESCHLOSSEN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ Gruppe/Team: [Kundenname] (Private)
✅ GROUP_ID: [GROUP_ID]
✅ Owners: [Namen]
✅ Members: hinzugefügt
Antworte mit 'weiter' um Phase 2 zu starten."

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHASE 2: INHALTE ANLEGEN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Starte Phase 2 NUR wenn der Nutzer "weiter" sagt.

SCHRITT 8: SharePoint Site holen
→ get_group_site (GROUP_ID)
→ Merke SITE_ID
→ Berichte: "✅ SharePoint Site: [webUrl]"

SCHRITT 9: Onboarding Checkliste
→ get_sharepoint_lists (SITE_ID)
→ Falls "Onboarding Checkliste" NICHT vorhanden:
  create_sharepoint_list (
    SITE_ID,
    display_name = "Onboarding Checkliste",
    columns = [
      {"name": "Status", "choice": {"choices": ["Offen", "In Bearbeitung", "Erledigt"]}},
      {"name": "Verantwortlich", "text": {}}
    ]
  )
→ Merke LIST_ID
→ Berichte: "✅ Liste angelegt"

SCHRITT 10: Checklisten-Items
→ get_sharepoint_list_items (SITE_ID, LIST_ID)
→ Nur fehlende Items anlegen:
  → create_sharepoint_list_item: {"Title": "Kick-off Termin vereinbaren", "Status": "Offen"}
  → create_sharepoint_list_item: {"Title": "Zugänge einrichten", "Status": "Offen"}
  → create_sharepoint_list_item: {"Title": "Erstdokumentation erstellen", "Status": "Offen"}
  → create_sharepoint_list_item: {"Title": "Abnahme & Übergabe", "Status": "Offen"}
→ Berichte: "✅ Items angelegt"

SCHRITT 11: Planner Plan
→ get_planner_plans (GROUP_ID)
→ Falls nicht vorhanden:
  create_planner_plan (GROUP_ID, title = "Onboarding " + Kundenname)
→ Merke PLAN_ID
→ Berichte: "✅ Plan angelegt"

SCHRITT 12: Onboarding Bucket erstellen
→ create_planner_bucket (PLAN_ID, name = "Onboarding")
→ Merke BUCKET_ID
→ Berichte: "✅ Bucket angelegt"

SCHRITT 13: Planner Tasks
→ get_planner_tasks (PLAN_ID)
→ Nur fehlende Tasks anlegen – alle mit bucket_id = BUCKET_ID:
  → create_planner_task: title = "Kick-off Termin vereinbaren", bucket_id = BUCKET_ID
  → create_planner_task: title = "Zugänge einrichten", bucket_id = BUCKET_ID
  → create_planner_task: title = "Erstdokumentation erstellen", bucket_id = BUCKET_ID
  → create_planner_task: title = "Abnahme & Übergabe", bucket_id = BUCKET_ID
→ Berichte: "✅ Tasks angelegt"

SCHRITT 14: Abschlusszusammenfassung
✅ Gruppe/Team: [Kundenname] (Private)
✅ Owners & Members: [Namen]
✅ SharePoint Site: [webUrl]
✅ Onboarding Checkliste: angelegt
✅ Planner Plan: angelegt mit Bucket "Onboarding" und 4 Tasks
💡 Die Owners können weitere Mitglieder direkt in Teams einladen

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WICHTIGE REGELN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- Überspringe KEINEN Schritt
- Zeige nach JEDEM Tool-Aufruf das Ergebnis
- Verwende NIEMALS Platzhalter wie [group_id] – hole fehlende IDs immer frisch
- Rufe find_user NIEMALS mehrfach auf
- Lege NIEMALS doppelte Strukturen an
- Antworte immer auf Deutsch
```

---

## 11. Laufender Betrieb

### Neues Image deployen (manuell solange Deploy-Job auskommentiert)

1. Code ändern und auf `main` pushen
2. GitHub Actions Build abwarten
3. Azure Portal → `cap-oskar-mcp-dev` → **Revisions and replicas → + Create new revision**
4. **Name/Suffix** eingeben (z.B. `v20`)
5. **Create**

### Client Secret rotieren

1. `entra.microsoft.com` → DEV-Tenant → `oskar-mcp-graph-connector` → **Certificates & secrets**
2. **+ New client secret** → Value kopieren
3. ACA → neue Revision mit neuem `CLIENT_SECRET` Environment Variable
4. Altes Secret löschen

### Logs einsehen

| Log-Typ | Wo |
|---|---|
| Real-time | Azure Portal → `cap-oskar-mcp-dev` → Revision → Logs → Real-time |
| Application (historisch) | Azure Portal → `cap-oskar-mcp-dev` → Revision → Logs → Historical → Application |
| System (historisch) | Azure Portal → `cap-oskar-mcp-dev` → Revision → Logs → Historical → System |

### Bekannte Fehler & Lösungen

| Fehler | Ursache | Lösung |
|---|---|---|
| `MANIFEST_UNKNOWN` | `ghcr.io` doppelt im Image-Pfad | Nur im Registry-Feld, nicht im Image-Pfad |
| `Attribute "app" not found` | `app` nicht auf Modul-Ebene | `app = mcp.streamable_http_app()` vor `if __name__` |
| `Invalid Host header` | DNS Rebinding Protection aktiv | `enable_dns_rebinding_protection=False` in FastMCP() |
| Container crasht sofort | Import-Fehler / fehlendes Modul | Logs → Historical → Application prüfen |
| Create-Button ausgegraut | Name/Suffix Feld leer | Name/Suffix bei neuer Revision ausfüllen |

---

## 12. Offene TODOs

### 🔴 Dringend (vor ~Juli 2026)

- [ ] **Client Secret rotieren** – läuft in 90 Tagen ab
- [ ] **App Registration `oskar-mcp-graph-connector` nach in2success** (Michael)
  → Dann Workload Identity Federation statt Client Secret möglich

### 🟡 Kurzfristig

- [ ] **GitHub Actions Deploy-Job aktivieren** – Michael muss `Contributor` auf `rg-oskar-mcp-env` für `oskar-github-actions-deploy` setzen
- [ ] **App Registration `oskar-github-actions-deploy` nach in2success** umziehen
- [ ] **Web-Suche in Copilot Studio deaktivieren**
- [ ] **Node.js 24** in `deploy.yml`: `FORCE_JAVASCRIPT_ACTIONS_TO_NODE24=true`

### 🟢 Mittelfristig

- [ ] Application Insights für Monitoring
- [ ] IP-Restriction für MCP Endpunkt
- [ ] Microsoft Loop Integration (sobald Graph API Beta stabilisiert)

---

## Anhang: Wichtige IDs & URLs

| Bezeichnung | Wert |
|---|---|
| in2success Tenant ID | `55d7a86a-b2ce-413d-b267-92be62ff9c9e` |
| DEV-Tenant ID (r6ly) | `4e2ce74e-a6a4-45f4-8111-3bc60634c8a9` |
| Admin User ID | `8f7e5b0e-9c50-4e87-93c0-8a526bb1ab4e` |
| ACA URL | `cap-oskar-mcp-dev.mangoisland-de1ecd4a.westeurope.azurecontainerapps.io` |
| MCP Endpunkt | `https://cap-oskar-mcp-dev.mangoisland-de1ecd4a.westeurope.azurecontainerapps.io/mcp` |
| GitHub Repo | `github.com/d-zurwerra/mcp-m365-graph-server` |
| Container Image | `ghcr.io/d-zurwerra/mcp-m365-graph-server:latest` |
| Copilot Studio | `copilotstudio.microsoft.com` → Oskar 4 Onboarding |
| Azure Portal | `portal.azure.com` → DZurwerra → rg-oskar-mcp-env |
| Entra DEV-Tenant | `entra.microsoft.com` → r6ly |
| Planner | `tasks.office.com` (mit r6ly Account einloggen) |
