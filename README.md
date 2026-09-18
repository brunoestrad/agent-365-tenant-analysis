# Agent Governance for Microsoft Entra

A two-part governance toolkit for AI agent identities in Microsoft Entra:

1. **`a365-tenant-analysis.ps1`** — PowerShell script that queries Microsoft Graph and exports a structured `.xlsx` report
2. **`governance-dashboard.html`** — Self-contained browser dashboard that ingests the exported report and renders it as an interactive UI

> "Identity is the silent author of every choice" - Bruno Estrada

---

## Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Prerequisites](#prerequisites)
- [Quick Start](#quick-start)
- [Script Reference](#script-reference)
  - [Configuration](#configuration)
  - [How It Works](#how-it-works)
  - [Output — Excel Workbook](#output--excel-workbook)
- [Dashboard Reference](#dashboard-reference)
  - [Loading a Report](#loading-a-report)
  - [Pages](#pages)
- [RAG Risk Rating Logic](#rag-risk-rating-logic)
- [Scope and Limitations](#scope-and-limitations)
- [Troubleshooting](#troubleshooting)

---

## Overview

As organisations deploy AI agents via Copilot Studio, Azure AI Foundry, and custom implementations,
the resulting identities accumulate in Entra without consistent governance.
This toolkit provides a **read-only, point-in-time snapshot** of:

- Every AI agent identity in the tenant (modern and classic)
- Blueprint health (credentials, identifier URIs, OAuth scopes)
- Ownership and accountability (owners, sponsors)
- Credential risk (secrets, expiry, certificate vs. secret)
- RAG (Red / Amber / Green) risk ratings per agent

No data is written back to the tenant. The script is entirely read-only.

---

## Architecture

### Two Types of Agent Identity

| Type | Description | Governance Standard |
|---|---|---|
| **Modern** | Backed by Microsoft Entra Agent ID (`agentIdentity` service principal) | ✅ Current |
| **Classic** | Plain service principal tagged `AgenticInstance` — pre-Agent ID | ⚠️ Migrate |

### Two Blueprint Endpoints

The script queries two distinct Graph endpoints to get a complete picture of blueprints:

```
/applications/microsoft.graph.agentIdentityBlueprint
    → agentIdentityBlueprint objects (inherits from 'application')
    → Carries: keyCredentials, passwordCredentials, identifierUris, api, signInAudience
    → Only blueprints YOUR tenant owns (Company-published)

/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal
    → agentIdentityBlueprintPrincipal objects (inherits from 'servicePrincipal')
    → Carries: id, appId, displayName — NOT credentials or API settings
    → ALL blueprint principals visible to your tenant (Company + external)
```

This split is intentional. Credential and API configuration live on the **application object**
in the publisher's tenant. For external blueprints (e.g. Microsoft-published Copilot Studio
shared blueprint), those objects are not visible — the script communicates this clearly
in the report via the `BlueprintIsExternal` flag rather than treating absent data as a finding.

---

## Prerequisites

### PowerShell Modules

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

| Module | Version | Purpose |
|---|---|---|
| `Microsoft.Graph.Authentication` | 2.x+ | `Connect-MgGraph`, `Invoke-MgGraphRequest` |
| `ImportExcel` | 7.x+ | `Export-Excel` — no Excel installation required |

### Microsoft Entra Role

| Role | Reason |
|---|---|
| **Global Reader** | Read-only access to all objects queried |

> A lower-privileged custom role scoped to the relevant resource types may be substituted
> if your organisation's policy restricts Global Reader assignment.

### Microsoft Graph Scopes

The script requests the following delegated scopes at runtime:

| Scope | Purpose |
|---|---|
| `AgentIdentity.Read.All` | List agent identities, blueprints, owners, sponsors |
| `Application.Read.All` | List blueprint application objects, `appRoleAssignments` |
| `User.Read.All` | Resolve display names and UPNs for owners and sponsors |

The session uses `-ContextScope Process` — the token is isolated to the current
PowerShell process and is not written to the persistent MSAL cache.

---

## Quick Start

```powershell
# 1. Clone or download the repository
# 2. Open the script and set your Tenant ID
$TenantId = "<YOUR_TENANT_ID>"   # Line ~36

# 3. Run
.\a365-tenant-analysis.ps1

# 4. Sign in when the browser prompt appears — use an account with Global Reader
# 5. The .xlsx report is saved to the same directory as the script
# 6. Open governance-dashboard.html in a browser and drop the .xlsx onto it
```

---

## Script Reference

### Configuration

Open `a365-tenant-analysis.ps1` and edit the `CONFIGURATION` block near the top:

```powershell
$TenantId           = "<YOUR_TENANT_ID>"   # Your Entra tenant ID (GUID)
$CredentialWarnDays = 60                   # Flag credentials expiring within N days
$InactiveDays       = 90                   # Reserved — sign-in activity not yet implemented
$ReportOutputPath   = "$PSScriptRoot\AgentGovernanceReport_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"
```

| Parameter | Default | Description |
|---|---|---|
| `$TenantId` | `<YOUR_TENANT_ID>` | **Required.** Replace with your tenant GUID before running. |
| `$CredentialWarnDays` | `60` | Credentials expiring within this window are flagged AMBER. |
| `$InactiveDays` | `90` | Reserved for future sign-in activity enrichment. Has no effect currently. |
| `$ReportOutputPath` | Script directory | Datestamped `.xlsx` filename. Change path if needed. |

### How It Works

The script runs in six sequential steps:

```
Step 1 — Fetch modern agent identities   (agentIdentity service principals)
Step 2 — Detect classic SP-backed agents (tagged AgenticInstance, deduplicated)
Step 3 — Fetch blueprints                (principals + application objects + permissions)
Step 4 — Enrich each modern agent        (owners, sponsors, blueprint health, RAG rating)
Step 5 — Build classic agent summary     (platform classification, static RED rating)
Step 6 — Build blueprint health report   (credential type, permission analysis)
```

**Step 3** includes a permission enrichment loop that:
- Fetches `appRoleAssignments` per blueprint principal
- Resolves role display names via a cached service principal lookup (one fetch per resource, reused across all blueprints and agents)
- Flags permissions that are granted in the tenant but absent from the blueprint's manifest as `⚠️ UNDECLARED`

**Step 4** uses a four-path detection strategy to identify Copilot Studio agents reliably,
because the platform tag is sometimes absent on newly created agents:

```
Path 1 → Tag:  AgentCreatedBy:CopilotStudio on the agent SP
Path 2 → Blueprint principal display name contains "Copilot Studio"
Path 3 → Blueprint application display name contains "Copilot Studio"
Path 4 → agentIdentityBlueprintId matches the known shared blueprint appId
```

### Output — Excel Workbook

The script writes a single `.xlsx` workbook with three sheets:

#### Sheet 1 — Modern Agents

One row per `agentIdentity` service principal.

| Column | Description |
|---|---|
| `AgentName` | Display name |
| `AgentId` | Object ID of the agent SP |
| `AgentAppId` | Application (client) ID |
| `IsEnabled` | Whether the SP is enabled |
| `CreatedDate` | ISO date of creation |
| `DaysSinceCreation` | Days since the agent was registered |
| `SignInAudience` | Token audience (AzureADMyOrg, etc.) |
| `IsClassic` / `IsModern` | Classification flags |
| `HasBlueprint` | Whether a blueprint is associated |
| `BlueprintId` | appId of the associated blueprint |
| `BlueprintName` | Display name of the blueprint |
| `BlueprintIsExternal` | `true` = publisher is external (Microsoft/3rd-party) |
| `BlueprintMultiTenant` | `true` = blueprint accepts tokens from other tenants |
| `BlueprintMissingUri` | Blueprint has no identifier URI |
| `BlueprintMissingScope` | Blueprint exposes no OAuth 2.0 scopes |
| `BlueprintNoCredential` | Blueprint has no credential configured |
| `HasClientSecret` | `true` = secret present (high risk) |
| `HasCertificate` | `true` = certificate present (preferred) |
| `CredentialExpired` | Any credential has passed its expiry date |
| `CredentialExpiringSoon` | Any credential expires within `$CredentialWarnDays` |
| `NearestCredentialExpiry` | ISO date of the soonest-expiring credential |
| `NoOwner` / `OwnerCount` / `Owners` / `OwnerUPNs` | Owner accountability fields |
| `HasDisabledOwner` | At least one owner account is disabled |
| `NoSponsor` / `SponsorCount` / `Sponsors` / `SponsorUPNs` | Sponsor accountability fields |
| `HasDisabledSponsor` | At least one sponsor account is disabled |
| `OwnershipPattern` | `Both` / `Owner Only` / `Sponsor Only` / `Neither` |
| `NeverSignedIn` | Placeholder — requires `Reports.Read.All` |
| `IsInactive` | Placeholder — requires sign-in log enrichment |
| `AgentPermissionCount` | Number of `appRoleAssignments` directly on the agent |
| `AgentPermissions` | Pipe-delimited list of directly granted permissions |
| `IsCopilotStudioAgent` | Identified as Copilot Studio via four-path detection |
| `RiskRating` | `RED` / `AMBER` / `GREEN` |

#### Sheet 2 — Blueprints

One row per `agentIdentityBlueprintPrincipal`.

| Column | Description |
|---|---|
| `BlueprintName` | Display name |
| `BlueprintAppId` | Application (client) ID |
| `BlueprintObjectId` | Object ID of the blueprint SP |
| `CreatedDate` | ISO creation date |
| `IsExternalBlueprint` | `true` = no matching application object in this tenant |
| `PublisherDomain` | Verified publisher domain (Company-owned only) |
| `IsMultiTenant` | Accepts tokens from external tenants |
| `ChildAgentCount` | Number of modern agents linked to this blueprint |
| `HasIdentifierUri` | Blueprint has a custom `api://` URI registered |
| `IdentifierUri` | The registered URI(s) |
| `HasOAuthScope` | Blueprint exposes at least one delegated scope |
| `HasClientSecret` | Secret present (`null` = external, data unavailable) |
| `HasCertificate` | Certificate present (`null` = external, data unavailable) |
| `CredentialType` | Human-readable credential summary with risk label |
| `RequiredPermissions` | Count of resource access entries in the manifest |
| `AgentNames` | Semicolon-delimited list of child agent display names |
| `GrantedPermissionCount` | Number of `appRoleAssignments` granted to the blueprint SP |
| `HasHighRiskPermissions` | Any granted permission matches a high-risk keyword |
| `GrantedPermissions` | Pipe-delimited permission list; `⚠️ UNDECLARED` suffix where applicable |

#### Sheet 3 — Classic Agents

One row per classic SP-backed agent (deduplicated against modern agents).

| Column | Description |
|---|---|
| `AgentName` | Display name |
| `AgentId` / `AgentAppId` | Object and application IDs |
| `IsEnabled` | SP enabled state |
| `CreatedDate` / `DaysSinceCreation` | Age data |
| `Platform` | `Copilot Studio` / `Foundry` / `Unknown` — derived from tags |
| `IsClassic` / `IsModern` | Always `true` / `false` |
| `HasBlueprint` | Always `false` — no blueprint association |
| `OwnershipPattern` | Static — verify via Enterprise Apps blade |
| `RiskRating` | Always `RED` |
| `Notes` | Remediation guidance |

---

## Dashboard Reference

`governance-dashboard.html` is a fully self-contained single-page application.
It requires no server, no build step, and no installation.
The only external dependency is the SheetJS (`xlsx`) library loaded from cdnjs,
used to parse the uploaded `.xlsx` file in-browser.

> All processing happens locally in the browser. No data is uploaded or transmitted.

### Loading a Report

1. Open `governance-dashboard.html` in any modern browser (Chrome, Edge, Firefox, Safari)
2. Drop the exported `.xlsx` file onto the upload area, or click to select it
3. The dashboard renders immediately — no page reload required
4. Use **↻ Load New Report** in the sidebar to swap files

### Pages

| Page | Description |
|---|---|
| **How to Read** | Field-by-field reference guide and RAG logic explanation |
| **Summary** | KPI cards, risk distribution, coverage breakdown, ownership and platform charts |
| **All Agents** | Combined modern + classic list with search and sort |
| **Modern Agents** | Filterable by `RED` / `AMBER` / `GREEN` with full detail expansion per agent |
| **Blueprints** | Blueprint health cards with permission detail and credential risk |
| **Classic Agents** | Pre-Agent ID agents with migration guidance |

Each agent card is expandable — click the row to reveal the full field set
including ownership, credential detail, and granted permissions with risk highlighting.

---

## RAG Risk Rating Logic

Ratings are evaluated in priority order. The first matching condition wins.

### 🔴 RED — Immediate action required

| Condition | Reason |
|---|---|
| Agent is classic (SP-backed) | No Agent ID governance, no blueprint, no sponsor enforcement |
| `HasClientSecret = true` | Client secrets on blueprints are high risk — rotate to certificate |
| `CredentialExpired = true` | Agent cannot authenticate — broken |
| `NoSponsor = true` | Business accountability is mandatory |
| `NeverSignedIn = true` and `DaysSinceCreation > 30` | Provisioned but never used — stale identity |

### 🟡 AMBER — Review required

| Condition | Notes |
|---|---|
| `NoOwner = true` (non-Copilot Studio) | Copilot Studio agents set sponsor only by platform design |
| `CredentialExpiringSoon = true` | Within the `$CredentialWarnDays` window |
| `IsInactive = true` | Placeholder — not yet populated |
| `BlueprintMissingUri` (non-Copilot Studio) | Copilot Studio uses Power Platform connectors, not blueprint OBO |
| `BlueprintMissingScope` (non-Copilot Studio) | Same suppression logic as URI |
| `BlueprintNoCredential` (Company-owned only) | External blueprints have no visible application object — not flagged |

### 🟢 GREEN — Fully governed

None of the above conditions apply.

---

## Scope and Limitations

### ✅ Covered

- Modern Entra Agent ID identities (`agentIdentity` service principals)
- Classic SP-backed agents tagged `AgenticInstance`
- Blueprint principals visible to the tenant (Company-owned and external)
- Credential health for Company-owned blueprints
- Ownership (owners and sponsors) for modern agents
- Granted permission analysis with undeclared permission detection

### ❌ Not Covered

| Gap | Reason |
|---|---|
| Sign-in activity (`NeverSignedIn`, `IsInactive`) | Requires `Reports.Read.All` and the beta sign-in logs endpoint |
| External blueprint credential health | Application objects live in the publisher's tenant — inaccessible |
| Agent Builder agents | Different identity model — not surfaced by these endpoints |
| Third-party agents with non-standard tags | Empirical tag taxonomy; non-conforming agents will not appear |
| Shadow agents | Requires network traffic inspection outside of Graph |

---

## Troubleshooting

### `Export-Excel: The term 'Export-Excel' is not recognized`
The `ImportExcel` module is not installed. Run:
```powershell
Install-Module ImportExcel -Scope CurrentUser
```

### `Invoke-MgGraphRequest: Insufficient privileges`
Ensure the signed-in account has **Global Reader** (or equivalent custom role)
and that all three Graph scopes were consented at the interactive sign-in prompt.
If consent was not granted, disconnect and re-run — the browser prompt will
include a consent screen on first use.

### `The workbook is empty / No recognised data found`
The dashboard expects sheets named exactly **Modern Agents**, **Blueprints**,
and **Classic Agents**. Do not rename sheets in the exported file.

### Dashboard shows no data for Blueprints or Classic Agents
If no classic agents exist, the **Classic Agents** sheet is omitted from the workbook
by design. The dashboard handles this gracefully. Blueprints will always be present
if blueprint principals are visible to the tenant.

### `Could not identify Copilot Studio shared blueprint`
This warning is printed if no blueprint principal with "Copilot Studio" in its display
name is found. Copilot Studio Path 4 detection (by known appId) will not function,
but Paths 1–3 remain active. This does not cause incorrect RAG ratings — it degrades
detection coverage for brand-new agents that also lack the platform tag.

---
## License
MIT License