#Requires -Modules Microsoft.Graph.Authentication, ImportExcel
<#
.SYNOPSIS
    Agent Governance Report
    Assesses the governance state of all AI agent identities in the Entra tenant.
.DESCRIPTION
    Queries Microsoft Graph to produce a structured governance report covering:
    - Inventory & classification (modern vs classic)
    - Ownership & accountability (owners, sponsors)
    - Blueprint health (credentials, URI, scope)
    - Credential risk (secrets, expiry)
    - Lifecycle & activity (last sign-in, never used)
    - Summary with RAG risk ratings per agent
.NOTES
    Required Entra Roles:  Global Reader  (read-only report)
    Required Graph Scopes:
        AgentIdentity.Read.All     — list agents, blueprints, owners
        Application.Read.All       — list blueprint apps, appRoleAssignments
        User.Read.All              — resolve user details
    ARCHITECTURE NOTE — Two blueprint endpoints:
        /applications/microsoft.graph.agentIdentityBlueprint
            → Returns agentIdentityBlueprint objects (inherits from 'application')
            → Has: keyCredentials, passwordCredentials, identifierUris, api, signInAudience
            → Only returns blueprints YOUR tenant owns (publisher = Company)
        /servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal
            → Returns agentIdentityBlueprintPrincipal objects (inherits from 'servicePrincipal')
            → Has: id, appId, displayName — but NOT credentials or API settings
            → Returns ALL blueprint principals visible to your tenant (Company + external)
    Run: .\a365-tenant-analysis.ps1
#>


# ─────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────
$TenantId = "<YOUR_TENANT_ID>"
$CredentialWarnDays = 60    # Flag credentials expiring within N days
$InactiveDays = 90    # Flag agents with no sign-in for N days (future use)
$ReportOutputPath = "$PSScriptRoot\AgentGovernanceReport_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"


# ─────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────
function Get-AllPages {
    param([string]$Uri, [hashtable]$Headers = @{})
    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -Headers $Headers
        if ($response.value) {
            foreach ($item in $response.value) { $results.Add($item) }
        }
        $nextUri = $response.'@odata.nextLink'
    } while ($nextUri)
    return , $results.ToArray()   # Comma operator forces array even for single item
}
function Get-SubResource {
    <#
    .SYNOPSIS Fetches a sub-resource (owners/sponsors) for a given SP.
              Returns an empty array gracefully on 404, 403, or missing value.
    #>
    param([string]$Uri)
    try {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        if ($null -eq $r.value) { return , @() } 
        return , @($r.value)
    }
    catch {
        return , @()
    }
}
function Get-RAGStatus {
    <#
    .SYNOPSIS Assigns a Red/Amber/Green risk rating based on governance flags.
    .NOTES    RED  = immediate risk requiring action
              AMBER = governance gap requiring review
              GREEN = fully governed
    #>
    param([PSCustomObject]$Agent)
    # ── RED ───────────────────────────────────────────────────────────────────
    if ($Agent.IsClassic) { return "RED" }   # No Agent ID governance
    # HasClientSecret is $null for external blueprints (data inaccessible)
    # $null evaluates as $false so this safely skips for external agents
    if ($Agent.HasClientSecret) { return "RED" }   # High credential risk
    if ($Agent.CredentialExpired) { return "RED" }   # Broken agent
    if ($Agent.NoSponsor) { return "RED" }   # Mandatory accountability missing
    if ($Agent.NeverSignedIn -and $Agent.DaysSinceCreation -gt 30) { return "RED" }
    # ── AMBER ─────────────────────────────────────────────────────────────────
    # Owner: Copilot Studio only sets sponsor by platform design — not a gap
    if ($Agent.NoOwner -and -not $Agent.IsCopilotStudioAgent) { return "AMBER" }
    if ($Agent.CredentialExpiringSoon) { return "AMBER" }
    if ($Agent.IsInactive) { return "AMBER" }
    # BlueprintMissingUri: suppress for Copilot Studio only
    # Other external blueprints (e.g. Conditional Access Agent) DO have URIs — check is valid for them
    # Copilot Studio uses Power Platform connectors, not blueprint OBO — URI not applicable
    if ($Agent.BlueprintMissingUri -and -not $Agent.IsCopilotStudioAgent) { return "AMBER" }
    # BlueprintMissingScope: same reasoning as URI
    if ($Agent.BlueprintMissingScope -and -not $Agent.IsCopilotStudioAgent) { return "AMBER" }
    # BlueprintNoCredential: suppress for ALL external blueprints
    # Credential data lives on the application object in the publisher's tenant — inaccessible
    # $null already set in $row for external blueprints so this would be $null anyway,
    # but guard explicitly for clarity
    if ($Agent.BlueprintNoCredential -and -not $Agent.BlueprintIsExternal) { return "AMBER" }
    return "GREEN"
}


# ─────────────────────────────────────────────────────────────────
# STEP 0 — CONNECT
# ─────────────────────────────────────────────────────────────────
Write-Host "`n📡 Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph `
    -TenantId $TenantId `
    -Scopes @(
    "AgentIdentity.Read.All",
    "Application.Read.All",
    "User.Read.All"
) `
    -NoWelcome `
    -ContextScope Process   # ← isolates this session's token from the persistent cache; enforces least-privilege per session
Write-Host "✅ Connected`n" -ForegroundColor Green


# ─────────────────────────────────────────────────────────────────
# STEP 1 — FETCH MODERN AGENT IDENTITIES
# ─────────────────────────────────────────────────────────────────
Write-Host "🔍 [1/6] Fetching modern agent identities..." -ForegroundColor Cyan
$modernAgents = Get-AllPages `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentity?`$select=id,displayName,appId,accountEnabled,createdDateTime,signInAudience,agentIdentityBlueprintId,servicePrincipalType,tags&`$count=true" `
    -Headers @{ ConsistencyLevel = "eventual" }
Write-Host "   Found $($modernAgents.Count) modern agent identities"


# ─────────────────────────────────────────────────────────────────
# STEP 2 — DETECT CLASSIC (SP-BACKED) AGENTS
# ─────────────────────────────────────────────────────────────────
Write-Host "🔍 [2/6] Identifying classic (SP-backed) agents..." -ForegroundColor Cyan
$classicAgents = Get-AllPages `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=tags/any(t:startsWith(t, 'Agentic'))&`$select=id,displayName,appId,accountEnabled,createdDateTime,servicePrincipalType,tags&`$count=true" `
    -Headers @{ ConsistencyLevel = "eventual" }
# Deduplicate — modern agents also carry AgenticInstance tag
$modernAgentIds = @($modernAgents | Select-Object -ExpandProperty id)
$classicOnly = @($classicAgents | Where-Object { $_.id -notin $modernAgentIds })
# Classify by platform using tags observed empirically
$classicCS = @($classicOnly | Where-Object { $_.tags -contains "AgentCreatedBy:CopilotStudio" })
$classicFoundry = @($classicOnly | Where-Object { $_.tags -contains "AgentCreatedBy:Foundry" })
$classicUnknown = @($classicOnly | Where-Object {
        $_.tags -notcontains "AgentCreatedBy:CopilotStudio" -and
        $_.tags -notcontains "AgentCreatedBy:Foundry"
    })
Write-Host "   Modern agent identities  : $($modernAgents.Count)"
Write-Host "   Classic — Copilot Studio : $($classicCS.Count)"
Write-Host "   Classic — Foundry        : $($classicFoundry.Count)"
Write-Host "   Classic — Unknown source : $($classicUnknown.Count)"
Write-Host "   Total                    : $($modernAgents.Count + $classicOnly.Count)"


# ─────────────────────────────────────────────────────────────────
# STEP 3 — FETCH BLUEPRINTS (TWO ENDPOINTS)
# ─────────────────────────────────────────────────────────────────
Write-Host "🔍 [3/6] Fetching agent identity blueprints..." -ForegroundColor Cyan
# 3a — Blueprint PRINCIPALS (service principals in your tenant)
#       Covers ALL blueprints (Company-owned + external/Microsoft-published)
#       Does NOT carry credentials, identifierUris, api — those are on the application object
$blueprintPrincipals = Get-AllPages `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal?`$select=id,appId,displayName,accountEnabled,createdDateTime,servicePrincipalType,oauth2PermissionScopes,signInAudience,servicePrincipalNames,appOwnerOrganizationId&`$count=true" `
    -Headers @{ ConsistencyLevel = "eventual" }
# 3b — Blueprint APPLICATION OBJECTS (only blueprints Company owns/published)
#       Has credentials, identifierUris, api.oauth2PermissionScopes, signInAudience
#       agentIdentityBlueprint inherits from 'application' — these are the fields we need
#       for credential health checks. External blueprints (e.g. Microsoft) won't appear here.
$blueprintApps = Get-AllPages `
    -Uri "https://graph.microsoft.com/v1.0/applications/microsoft.graph.agentIdentityBlueprint?`$select=id,appId,displayName,identifierUris,api,keyCredentials,passwordCredentials,signInAudience,publisherDomain,requiredResourceAccess,createdDateTime&`$count=true" `
    -Headers @{ ConsistencyLevel = "eventual" }
# Build lookup maps keyed by BOTH id and appId for resilience
$blueprintPrincipalMap = @{}
foreach ($bp in $blueprintPrincipals) {
    $blueprintPrincipalMap[$bp.id] = $bp
    $blueprintPrincipalMap[$bp.appId] = $bp
}
# Credential/API data lives on the application object — separate map
$blueprintAppMap = @{}
foreach ($bpa in $blueprintApps) {
    $blueprintAppMap[$bpa.id] = $bpa
    $blueprintAppMap[$bpa.appId] = $bpa
}
Write-Host "   Blueprint principals (all tenants) : $($blueprintPrincipals.Count)"
Write-Host "   Blueprint apps (Company-owned only)  : $($blueprintApps.Count)"
# ── Step 3c — Fetch granted permissions per blueprint principal ───────────────
Write-Host "   Fetching granted permissions per blueprint..." -ForegroundColor Gray
$blueprintPermissionsMap = @{}
$resourceSpCache = @{}
foreach ($bp in $blueprintPrincipals) {
    $assignments = Get-SubResource `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($bp.id)/appRoleAssignments?`$select=appRoleId,resourceDisplayName,resourceId,createdDateTime"
    if ($assignments.Count -eq 0) {
        $blueprintPermissionsMap[$bp.id] = @()
        continue
    }
    # Build a set of declared permission GUIDs from requiredResourceAccess
    # so we can flag grants that exist in the tenant but were never declared in the manifest
    # Only possible for Company-owned blueprints where we have the application object
    $bpaForThisBp = $blueprintAppMap[$bp.appId]
    $declaredRoleIds = @{}
    if ($bpaForThisBp -and $bpaForThisBp.requiredResourceAccess) {
        foreach ($resource in $bpaForThisBp.requiredResourceAccess) {
            foreach ($access in $resource.resourceAccess) {
                $declaredRoleIds[$access.id] = $true
            }
        }
    }
    $permissionStrings = foreach ($assignment in $assignments) {
        $roleName = if ($assignment.appRoleId -eq "00000000-0000-0000-0000-000000000000") {
            "Default Access"
        }
        else {
            # Cache check — fetch resource SP's appRoles only once across all blueprints
            if (-not $resourceSpCache.ContainsKey($assignment.resourceId)) {
                $resourceSp = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($assignment.resourceId)?`$select=appRoles" `
                    -ErrorAction SilentlyContinue
                $resourceSpCache[$assignment.resourceId] = @($resourceSp.appRoles)
            }
            $matchedRole = $resourceSpCache[$assignment.resourceId] |
            Where-Object { $_.id -eq $assignment.appRoleId }
            if ($matchedRole) { $matchedRole.value } else { $assignment.appRoleId }
        }
        # Flag permissions that are granted but not declared in the manifest
        # Only flagged for Company-owned blueprints — external blueprints have no manifest visible
        $isDeclared = if ($null -eq $bpaForThisBp) {
            $true    # Can't check manifest for external blueprints — don't flag as undeclared
        }
        else {
            $declaredRoleIds.ContainsKey($assignment.appRoleId)
        }
        $suffix = if (-not $isDeclared) { " ⚠️ UNDECLARED" } else { "" }
        "$($assignment.resourceDisplayName) → $roleName$suffix"
    }
    $blueprintPermissionsMap[$bp.id] = @($permissionStrings)
}
Write-Host "   ✅ Permissions fetched for $($blueprintPrincipals.Count) blueprints"
$copilotStudioBlueprintAppId = ($blueprintPrincipals |
    Where-Object { $_.displayName -like "*Copilot Studio*" } |
    Select-Object -First 1 -ExpandProperty appId)
if ($copilotStudioBlueprintAppId) {
    Write-Host "   Copilot Studio shared blueprint appId: $copilotStudioBlueprintAppId" -ForegroundColor Gray
}
else {
    Write-Host "   ⚠️ Could not identify Copilot Studio shared blueprint" -ForegroundColor Yellow
}

# ─────────────────────────────────────────────────────────────────
# STEP 4 — ENRICH EACH MODERN AGENT
# ─────────────────────────────────────────────────────────────────
Write-Host "🔍 [4/6] Enriching modern agents with ownership + blueprint data..." -ForegroundColor Cyan
Write-Host "   (One API call per agent — may take a moment)`n"
$agentReport = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0
foreach ($agent in $modernAgents) {
    $counter++
    Write-Progress `
        -Activity   "Enriching agent identities" `
        -Status     "$counter / $($modernAgents.Count) — $($agent.displayName)" `
        -PercentComplete (($counter / [Math]::Max($modernAgents.Count, 1)) * 100)
    # ── Owners ────────────────────────────────────────────────────
    # URL uses cast segment /microsoft.graph.agentIdentity/ as per Graph docs
    $owners = Get-SubResource `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($agent.id)/microsoft.graph.agentIdentity/owners?`$select=id,displayName,userPrincipalName,accountEnabled"
    # ── Sponsors ──────────────────────────────────────────────────
    $sponsors = Get-SubResource `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($agent.id)/microsoft.graph.agentIdentity/sponsors?`$select=id,displayName,userPrincipalName,accountEnabled"

    # ── Agent Identity Permissions ────────────────────────────────────
    # Direct appRoleAssignments on the agent identity itself — separate from blueprint permissions
    # Blueprint permissions = what ALL agents from this blueprint can do as a class
    # Agent permissions    = what THIS specific agent instance has been directly granted
    $agentAssignments = Get-SubResource `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($agent.id)/appRoleAssignments?`$select=appRoleId,resourceDisplayName,resourceId"
    $agentPermissionStrings = foreach ($assignment in $agentAssignments) {
        $roleName = if ($assignment.appRoleId -eq "00000000-0000-0000-0000-000000000000") {
            "Default Access"
        }
        else {
            # Reuse $resourceSpCache built in Step 3c — no duplicate lookups
            if (-not $resourceSpCache.ContainsKey($assignment.resourceId)) {
                $resourceSp = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($assignment.resourceId)?`$select=appRoles" `
                    -ErrorAction SilentlyContinue
                $resourceSpCache[$assignment.resourceId] = @($resourceSp.appRoles)
            }
            $matchedRole = $resourceSpCache[$assignment.resourceId] |
            Where-Object { $_.id -eq $assignment.appRoleId }
            if ($matchedRole) { $matchedRole.value } else { $assignment.appRoleId }
        }
        "$($assignment.resourceDisplayName) → $roleName"
    }
    $agentPermissionsString = $agentPermissionStrings -join " | "
    $agentPermissionCount = $agentAssignments.Count

    # ── Blueprint Lookup ──────────────────────────────────────────
    $bpPrincipal = $null
    $bpApp = $null   # application object — has credentials
    $hasBlueprint = $false
    $bpMissingUri = $false
    $bpMissingScope = $false
    $bpNoCredential = $false
    $hasSecret = $false
    $hasCert = $false
    $credExpired = $false
    $credExpiringSoon = $false
    $credNearestExpiry = $null
    if ($agent.agentIdentityBlueprintId) {
        $bpPrincipal = $blueprintPrincipalMap[$agent.agentIdentityBlueprintId]
        $bpApp = $blueprintAppMap[$agent.agentIdentityBlueprintId]
        $hasBlueprint = $true
        if ($bpPrincipal) {
            # Identifier URI — present in servicePrincipalNames as anything beyond the bare appId
            $bpMissingUri = (
                @($bpPrincipal.servicePrincipalNames |
                    Where-Object { $_ -ne $bpPrincipal.appId }
                ).Count -eq 0
            )
            # OAuth2 scope — oauth2PermissionScopes is on the SP
            $bpMissingScope = (
                $null -eq $bpPrincipal.oauth2PermissionScopes -or
                $bpPrincipal.oauth2PermissionScopes.Count -eq 0
            )
            # ── Credential checks — application object ONLY (Company-owned blueprints) ─
            if ($bpApp) {
                $allCreds = [System.Collections.Generic.List[object]]::new()
                if ($bpApp.passwordCredentials) { foreach ($c in $bpApp.passwordCredentials) { $allCreds.Add($c) } }
                if ($bpApp.keyCredentials) { foreach ($c in $bpApp.keyCredentials) { $allCreds.Add($c) } }
                $bpNoCredential = ($allCreds.Count -eq 0)
                $hasSecret = ($null -ne $bpApp.passwordCredentials -and $bpApp.passwordCredentials.Count -gt 0)
                $hasCert = ($null -ne $bpApp.keyCredentials -and $bpApp.keyCredentials.Count -gt 0)
                $now = Get-Date
                foreach ($cred in $allCreds) {
                    if ($null -ne $cred.endDateTime) {
                        $expiry = [datetime]$cred.endDateTime
                        if ($expiry -lt $now) { $credExpired = $true }
                        elseif ($expiry -lt $now.AddDays($CredentialWarnDays)) { $credExpiringSoon = $true }
                        if ($null -eq $credNearestExpiry -or $expiry -lt $credNearestExpiry) {
                            $credNearestExpiry = $expiry
                        }
                    }
                }
            }
            # If $bpApp is null (external blueprint): $bpNoCredential/$hasSecret/$hasCert
            # remain $false — their meaning shifts to "unknown" not "absent".
            # BlueprintIsExternal = $true in the report communicates this.
        }
    }

    # Copilot Studio agents share one tenant-wide blueprint
    # The shared blueprint has a well-known display name
    # Path 1: tag on the agent SP (may not be present on brand-new agents)
    $isCopilotStudioAgent = ($agent.tags -contains "AgentCreatedBy:CopilotStudio")
    # Path 2: blueprint principal display name (requires blueprint map to resolve)
    if (-not $isCopilotStudioAgent -and $null -ne $bpPrincipal) {
        $isCopilotStudioAgent = ($bpPrincipal.displayName -like "*Copilot Studio*")
    }
    # Path 3: blueprint app display name (from blueprintAppMap — Company-owned blueprints)
    if (-not $isCopilotStudioAgent -and $null -ne $bpApp) {
        $isCopilotStudioAgent = ($bpApp.displayName -like "*Copilot Studio*")
    }
    if (-not $isCopilotStudioAgent -and $null -ne $copilotStudioBlueprintAppId) {
        $isCopilotStudioAgent = ($agent.agentIdentityBlueprintId -eq $copilotStudioBlueprintAppId) # Path 4
    }

    # ── Ownership Flags ───────────────────────────────────────────
    $noOwner = ($owners.Count -eq 0)
    $noSponsor = ($sponsors.Count -eq 0)
    $disabledOwner = ($owners   | Where-Object { $_.accountEnabled -eq $false }).Count -gt 0
    $disabledSponsor = ($sponsors | Where-Object { $_.accountEnabled -eq $false }).Count -gt 0
    # ── Age & Lifecycle ───────────────────────────────────────────
    $createdDate = if ($null -ne $agent.createdDateTime) { [datetime]$agent.createdDateTime } else { $null }
    $daysSinceCreate = if ($null -ne $createdDate) { ([datetime]::UtcNow - $createdDate).Days } else { $null }
    $neverSignedIn = $false   # Placeholder — needs Reports.Read.All + beta sign-in endpoint
    $isInactive = $false   # Placeholder
    # ── BUG FIX: Pre-compute all nullable strings BEFORE [PSCustomObject]@{} ──
    # Reason: ?. null-conditional has parser ambiguity inside hashtable literal @{ Key = $obj?.Prop }
    # even on PS 7.6.x. Pre-computing into plain variables avoids the issue entirely.
    $createdDateStr = if ($null -ne $createdDate) { $createdDate.ToString("yyyy-MM-dd") } else { $null }
    $nearestExpiryStr = if ($null -ne $credNearestExpiry) { $credNearestExpiry.ToString("yyyy-MM-dd") } else { $null }
    $blueprintDisplayName = if ($null -ne $bpPrincipal) { $bpPrincipal.displayName } else { $null }
    $blueprintIsExternal = if ($null -ne $bpPrincipal -and $null -eq $bpApp) { $true } else { $false }
    $blueprintMultiTenant = if ($null -ne $bpApp) { $bpApp.signInAudience -ne "AzureADMyOrg" } else { $null }
    # BUG FIX: Pre-compute switch statement — switch as hashtable value is technically a
    # statement used as expression, which can behave unexpectedly in some PS builds.
    $ownershipPattern = switch ($true) {
        ($noOwner -and $noSponsor) { "Neither"; break }
        ($noOwner -and -not $noSponsor) { "Sponsor Only"; break }
        (-not $noOwner -and $noSponsor) { "Owner Only"; break }
        default { "Both" }
    }
    # ── Build Report Row ──────────────────────────────────────────
    $row = [PSCustomObject]@{
        # Identity
        AgentName               = $agent.displayName
        AgentId                 = $agent.id
        AgentAppId              = $agent.appId
        IsEnabled               = $agent.accountEnabled
        CreatedDate             = $createdDateStr          # pre-computed — no ?. in hashtable
        DaysSinceCreation       = $daysSinceCreate
        SignInAudience          = $agent.signInAudience
        # Classification
        IsClassic               = $false
        IsModern                = $true
        HasBlueprint            = $hasBlueprint
        BlueprintId             = $agent.agentIdentityBlueprintId
        BlueprintName           = $blueprintDisplayName    # pre-computed — no ?. in hashtable
        BlueprintIsExternal     = $blueprintIsExternal     # true = Microsoft/3rd party blueprint, creds not visible
        BlueprintMultiTenant    = $blueprintMultiTenant    # pre-computed — no ?. in hashtable
        # Blueprint Health (only populated for Company-owned blueprints)
        BlueprintMissingUri     = $bpMissingUri
        BlueprintMissingScope   = $bpMissingScope
        BlueprintNoCredential   = if ($isCopilotStudioAgent) { $null } else { $bpNoCredential }
        # Credential Risk
        HasClientSecret         = if ($isCopilotStudioAgent) { $null } else { $hasSecret }
        HasCertificate          = if ($isCopilotStudioAgent) { $null } else { $hasCert }
        CredentialExpired       = $credExpired
        CredentialExpiringSoon  = $credExpiringSoon
        NearestCredentialExpiry = $nearestExpiryStr        # pre-computed — no ?. in hashtable
        # Ownership
        NoOwner                 = $noOwner
        OwnerCount              = $owners.Count
        Owners                  = ($owners   | Select-Object -ExpandProperty displayName) -join "; "
        OwnerUPNs               = ($owners   | Select-Object -ExpandProperty userPrincipalName) -join "; "
        HasDisabledOwner        = $disabledOwner
        NoSponsor               = $noSponsor
        SponsorCount            = $sponsors.Count
        Sponsors                = ($sponsors | Select-Object -ExpandProperty displayName) -join "; "
        SponsorUPNs             = ($sponsors | Select-Object -ExpandProperty userPrincipalName) -join "; "
        HasDisabledSponsor      = $disabledSponsor
        OwnershipPattern        = $ownershipPattern        # pre-computed — no switch in hashtable
        # Activity (placeholder — add sign-in log enrichment in a future step)
        NeverSignedIn           = $neverSignedIn
        IsInactive              = $isInactive
        # LastSignInDate          = "N/A — add Reports.Read.All + beta sign-in query"
        # Permissions — direct grants on this agent identity (not inherited from blueprint)
        AgentPermissionCount    = $agentPermissionCount
        AgentPermissions        = $agentPermissionsString   # e.g. "Microsoft Graph → AgentIdentity.Read.All"
        IsCopilotStudioAgent    = $isCopilotStudioAgent
        # RAG Risk — set AFTER object is created
        RiskRating              = ""
    }
    # Set RAG now that $row is a valid object
    $row.RiskRating = Get-RAGStatus -Agent $row
    $agentReport.Add($row)
}
Write-Progress -Activity "Enriching agent identities" -Completed


# ─────────────────────────────────────────────────────────────────
# STEP 5 — CLASSIC AGENT SUMMARY
# ─────────────────────────────────────────────────────────────────
Write-Host "`n🔍 [5/6] Building classic agent summary..." -ForegroundColor Cyan
$classicReport = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($sp in $classicOnly) {
    $spCreatedDate = if ($null -ne $sp.createdDateTime) { [datetime]$sp.createdDateTime } else { $null }
    $spCreatedDateStr = if ($null -ne $spCreatedDate) { $spCreatedDate.ToString("yyyy-MM-dd") } else { $null }
    $spDaysSince = if ($null -ne $spCreatedDate) { ([datetime]::UtcNow - $spCreatedDate).Days } else { $null }
    $platform = switch ($true) {
        ($sp.tags -contains "AgentCreatedBy:CopilotStudio") { "Copilot Studio"; break }
        ($sp.tags -contains "AgentCreatedBy:Foundry") { "Foundry"; break }
        default { "Unknown" }
    }
    $classicRow = [PSCustomObject]@{
        AgentName         = $sp.displayName
        AgentId           = $sp.id
        AgentAppId        = $sp.appId
        IsEnabled         = $sp.accountEnabled
        CreatedDate       = $spCreatedDateStr
        DaysSinceCreation = $spDaysSince
        Platform          = $platform
        IsClassic         = $true
        IsModern          = $false
        HasBlueprint      = $false
        BlueprintId       = $null
        OwnershipPattern  = "Unknown — check Enterprise Apps blade"
        RiskRating        = "RED"
        Notes             = "Classic SP-backed agent. No Agent ID governance. Migrate to Agent ID."
    }
    $classicReport.Add($classicRow)
}
Write-Host "   Built $($classicReport.Count) classic agent summary rows"


# ─────────────────────────────────────────────────────────────────
# STEP 6 — BLUEPRINT HEALTH REPORT
# ─────────────────────────────────────────────────────────────────
Write-Host "🔍 [6/6] Building blueprint health report..." -ForegroundColor Cyan
# Flag high-risk permissions — governance red flag if agent has write/delete access
$highRiskKeywords = @("Write", "ReadWrite", "Delete", "FullControl", "Mail.Send", "Files.ReadWrite")
$blueprintReport = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($bp in $blueprintPrincipals) {
    $grantedPermissions = @($blueprintPermissionsMap[$bp.id])
    $permissionCount = $grantedPermissions.Count
    $permissionsString = $grantedPermissions -join " | "   # pipe-delimited for Excel readability
    $hasHighRiskPermissions = ($grantedPermissions | Where-Object {
            $perm = $_
            $highRiskKeywords | Where-Object { $perm -like "*$_*" }
        }).Count -gt 0
    # Find matching application object for credential data (Company-owned only)
    $bpApp = $blueprintAppMap[$bp.appId]
    $isExternal = ($null -eq $bpApp)
    $bpHasSecret = if ($bpApp) { $null -ne $bpApp.passwordCredentials -and $bpApp.passwordCredentials.Count -gt 0 } else { $null }
    $bpHasCert = if ($bpApp) { $null -ne $bpApp.keyCredentials -and $bpApp.keyCredentials.Count -gt 0 } else { $null }
    $bpHasUri = if ($bpApp) { $null -ne $bpApp.identifierUris -and $bpApp.identifierUris.Count -gt 0 } else { $null }
    $bpHasScope = if ($bpApp) { $null -ne $bpApp.api -and $null -ne $bpApp.api.oauth2PermissionScopes -and $bpApp.api.oauth2PermissionScopes.Count -gt 0 } else { $null }
    $bpSI = if ($bpApp) { $bpApp.signInAudience } else { $null }
    $bpPubDomain = if ($bpApp) { $bpApp.publisherDomain } else { "External (not visible)" }
    $bpIdentUri = if ($bpApp) { ($bpApp.identifierUris -join "; ") } else { "External (not visible)" }
    $credType = if ($isExternal) {
        "External — credentials not visible from this tenant"
    }
    else {
        switch ($true) {
            ($bpHasSecret -and $bpHasCert) { "Secret + Certificate (HIGH RISK — remove secret)"; break }
            $bpHasSecret { "Secret Only (HIGH RISK)"; break }
            $bpHasCert { "Certificate"; break }
            default { "None (BROKEN — no credential configured)" }
        }
    }
    $agentsFromBp = @($agentReport | Where-Object { $_.BlueprintId -eq $bp.appId })
    $bpCreatedStr = if ($null -ne $bp.createdDateTime) { ([datetime]$bp.createdDateTime).ToString("yyyy-MM-dd") } else { $null }
    $bpRow = [PSCustomObject]@{
        BlueprintName          = $bp.displayName
        BlueprintAppId         = $bp.appId
        BlueprintObjectId      = $bp.id
        CreatedDate            = $bpCreatedStr
        IsExternalBlueprint    = $isExternal
        PublisherDomain        = $bpPubDomain
        IsMultiTenant          = if ($null -ne $bpSI) { $bpSI -ne "AzureADMyOrg" } else { $null }
        ChildAgentCount        = $agentsFromBp.Count
        HasIdentifierUri       = $bpHasUri
        IdentifierUri          = $bpIdentUri
        HasOAuthScope          = $bpHasScope
        HasClientSecret        = $bpHasSecret
        HasCertificate         = $bpHasCert
        CredentialType         = $credType
        RequiredPermissions    = if ($bpApp) { ($bpApp.requiredResourceAccess | Measure-Object).Count } else { $null }
        AgentNames             = ($agentsFromBp | Select-Object -ExpandProperty AgentName) -join "; "
        GrantedPermissionCount = $permissionCount
        HasHighRiskPermissions = $hasHighRiskPermissions
        GrantedPermissions     = $permissionsString
    }
    $blueprintReport.Add($bpRow)
}


# ─────────────────────────────────────────────────────────────────
# CONSOLE SUMMARY
# ─────────────────────────────────────────────────────────────────
$allModern = @($agentReport)
$totalAgents = $allModern.Count + $classicReport.Count
$redCount = ($allModern | Where-Object { $_.RiskRating -eq "RED" }).Count + $classicReport.Count
$amberCount = ($allModern | Where-Object { $_.RiskRating -eq "AMBER" }).Count
$greenCount = ($allModern | Where-Object { $_.RiskRating -eq "GREEN" }).Count
Write-Host ("`n" + ("=" * 65)) -ForegroundColor White
Write-Host "  AGENT GOVERNANCE REPORT" -ForegroundColor White
Write-Host "  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Gray
Write-Host ("=" * 65) -ForegroundColor White
Write-Host "`n📊 INVENTORY" -ForegroundColor Cyan
Write-Host "   Total agents in scope        : $totalAgents"
Write-Host "   Modern (Agent ID-backed)     : $($allModern.Count)"
Write-Host "   Classic (SP-backed)          : $($classicReport.Count)"
Write-Host "     ↳ Copilot Studio classic   : $($classicCS.Count)"
Write-Host "     ↳ Foundry classic          : $($classicFoundry.Count)"
Write-Host "     ↳ Unknown platform         : $($classicUnknown.Count)"
Write-Host "   Blueprint principals (total) : $($blueprintPrincipals.Count)"
Write-Host "   Blueprint apps (Company-owned): $($blueprintApps.Count)"
Write-Host "`n🏷️  OWNERSHIP (modern agents only)" -ForegroundColor Cyan
Write-Host "   Both owner + sponsor         : $(($allModern | Where-Object { $_.OwnershipPattern -eq 'Both'         }).Count)"
Write-Host "   Sponsor only                 : $(($allModern | Where-Object { $_.OwnershipPattern -eq 'Sponsor Only' }).Count)"
Write-Host "   Owner only                   : $(($allModern | Where-Object { $_.OwnershipPattern -eq 'Owner Only'   }).Count)"
$neitherCount = ($allModern | Where-Object { $_.OwnershipPattern -eq 'Neither' }).Count
$neitherColor = if ($neitherCount -gt 0) { "Red" } else { "White" }
Write-Host "   Neither owner nor sponsor    : $neitherCount" -ForegroundColor $neitherColor
Write-Host "   Disabled owner account       : $(($allModern | Where-Object { $_.HasDisabledOwner   }).Count)"
Write-Host "   Disabled sponsor account     : $(($allModern | Where-Object { $_.HasDisabledSponsor }).Count)"
Write-Host "`n🔐 CREDENTIAL RISK (Company-owned blueprints only)" -ForegroundColor Cyan
$secretCount = ($blueprintReport | Where-Object { $_.HasClientSecret -eq $true }).Count
$brokenCount = ($blueprintReport | Where-Object { $_.CredentialType -like '*BROKEN*' }).Count
$expiredCount = ($allModern       | Where-Object { $_.CredentialExpired }).Count
$expiringCount = ($allModern       | Where-Object { $_.CredentialExpiringSoon }).Count
Write-Host "   Blueprints with secrets      : $secretCount"  -ForegroundColor $(if ($secretCount -gt 0) { "Red" } else { "White" })
Write-Host "   Blueprints with certs        : $(($blueprintReport | Where-Object { $_.HasCertificate -eq $true }).Count)"
Write-Host "   Blueprints with no cred      : $brokenCount"  -ForegroundColor $(if ($brokenCount -gt 0) { "Red" } else { "White" })
Write-Host "   Credentials expired          : $expiredCount" -ForegroundColor $(if ($expiredCount -gt 0) { "Red" } else { "White" })
Write-Host "   Expiring within $CredentialWarnDays days      : $expiringCount" -ForegroundColor $(if ($expiringCount -gt 0) { "Yellow" } else { "White" })
Write-Host "`n🏗️  BLUEPRINT HEALTH" -ForegroundColor Cyan
$missingUriCount = ($blueprintReport | Where-Object { $_.HasIdentifierUri -eq $false }).Count
Write-Host "   Missing identifier URI       : $missingUriCount"   -ForegroundColor $(if ($missingUriCount -gt 0) { "Yellow" } else { "White" })
Write-Host "   Missing OAuth scope          : $(($blueprintReport | Where-Object { $_.HasOAuthScope -eq $false }).Count)"
Write-Host "   Zero child agents            : $(($blueprintReport | Where-Object { $_.ChildAgentCount -eq 0 }).Count)"
Write-Host "   Multi-tenant blueprints      : $(($blueprintReport | Where-Object { $_.IsMultiTenant -eq $true }).Count)"
Write-Host "   External (Microsoft/3rd-pty) : $(($blueprintReport | Where-Object { $_.IsExternalBlueprint }).Count)"
Write-Host "`n🚦 RISK SUMMARY" -ForegroundColor Cyan
Write-Host "   🔴 RED   : $redCount agents"   -ForegroundColor Red
Write-Host "   🟡 AMBER : $amberCount agents"  -ForegroundColor Yellow
Write-Host "   🟢 GREEN : $greenCount agents"  -ForegroundColor Green
Write-Host ("`n" + ("=" * 65)) -ForegroundColor White


# ─────────────────────────────────────────────────────────────────
# EXPORT — Single .xlsx workbook with three sheets
# ─────────────────────────────────────────────────────────────────
Write-Host "`n💾 Exporting report to Excel workbook..." -ForegroundColor Cyan
# Sheet 1 — Modern Agents
$agentReport | Export-Excel `
    -Path        $ReportOutputPath `
    -WorksheetName "Modern Agents" `
    -TableName   "ModernAgents" `
    -TableStyle  Medium6 `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow
# Sheet 2 — Blueprints (append to same file — no -ClearSheet, no overwrite)
$blueprintReport | Export-Excel `
    -Path        $ReportOutputPath `
    -WorksheetName "Blueprints" `
    -TableName   "Blueprints" `
    -TableStyle  Medium4 `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow
# Sheet 3 — Classic Agents
if ($classicReport.Count -gt 0) {
    $classicReport | Export-Excel `
        -Path        $ReportOutputPath `
        -WorksheetName "Classic Agents" `
        -TableName   "ClassicAgents" `
        -TableStyle  Medium7 `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow
}
Write-Host "   ✅ Workbook saved : $ReportOutputPath" -ForegroundColor Green
Write-Host "`n✅ Report complete.`n"                  -ForegroundColor Green