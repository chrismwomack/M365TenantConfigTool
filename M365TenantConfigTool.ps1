<#
.SYNOPSIS
    Microsoft 365 Tenant Configuration Snapshot Tool using Graph Beta UTCM APIs.

.DESCRIPTION
    Menu-driven PowerShell tool for creating, managing, and viewing tenant configuration
    snapshots via the Unified Tenant Configuration Management (UTCM) APIs in Microsoft
    Graph Beta. Supports all six workloads: Entra, Exchange, Intune, Teams,
    Security & Compliance (Defender/Purview).

.NOTES
    Requires: Microsoft.Graph.Authentication module
    API Permission: ConfigurationMonitoring.ReadWrite.All
    License: Microsoft 365 E3 or Microsoft Entra ID P1
#>

#Requires -Version 5.1

# ============================================================================
# Configuration
# ============================================================================

$script:GraphBaseUri = "https://graph.microsoft.com/beta"
$script:UTCMBase = "$script:GraphBaseUri/admin/configurationManagement"
$script:SnapshotJobsEndpoint = "$script:UTCMBase/configurationSnapshotJobs"
$script:CreateSnapshotEndpoint = "$script:UTCMBase/configurationSnapshots/createSnapshot"
$script:SnapshotsEndpoint = "$script:UTCMBase/configurationSnapshots"
$script:RequiredScopes = @("ConfigurationMonitoring.ReadWrite.All")
$script:UTCMAppId = "03b07b79-c5bc-4b5e-9bfa-13acf4a99998"
$script:GraphAppId = "00000003-0000-0000-c000-000000000000"
$script:IsConnected = $false

# ============================================================================
# Resource Type Definitions by Workload
# ============================================================================

$script:Workloads = [ordered]@{

    "Microsoft Entra" = @(
        "microsoft.entra.administrativeunit"
        "microsoft.entra.application"
        "microsoft.entra.authenticationcontextclassreference"
        "microsoft.entra.authenticationmethodpolicy"
        "microsoft.entra.authenticationmethodpolicyauthenticator"
        "microsoft.entra.authenticationmethodpolicyemail"
        "microsoft.entra.authenticationmethodpolicyfido2"
        "microsoft.entra.authenticationmethodpolicysms"
        "microsoft.entra.authenticationmethodpolicysoftware"
        "microsoft.entra.authenticationmethodpolicytemporary"
        "microsoft.entra.authenticationmethodpolicyvoice"
        "microsoft.entra.authenticationmethodpolicyx509"
        "microsoft.entra.authenticationstrengthpolicy"
        "microsoft.entra.authorizationpolicy"
        "microsoft.entra.conditionalaccesspolicy"
        "microsoft.entra.crosstenantaccesspolicy"
        "microsoft.entra.crosstenantaccesspolicyconfigurationdefault"
        "microsoft.entra.crosstenantaccesspolicyconfigurationpartner"
        "microsoft.entra.entitlementmanagementaccesspackage"
        "microsoft.entra.entitlementmanagementaccesspackageassignmentpolicy"
        "microsoft.entra.entitlementmanagementaccesspackagecatalog"
        "microsoft.entra.entitlementmanagementaccesspackagecatalogresource"
        "microsoft.entra.entitlementmanagementconnectedorganization"
        "microsoft.entra.externalidentitypolicy"
        "microsoft.entra.group"
        "microsoft.entra.grouplifecyclepolicy"
        "microsoft.entra.namedlocationpolicy"
        "microsoft.entra.roledefinition"
        "microsoft.entra.roleeligibilityschedulerequest"
        "microsoft.entra.rolesetting"
        "microsoft.entra.securitydefaults"
        "microsoft.entra.serviceprincipal"
        "microsoft.entra.socialidentityprovider"
        "microsoft.entra.tenantdetails"
        "microsoft.entra.tokenlifetimepolicy"
        "microsoft.entra.user"
    )

    "Microsoft Exchange" = @(
        "microsoft.exchange.accepteddomain"
        "microsoft.exchange.activesyncdeviceaccessrule"
        "microsoft.exchange.antiphishpolicy"
        "microsoft.exchange.antiphishrule"
        "microsoft.exchange.applicationaccesspolicy"
        "microsoft.exchange.atppolicyforo365"
        "microsoft.exchange.authenticationpolicy"
        "microsoft.exchange.authenticationpolicyassignment"
        "microsoft.exchange.availabilityaddressspace"
        "microsoft.exchange.availabilityconfig"
        "microsoft.exchange.calendarprocessing"
        "microsoft.exchange.casmailboxplan"
        "microsoft.exchange.casmailboxsettings"
        "microsoft.exchange.dataclassification"
        "microsoft.exchange.dataencryptionpolicy"
        "microsoft.exchange.distributiongroup"
        "microsoft.exchange.dkimsigningconfig"
        "microsoft.exchange.emailaddresspolicy"
        "microsoft.exchange.groupsettings"
        "microsoft.exchange.hostedconnectionfilterpolicy"
        "microsoft.exchange.hostedcontentfilterpolicy"
        "microsoft.exchange.hostedcontentfilterrule"
        "microsoft.exchange.hostedoutboundspamfilterpolicy"
        "microsoft.exchange.hostedoutboundspamfilterrule"
        "microsoft.exchange.inboundconnector"
        "microsoft.exchange.intraorganizationconnector"
        "microsoft.exchange.irmconfiguration"
        "microsoft.exchange.journalrule"
        "microsoft.exchange.mailboxautoreplyconfiguration"
        "microsoft.exchange.mailboxcalendarfolder"
        "microsoft.exchange.mailboxpermission"
        "microsoft.exchange.mailboxplan"
        "microsoft.exchange.mailboxsettings"
        "microsoft.exchange.mailcontact"
        "microsoft.exchange.mailtips"
        "microsoft.exchange.malwarefilterpolicy"
        "microsoft.exchange.malwarefilterrule"
        "microsoft.exchange.managementrole"
        "microsoft.exchange.managementroleassignment"
        "microsoft.exchange.messageclassification"
        "microsoft.exchange.mobiledevicemailboxpolicy"
        "microsoft.exchange.omeconfiguration"
        "microsoft.exchange.onpremisesorganization"
        "microsoft.exchange.organizationconfig"
        "microsoft.exchange.organizationrelationship"
        "microsoft.exchange.outboundconnector"
        "microsoft.exchange.owamailboxpolicy"
        "microsoft.exchange.partnerapplication"
        "microsoft.exchange.perimeterconfiguration"
        "microsoft.exchange.place"
        "microsoft.exchange.policytipconfig"
        "microsoft.exchange.quarantinepolicy"
        "microsoft.exchange.recipientpermission"
        "microsoft.exchange.remotedomain"
        "microsoft.exchange.reportsubmissionpolicy"
        "microsoft.exchange.reportsubmissionrule"
        "microsoft.exchange.resourceconfiguration"
        "microsoft.exchange.roleassignmentpolicy"
        "microsoft.exchange.rolegroup"
        "microsoft.exchange.safeattachmentpolicy"
        "microsoft.exchange.safeattachmentrule"
        "microsoft.exchange.safelinkspolicy"
        "microsoft.exchange.safelinksrule"
        "microsoft.exchange.sharedmailbox"
        "microsoft.exchange.sharingpolicy"
        "microsoft.exchange.transportconfig"
        "microsoft.exchange.transportrule"
    )

    "Microsoft Intune" = @(
        "microsoft.intune.accountprotectionlocalusergroupmembershippolicy"
        "microsoft.intune.accountprotectionpolicy"
        "microsoft.intune.antiviruspolicywindows10settingcatalog"
        "microsoft.intune.appconfigurationpolicy"
        "microsoft.intune.applicationcontrolpolicywindows10"
        "microsoft.intune.appprotectionpolicyandroid"
        "microsoft.intune.appprotectionpolicyios"
        "microsoft.intune.attacksurfacereductionrulespolicywindows10configmanager"
        "microsoft.intune.deviceandappmanagementassignmentfilter"
        "microsoft.intune.devicecategory"
        "microsoft.intune.devicecleanuprule"
        "microsoft.intune.devicecompliancepolicyandroid"
        "microsoft.intune.devicecompliancepolicyandroiddeviceowner"
        "microsoft.intune.devicecompliancepolicyandroidworkprofile"
        "microsoft.intune.devicecompliancepolicyios"
        "microsoft.intune.devicecompliancepolicymacos"
        "microsoft.intune.devicecompliancepolicywindows10"
        "microsoft.intune.deviceconfigurationadministrativetemplatepolicywindows10"
        "microsoft.intune.deviceconfigurationcustompolicywindows10"
        "microsoft.intune.deviceconfigurationdefenderforendpointonboardingpolicywindows10"
        "microsoft.intune.deviceconfigurationdeliveryoptimizationpolicywindows10"
        "microsoft.intune.deviceconfigurationdomainjoinpolicywindows10"
        "microsoft.intune.deviceconfigurationemailprofilepolicywindows10"
        "microsoft.intune.deviceconfigurationendpointprotectionpolicywindows10"
        "microsoft.intune.deviceconfigurationfirmwareinterfacepolicywindows10"
        "microsoft.intune.deviceconfigurationhealthmonitoringconfigurationpolicywindows10"
        "microsoft.intune.deviceconfigurationidentityprotectionpolicywindows10"
        "microsoft.intune.deviceconfigurationimportedpfxcertificatepolicywindows10"
        "microsoft.intune.deviceconfigurationkioskpolicywindows10"
        "microsoft.intune.deviceconfigurationnetworkboundarypolicywindows10"
        "microsoft.intune.deviceconfigurationpkcscertificatepolicywindows10"
        "microsoft.intune.deviceconfigurationpolicyandroiddeviceadministrator"
        "microsoft.intune.deviceconfigurationpolicyandroiddeviceowner"
        "microsoft.intune.deviceconfigurationpolicyandroidopensourceproject"
        "microsoft.intune.deviceconfigurationpolicyandroidworkprofile"
        "microsoft.intune.deviceconfigurationpolicyios"
        "microsoft.intune.deviceconfigurationpolicymacos"
        "microsoft.intune.deviceconfigurationpolicywindows10"
        "microsoft.intune.deviceconfigurationscepcertificatepolicywindows10"
        "microsoft.intune.deviceconfigurationsecureassessmentpolicywindows10"
        "microsoft.intune.deviceconfigurationsharedmultidevicepolicywindows10"
        "microsoft.intune.deviceconfigurationtrustedcertificatepolicywindows10"
        "microsoft.intune.deviceconfigurationvpnpolicywindows10"
        "microsoft.intune.deviceconfigurationwindowsteampolicywindows10"
        "microsoft.intune.deviceconfigurationwirednetworkpolicywindows10"
        "microsoft.intune.deviceenrollmentlimitrestriction"
        "microsoft.intune.deviceenrollmentplatformrestriction"
        "microsoft.intune.deviceenrollmentstatuspagewindows10"
        "microsoft.intune.endpointdetectionandresponsepolicywindows10"
        "microsoft.intune.exploitprotectionpolicywindows10settingcatalog"
        "microsoft.intune.policysets"
        "microsoft.intune.roleassignment"
        "microsoft.intune.roledefinition"
        "microsoft.intune.settingcatalogasrrulespolicywindows10"
        "microsoft.intune.settingcatalogcustompolicywindows10"
        "microsoft.intune.wificonfigurationpolicyandroiddeviceadministrator"
        "microsoft.intune.wificonfigurationpolicyandroidenterprisedeviceowner"
        "microsoft.intune.wificonfigurationpolicyandroidenterpriseworkprofile"
        "microsoft.intune.wificonfigurationpolicyandroidforwork"
        "microsoft.intune.wificonfigurationpolicyandroidopensourceproject"
        "microsoft.intune.wificonfigurationpolicyios"
        "microsoft.intune.wificonfigurationpolicymacos"
        "microsoft.intune.wificonfigurationpolicywindows10"
        "microsoft.intune.windowsautopilotdeploymentprofileazureadhybridjoined"
        "microsoft.intune.windowsautopilotdeploymentprofileazureadjoined"
        "microsoft.intune.windowsinformationprotectionpolicywindows10mdmenrolled"
        "microsoft.intune.windowsupdateforbusinessfeatureupdateprofilewindows10"
        "microsoft.intune.windowsupdateforbusinessringupdateprofilewindows10"
    )

    "Microsoft Teams" = @(
        "microsoft.teams.appPermissionPolicy"
        "microsoft.teams.appSetupPolicy"
        "microsoft.teams.audioConferencingPolicy"
        "microsoft.teams.callHoldPolicy"
        "microsoft.teams.callingPolicy"
        "microsoft.teams.callParkPolicy"
        "microsoft.teams.callQueue"
        "microsoft.teams.channelsPolicy"
        "microsoft.teams.clientConfiguration"
        "microsoft.teams.complianceRecordingPolicy"
        "microsoft.teams.cortanaPolicy"
        "microsoft.teams.dialInConferencingTenantSettings"
        "microsoft.teams.emergencyCallingPolicy"
        "microsoft.teams.emergencyCallRoutingPolicy"
        "microsoft.teams.enhancedEncryptionPolicy"
        "microsoft.teams.eventsPolicy"
        "microsoft.teams.federationConfiguration"
        "microsoft.teams.feedbackPolicy"
        "microsoft.teams.filesPolicy"
        "microsoft.teams.groupPolicyAssignment"
        "microsoft.teams.guestCallingConfiguration"
        "microsoft.teams.guestMeetingConfiguration"
        "microsoft.teams.guestMessagingConfiguration"
        "microsoft.teams.ipPhonePolicy"
        "microsoft.teams.meetingBroadcastConfiguration"
        "microsoft.teams.meetingBroadcastPolicy"
        "microsoft.teams.meetingConfiguration"
        "microsoft.teams.meetingPolicy"
        "microsoft.teams.messagingPolicy"
        "microsoft.teams.mobilityPolicy"
        "microsoft.teams.networkRoamingPolicy"
        "microsoft.teams.onlineVoicemailPolicy"
        "microsoft.teams.onlineVoicemailUserSettings"
        "microsoft.teams.onlineVoiceUser"
        "microsoft.teams.pstnUsage"
        "microsoft.teams.shiftsPolicy"
        "microsoft.teams.templatesPolicy"
        "microsoft.teams.tenantDialPlan"
        "microsoft.teams.tenantNetworkRegion"
        "microsoft.teams.tenantNetworkSite"
        "microsoft.teams.tenantNetworkSubnet"
        "microsoft.teams.tenantTrustedIPAddress"
        "microsoft.teams.translationRule"
        "microsoft.teams.unassignedNumberTreatment"
        "microsoft.teams.updateManagementPolicy"
        "microsoft.teams.upgradeConfiguration"
        "microsoft.teams.upgradePolicy"
        "microsoft.teams.user"
        "microsoft.teams.userCallingSettings"
        "microsoft.teams.userPolicyAssignment"
        "microsoft.teams.vdiPolicy"
        "microsoft.teams.voiceRoute"
        "microsoft.teams.voiceRoutingPolicy"
        "microsoft.teams.workloadPolicy"
    )

    "Security & Compliance" = @(
        "microsoft.securityandcompliance.autosensitivitylabelpolicy"
        "microsoft.securityandcompliance.caseholdpolicy"
        "microsoft.securityandcompliance.caseholdrule"
        "microsoft.securityandcompliance.compliancecase"
        "microsoft.securityandcompliance.compliancesearch"
        "microsoft.securityandcompliance.compliancesearchaction"
        "microsoft.securityandcompliance.compliancetag"
        "microsoft.securityandcompliance.deviceconditionalaccesspolicy"
        "microsoft.securityandcompliance.deviceconfigurationpolicy"
        "microsoft.securityandcompliance.dlpcompliancepolicy"
        "microsoft.securityandcompliance.fileplanpropertyauthority"
        "microsoft.securityandcompliance.fileplanpropertycategory"
        "microsoft.securityandcompliance.fileplanpropertycitation"
        "microsoft.securityandcompliance.fileplanpropertydepartment"
        "microsoft.securityandcompliance.fileplanpropertyreferenceid"
        "microsoft.securityandcompliance.fileplanpropertysubcategory"
        "microsoft.securityandcompliance.labelpolicy"
        "microsoft.securityandcompliance.protectionalert"
        "microsoft.securityandcompliance.retentioncompliancepolicy"
        "microsoft.securityandcompliance.retentioncompliancerule"
        "microsoft.securityandcompliance.retentioneventtype"
        "microsoft.securityandcompliance.securityfilter"
        "microsoft.securityandcompliance.supervisoryreviewpolicy"
        "microsoft.securityandcompliance.supervisoryreviewrule"
    )
}

# ============================================================================
# UI Helper Functions
# ============================================================================

function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   M365 Tenant Configuration Snapshot Tool" -ForegroundColor Cyan
    Write-Host "   Microsoft Graph Beta - UTCM APIs (Preview)" -ForegroundColor DarkCyan
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-MenuHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "  --- $Title ---" -ForegroundColor Yellow
    Write-Host ""
}

function Write-StatusLine {
    param(
        [string]$Label,
        [string]$Value,
        [ConsoleColor]$ValueColor = "White"
    )
    Write-Host "  $Label : " -NoNewline -ForegroundColor Gray
    Write-Host $Value -ForegroundColor $ValueColor
}

function Write-Success {
    param([string]$Message)
    Write-Host "  [OK] $Message" -ForegroundColor Green
}

function Write-Failure {
    param([string]$Message)
    Write-Host "  [ERROR] $Message" -ForegroundColor Red
}

function Write-Info {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor DarkGray
}

function Write-Waiting {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor DarkYellow
}

function Pause-ForKey {
    Write-Host ""
    Write-Host "  Press any key to continue..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Read-MenuChoice {
    param(
        [string]$Prompt = "Select an option",
        [string[]]$ValidChoices
    )
    Write-Host ""
    Write-Host "  $Prompt" -NoNewline -ForegroundColor White
    Write-Host ": " -NoNewline
    $choice = Read-Host
    if ($ValidChoices -and $choice -notin $ValidChoices) {
        Write-Failure "Invalid selection: '$choice'"
        return $null
    }
    return $choice
}

# ============================================================================
# Authentication Functions
# ============================================================================

function Connect-ToGraph {
    Write-Banner
    Write-MenuHeader "Connect to Microsoft Graph"

    $moduleName = "Microsoft.Graph.Authentication"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Waiting "Installing $moduleName module..."
        try {
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
            Write-Success "Module installed."
        }
        catch {
            Write-Failure "Failed to install $moduleName. $_"
            Pause-ForKey
            return
        }
    }

    Import-Module $moduleName -ErrorAction SilentlyContinue

    Write-Info "Connecting with scope: ConfigurationMonitoring.ReadWrite.All"
    Write-Host ""

    try {
        Connect-MgGraph -Scopes $script:RequiredScopes -NoWelcome
        $context = Get-MgContext
        if ($context) {
            $script:IsConnected = $true
            Write-Host ""
            Write-Success "Connected to Microsoft Graph"
            Write-StatusLine "Account" $context.Account
            Write-StatusLine "Tenant ID" $context.TenantId
            Write-StatusLine "Auth Type" $context.AuthType
        }
        else {
            Write-Failure "Connection returned no context."
        }
    }
    catch {
        Write-Failure "Failed to connect: $_"
    }

    Pause-ForKey
}

function Test-GraphConnection {
    if (-not $script:IsConnected) {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            $script:IsConnected = $true
            return $true
        }
        Write-Failure "Not connected to Microsoft Graph. Please connect first (Option 1)."
        Pause-ForKey
        return $false
    }
    return $true
}

# ============================================================================
# Graph API Request Helpers
# ============================================================================

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body
    )

    $params = @{
        Method = $Method
        Uri    = $Uri
    }

    if ($Body) {
        $params["Body"]        = $Body
        $params["ContentType"] = "application/json"
    }

    try {
        $response = Invoke-MgGraphRequest @params
        return $response
    }
    catch {
        $statusCode = $null
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }

        $errorMessage = $_.Exception.Message
        if ($_.ErrorDetails.Message) {
            try {
                $errorDetail = $_.ErrorDetails.Message | ConvertFrom-Json
                if ($errorDetail.error.message) {
                    $errorMessage = $errorDetail.error.message
                }
            }
            catch { }
        }

        Write-Failure "API Error ($statusCode): $errorMessage"
        return $null
    }
}

# ============================================================================
# Resource Type Selection Functions
# ============================================================================

function Show-WorkloadMenu {
    Write-MenuHeader "Select Workload"

    $workloadNames = @($script:Workloads.Keys)
    for ($i = 0; $i -lt $workloadNames.Count; $i++) {
        $count = $script:Workloads[$workloadNames[$i]].Count
        Write-Host "  [$($i + 1)] $($workloadNames[$i])" -NoNewline -ForegroundColor White
        Write-Host " ($count resource types)" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray

    $valid = @("0")
    for ($i = 1; $i -le $workloadNames.Count; $i++) { $valid += "$i" }

    $choice = Read-MenuChoice -Prompt "Select workload" -ValidChoices $valid
    if ($null -eq $choice -or $choice -eq "0") { return $null }

    return $workloadNames[[int]$choice - 1]
}

function Show-ResourceTypePicker {
    param(
        [Parameter(Mandatory)][string]$WorkloadName,
        [switch]$MultiSelect
    )

    $resources = $script:Workloads[$WorkloadName]

    Write-MenuHeader "Resource Types - $WorkloadName"
    Write-Info "Showing $($resources.Count) resource types."
    Write-Host ""

    # Display in pages of 20
    $pageSize = 20
    $totalPages = [Math]::Ceiling($resources.Count / $pageSize)
    $currentPage = 0
    $selected = @{}

    while ($true) {
        $startIdx = $currentPage * $pageSize
        $endIdx = [Math]::Min($startIdx + $pageSize - 1, $resources.Count - 1)

        Write-Host ""
        Write-Host "  Page $($currentPage + 1) of $totalPages" -ForegroundColor DarkCyan
        Write-Host "  -------------------------------------------------------" -ForegroundColor DarkGray

        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $num = $i + 1
            $shortName = $resources[$i] -replace "^microsoft\.\w+\.", ""
            $marker = if ($selected.ContainsKey($i)) { "[X]" } else { "[ ]" }
            $color = if ($selected.ContainsKey($i)) { "Green" } else { "White" }
            Write-Host "  $marker " -NoNewline -ForegroundColor $color
            Write-Host "$($num.ToString().PadLeft(3)). " -NoNewline -ForegroundColor DarkGray
            Write-Host "$shortName" -ForegroundColor $color
        }

        Write-Host ""
        Write-Host "  -------------------------------------------------------" -ForegroundColor DarkGray
        if ($MultiSelect) {
            Write-Host "  Commands:" -ForegroundColor DarkCyan
            Write-Host "    <number>      Toggle a resource type" -ForegroundColor Gray
            Write-Host "    <start-end>   Toggle a range (e.g. 1-5)" -ForegroundColor Gray
            Write-Host "    all           Select all in this workload" -ForegroundColor Gray
            Write-Host "    none          Deselect all" -ForegroundColor Gray
            Write-Host "    n / p         Next / Previous page" -ForegroundColor Gray
            Write-Host "    done          Confirm selection" -ForegroundColor Gray
            Write-Host "    cancel        Cancel" -ForegroundColor Gray
            Write-Host ""
            Write-Host "  Selected: $($selected.Count) resource type(s)" -ForegroundColor Cyan
        }
        else {
            Write-Host "  Enter a number to select, n/p to page, or cancel" -ForegroundColor Gray
        }

        $input_val = Read-MenuChoice -Prompt "Enter command"
        if ($null -eq $input_val) { $input_val = "" }
        $input_val = $input_val.Trim().ToLower()

        switch -Regex ($input_val) {
            "^cancel$" { return $null }
            "^done$" {
                if ($selected.Count -eq 0) {
                    Write-Failure "No resource types selected."
                    continue
                }
                return $selected.Keys | Sort-Object | ForEach-Object { $resources[$_] }
            }
            "^all$" {
                for ($i = 0; $i -lt $resources.Count; $i++) { $selected[$i] = $true }
                continue
            }
            "^none$" {
                $selected = @{}
                continue
            }
            "^n$" {
                if ($currentPage -lt $totalPages - 1) { $currentPage++ }
                continue
            }
            "^p$" {
                if ($currentPage -gt 0) { $currentPage-- }
                continue
            }
            "^(\d+)-(\d+)$" {
                $rangeStart = [int]$Matches[1]
                $rangeEnd = [int]$Matches[2]
                if ($rangeStart -ge 1 -and $rangeEnd -le $resources.Count -and $rangeStart -le $rangeEnd) {
                    for ($i = $rangeStart - 1; $i -le $rangeEnd - 1; $i++) {
                        if ($selected.ContainsKey($i)) { $selected.Remove($i) }
                        else { $selected[$i] = $true }
                    }
                }
                else {
                    Write-Failure "Invalid range. Valid: 1-$($resources.Count)"
                }
                continue
            }
            "^\d+$" {
                $num = [int]$input_val
                if ($num -ge 1 -and $num -le $resources.Count) {
                    $idx = $num - 1
                    if ($MultiSelect) {
                        if ($selected.ContainsKey($idx)) { $selected.Remove($idx) }
                        else { $selected[$idx] = $true }
                    }
                    else {
                        return @($resources[$idx])
                    }
                }
                else {
                    Write-Failure "Invalid number. Valid: 1-$($resources.Count)"
                }
                continue
            }
            default {
                Write-Failure "Unknown command: '$input_val'"
                continue
            }
        }
    }
}

# ============================================================================
# Snapshot Operations
# ============================================================================

function New-TenantSnapshot {
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [string]$Description = "",
        [Parameter(Mandatory)][string[]]$Resources
    )

    $body = @{
        displayName = $DisplayName
        description = $Description
        resources   = $Resources
    }

    Write-Info "Creating snapshot with $($Resources.Count) resource type(s)..."
    Write-Host ""

    $result = Invoke-GraphRequest -Method "POST" -Uri $script:CreateSnapshotEndpoint -Body $body

    if ($result) {
        Write-Success "Snapshot job created successfully!"
        Write-Host ""
        Write-StatusLine "Job ID" $result.id "Cyan"
        Write-StatusLine "Display Name" $result.displayName
        Write-StatusLine "Status" $result.status "Yellow"
        Write-StatusLine "Created" $result.createdDateTime
        Write-StatusLine "Resources" "$($Resources.Count) type(s)"
        Write-Host ""

        $poll = Read-MenuChoice -Prompt "Poll for completion? (y/n)" -ValidChoices @("y","n")
        if ($poll -eq "y") {
            Wait-ForSnapshotCompletion -JobId $result.id
        }
    }

    return $result
}

function Wait-ForSnapshotCompletion {
    param([Parameter(Mandatory)][string]$JobId)

    $maxAttempts = 60
    $delaySeconds = 5
    $attempt = 0

    Write-Host ""
    Write-Waiting "Polling for job completion (every ${delaySeconds}s, max ${maxAttempts} attempts)..."

    while ($attempt -lt $maxAttempts) {
        $attempt++
        Start-Sleep -Seconds $delaySeconds

        $job = Invoke-GraphRequest -Method "GET" -Uri "$script:SnapshotJobsEndpoint/$JobId"
        if (-not $job) {
            Write-Failure "Lost connection to job. Check status manually."
            return
        }

        $status = $job.status
        Write-Host "  [$attempt/$maxAttempts] Status: $status" -ForegroundColor DarkGray

        switch ($status) {
            "succeeded" {
                Write-Host ""
                Write-Success "Snapshot completed successfully!"
                Write-StatusLine "Completed" $job.completedDateTime
                if ($job.resourceLocation) {
                    Write-StatusLine "Resource Location" $job.resourceLocation "Cyan"
                }
                return
            }
            "failed" {
                Write-Host ""
                Write-Failure "Snapshot job failed."
                if ($job.errorDetails) {
                    Write-Host ""
                    Write-Host "  Error Details:" -ForegroundColor Red
                    foreach ($err in $job.errorDetails) {
                        Write-Host "    - $($err.code): $($err.message)" -ForegroundColor Red
                    }
                }
                return
            }
            "partiallySucceeded" {
                Write-Host ""
                Write-Host "  [WARN] Snapshot partially succeeded." -ForegroundColor Yellow
                Write-StatusLine "Completed" $job.completedDateTime
                if ($job.errorDetails) {
                    Write-Host ""
                    Write-Host "  Partial Errors:" -ForegroundColor Yellow
                    foreach ($err in $job.errorDetails) {
                        Write-Host "    - $($err.code): $($err.message)" -ForegroundColor Yellow
                    }
                }
                return
            }
        }
    }

    Write-Failure "Timed out waiting for snapshot completion after $($maxAttempts * $delaySeconds) seconds."
    Write-Info "Job ID: $JobId - check status manually."
}

# ============================================================================
# Menu Action: Create Snapshot by Selected Resource Types
# ============================================================================

function Invoke-CreateSnapshotByResourceTypes {
    if (-not (Test-GraphConnection)) { return }

    Write-Banner
    Write-MenuHeader "Create Snapshot - Select Resource Types"

    # Step 1: Pick a workload
    $workloadName = Show-WorkloadMenu
    if (-not $workloadName) { return }

    # Step 2: Pick resource types from that workload
    Write-Banner
    $selectedResources = Show-ResourceTypePicker -WorkloadName $workloadName -MultiSelect
    if (-not $selectedResources) { return }

    Write-Banner
    Write-MenuHeader "Create Snapshot"
    Write-Host "  Selected $($selectedResources.Count) resource type(s) from $workloadName" -ForegroundColor Cyan
    Write-Host ""

    foreach ($r in $selectedResources) {
        $short = $r -replace "^microsoft\.\w+\.", ""
        Write-Host "    - $short" -ForegroundColor Gray
    }

    Write-Host ""

    # Step 3: Get snapshot name
    Write-Host "  Snapshot display name" -NoNewline -ForegroundColor White
    Write-Host ": " -NoNewline
    $displayName = Read-Host
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = "$workloadName Snapshot - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    }

    Write-Host "  Description (optional)" -NoNewline -ForegroundColor White
    Write-Host ": " -NoNewline
    $description = Read-Host

    New-TenantSnapshot -DisplayName $displayName -Description $description -Resources $selectedResources
    Pause-ForKey
}

# ============================================================================
# Menu Action: Create Snapshot of Entire Workload
# ============================================================================

function Invoke-CreateSnapshotByWorkload {
    if (-not (Test-GraphConnection)) { return }

    Write-Banner
    Write-MenuHeader "Create Snapshot - Entire Workload"

    $workloadName = Show-WorkloadMenu
    if (-not $workloadName) { return }

    $resources = $script:Workloads[$workloadName]

    Write-Banner
    Write-MenuHeader "Confirm Workload Snapshot"
    Write-Host "  Workload   : $workloadName" -ForegroundColor Cyan
    Write-Host "  Resources  : $($resources.Count) resource type(s)" -ForegroundColor Cyan
    Write-Host ""

    $confirm = Read-MenuChoice -Prompt "Create snapshot for all $($resources.Count) resource types? (y/n)" -ValidChoices @("y","n")
    if ($confirm -ne "y") {
        Write-Info "Cancelled."
        Pause-ForKey
        return
    }

    $displayName = "$workloadName - Full Workload Snapshot - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

    Write-Host ""
    Write-Host "  Snapshot display name" -NoNewline -ForegroundColor White
    Write-Host " [$displayName]: " -NoNewline -ForegroundColor DarkGray
    $customName = Read-Host
    if (-not [string]::IsNullOrWhiteSpace($customName)) {
        $displayName = $customName
    }

    New-TenantSnapshot -DisplayName $displayName -Description "Full workload snapshot of $workloadName" -Resources $resources
    Pause-ForKey
}

# ============================================================================
# Menu Action: Create Snapshot of All Workloads
# ============================================================================

function Invoke-CreateSnapshotAllWorkloads {
    if (-not (Test-GraphConnection)) { return }

    Write-Banner
    Write-MenuHeader "Create Snapshot - All Workloads"

    $allResources = @()
    foreach ($workloadName in $script:Workloads.Keys) {
        $allResources += $script:Workloads[$workloadName]
    }

    Write-Host "  This will create a snapshot covering:" -ForegroundColor Cyan
    Write-Host ""
    foreach ($workloadName in $script:Workloads.Keys) {
        $count = $script:Workloads[$workloadName].Count
        Write-Host "    $workloadName" -NoNewline -ForegroundColor White
        Write-Host " ($count types)" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "  Total: $($allResources.Count) resource types" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "  [WARN] Large snapshots count against the 20,000 resource/month quota." -ForegroundColor Yellow
    Write-Host "  [WARN] Max 12 visible snapshot jobs allowed at a time." -ForegroundColor Yellow
    Write-Host ""

    $confirm = Read-MenuChoice -Prompt "Create snapshot for ALL workloads? (y/n)" -ValidChoices @("y","n")
    if ($confirm -ne "y") {
        Write-Info "Cancelled."
        Pause-ForKey
        return
    }

    $displayName = "All Workloads Snapshot - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

    Write-Host ""
    Write-Host "  Snapshot display name" -NoNewline -ForegroundColor White
    Write-Host " [$displayName]: " -NoNewline -ForegroundColor DarkGray
    $customName = Read-Host
    if (-not [string]::IsNullOrWhiteSpace($customName)) {
        $displayName = $customName
    }

    New-TenantSnapshot -DisplayName $displayName -Description "Complete tenant configuration snapshot across all workloads" -Resources $allResources
    Pause-ForKey
}

# ============================================================================
# Menu Action: List Snapshot Jobs
# ============================================================================

function Invoke-ListSnapshotJobs {
    if (-not (Test-GraphConnection)) { return }

    Write-Banner
    Write-MenuHeader "Snapshot Jobs"
    Write-Info "Fetching snapshot jobs..."
    Write-Host ""

    $result = Invoke-GraphRequest -Method "GET" -Uri $script:SnapshotJobsEndpoint
    if (-not $result) {
        Pause-ForKey
        return
    }

    $jobs = $result.value
    if (-not $jobs -or $jobs.Count -eq 0) {
        Write-Info "No snapshot jobs found."
        Pause-ForKey
        return
    }

    Write-Host "  Found $($jobs.Count) snapshot job(s):" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  {0,-4} {1,-40} {2,-20} {3,-22} {4}" -f "#", "Display Name", "Status", "Created", "Resources" -ForegroundColor DarkCyan
    Write-Host "  $('-' * 110)" -ForegroundColor DarkGray

    $jobIndex = 1
    foreach ($job in $jobs) {
        $statusColor = switch ($job.status) {
            "succeeded"          { "Green" }
            "failed"             { "Red" }
            "running"            { "Yellow" }
            "notStarted"         { "DarkYellow" }
            "partiallySucceeded" { "Yellow" }
            default              { "Gray" }
        }

        $name = if ($job.displayName.Length -gt 38) { $job.displayName.Substring(0, 35) + "..." } else { $job.displayName }
        $created = if ($job.createdDateTime) {
            try { ([datetime]$job.createdDateTime).ToString("yyyy-MM-dd HH:mm:ss") } catch { $job.createdDateTime }
        } else { "N/A" }
        $resourceCount = if ($job.resources) { "$($job.resources.Count) type(s)" } else { "N/A" }

        Write-Host "  $($jobIndex.ToString().PadLeft(3)). " -NoNewline -ForegroundColor DarkGray
        Write-Host "$($name.PadRight(40)) " -NoNewline -ForegroundColor White
        Write-Host "$($job.status.PadRight(20)) " -NoNewline -ForegroundColor $statusColor
        Write-Host "$($created.PadRight(22)) " -NoNewline -ForegroundColor Gray
        Write-Host "$resourceCount" -ForegroundColor Gray

        $jobIndex++
    }

    Write-Host ""
    Write-Host "  [HINT] Use 'View Snapshot Job Details' to see full info and download content." -ForegroundColor DarkGray
    Pause-ForKey
}

# ============================================================================
# Menu Action: View Snapshot Job Details
# ============================================================================

function Invoke-ViewSnapshotJobDetails {
    if (-not (Test-GraphConnection)) { return }

    Write-Banner
    Write-MenuHeader "View Snapshot Job Details"

    # Fetch jobs to let user pick
    Write-Info "Fetching snapshot jobs..."
    $result = Invoke-GraphRequest -Method "GET" -Uri $script:SnapshotJobsEndpoint
    if (-not $result) {
        Pause-ForKey
        return
    }

    $jobs = $result.value
    if (-not $jobs -or $jobs.Count -eq 0) {
        Write-Info "No snapshot jobs found."
        Pause-ForKey
        return
    }

    # List jobs for selection
    Write-Host ""
    for ($i = 0; $i -lt $jobs.Count; $i++) {
        $job = $jobs[$i]
        $statusColor = switch ($job.status) {
            "succeeded"          { "Green" }
            "failed"             { "Red" }
            "running"            { "Yellow" }
            "notStarted"         { "DarkYellow" }
            "partiallySucceeded" { "Yellow" }
            default              { "Gray" }
        }
        Write-Host "  [$($i + 1)] " -NoNewline -ForegroundColor DarkGray
        Write-Host "$($job.displayName)" -NoNewline -ForegroundColor White
        Write-Host " ($($job.status))" -ForegroundColor $statusColor
    }
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray

    $valid = @("0")
    for ($i = 1; $i -le $jobs.Count; $i++) { $valid += "$i" }

    $choice = Read-MenuChoice -Prompt "Select job to view" -ValidChoices $valid
    if ($null -eq $choice -or $choice -eq "0") { return }

    $selectedJob = $jobs[[int]$choice - 1]

    # Display full details
    Write-Banner
    Write-MenuHeader "Snapshot Job Details"

    Write-StatusLine "Job ID" $selectedJob.id "Cyan"
    Write-StatusLine "Display Name" $selectedJob.displayName
    Write-StatusLine "Description" $(if ($selectedJob.description) { $selectedJob.description } else { "(none)" })
    Write-StatusLine "Tenant ID" $selectedJob.tenantId

    $statusColor = switch ($selectedJob.status) {
        "succeeded"          { "Green" }
        "failed"             { "Red" }
        "running"            { "Yellow" }
        "notStarted"         { "DarkYellow" }
        "partiallySucceeded" { "Yellow" }
        default              { "Gray" }
    }
    Write-StatusLine "Status" $selectedJob.status $statusColor

    $createdStr = if ($selectedJob.createdDateTime) {
        try { ([datetime]$selectedJob.createdDateTime).ToString("yyyy-MM-dd HH:mm:ss UTC") } catch { $selectedJob.createdDateTime }
    } else { "N/A" }
    Write-StatusLine "Created" $createdStr

    $completedStr = if ($selectedJob.completedDateTime) {
        try { ([datetime]$selectedJob.completedDateTime).ToString("yyyy-MM-dd HH:mm:ss UTC") } catch { $selectedJob.completedDateTime }
    } else { "N/A" }
    Write-StatusLine "Completed" $completedStr

    if ($selectedJob.createdBy.user) {
        Write-StatusLine "Created By" "$($selectedJob.createdBy.user.displayName) ($($selectedJob.createdBy.user.id))"
    }

    # Resource types
    if ($selectedJob.resources -and $selectedJob.resources.Count -gt 0) {
        Write-Host ""
        Write-Host "  Resource Types ($($selectedJob.resources.Count)):" -ForegroundColor DarkCyan
        foreach ($r in $selectedJob.resources) {
            $short = $r -replace "^microsoft\.\w+\.", ""
            $workload = if ($r -match "^microsoft\.(\w+)\.") { $Matches[1] } else { "unknown" }
            Write-Host "    [$workload] $short" -ForegroundColor Gray
        }
    }

    # Error details
    if ($selectedJob.errorDetails -and $selectedJob.errorDetails.Count -gt 0) {
        Write-Host ""
        Write-Host "  Error Details:" -ForegroundColor Red
        foreach ($err in $selectedJob.errorDetails) {
            Write-Host "    Code    : $($err.code)" -ForegroundColor Red
            Write-Host "    Message : $($err.message)" -ForegroundColor Red
            Write-Host ""
        }
    }

    # Snapshot content download
    if ($selectedJob.status -eq "succeeded" -and $selectedJob.resourceLocation) {
        Write-Host ""
        $download = Read-MenuChoice -Prompt "Download snapshot content? (y/n)" -ValidChoices @("y","n")
        if ($download -eq "y") {
            Get-SnapshotContent -SnapshotId $selectedJob.resourceLocation -DisplayName $selectedJob.displayName
        }
    }
    elseif ($selectedJob.status -in @("notStarted", "running")) {
        Write-Host ""
        $poll = Read-MenuChoice -Prompt "Job is still in progress. Poll for completion? (y/n)" -ValidChoices @("y","n")
        if ($poll -eq "y") {
            Wait-ForSnapshotCompletion -JobId $selectedJob.id
        }
    }

    Pause-ForKey
}

# ============================================================================
# Snapshot Content Download
# ============================================================================

function Get-SnapshotContent {
    param(
        [Parameter(Mandatory)][string]$SnapshotId,
        [string]$DisplayName = "snapshot"
    )

    Write-Host ""
    Write-Waiting "Downloading snapshot content..."

    $snapshotUri = "$script:SnapshotsEndpoint('$SnapshotId')"
    $content = Invoke-GraphRequest -Method "GET" -Uri $snapshotUri

    if (-not $content) {
        Write-Failure "Failed to download snapshot content."
        return
    }

    Write-Success "Snapshot content retrieved!"
    Write-Host ""

    # Show summary of snapshot resources
    if ($content.resources) {
        Write-Host "  Snapshot contains $($content.resources.Count) resource instance(s):" -ForegroundColor Cyan
        Write-Host ""

        $grouped = $content.resources | Group-Object -Property resourceType
        foreach ($group in $grouped) {
            Write-Host "  $($group.Name)" -ForegroundColor DarkCyan
            Write-Host "    $($group.Count) instance(s)" -ForegroundColor Gray
            foreach ($item in $group.Group | Select-Object -First 5) {
                $itemName = if ($item.displayName) { $item.displayName } elseif ($item.properties.DisplayName) { $item.properties.DisplayName } else { "(unnamed)" }
                Write-Host "      - $itemName" -ForegroundColor DarkGray
            }
            if ($group.Count -gt 5) {
                Write-Host "      ... and $($group.Count - 5) more" -ForegroundColor DarkGray
            }
        }
    }

    # Offer to export
    Write-Host ""
    $export = Read-MenuChoice -Prompt "Export to JSON file? (y/n)" -ValidChoices @("y","n")
    if ($export -eq "y") {
        $safeName = ($DisplayName -replace '[\\/:*?"<>|]', '_')
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $defaultFile = "UTCM_Snapshot_${safeName}_${timestamp}.json"

        Write-Host "  Filename" -NoNewline -ForegroundColor White
        Write-Host " [$defaultFile]: " -NoNewline -ForegroundColor DarkGray
        $filename = Read-Host
        if ([string]::IsNullOrWhiteSpace($filename)) {
            $filename = $defaultFile
        }

        try {
            $content | ConvertTo-Json -Depth 20 | Out-File -FilePath $filename -Encoding UTF8
            $fullPath = (Resolve-Path $filename).Path
            Write-Success "Exported to: $fullPath"
        }
        catch {
            Write-Failure "Failed to export: $_"
        }
    }
}

# ============================================================================
# Menu Action: Delete Snapshot Job
# ============================================================================

function Invoke-DeleteSnapshotJob {
    if (-not (Test-GraphConnection)) { return }

    Write-Banner
    Write-MenuHeader "Delete Snapshot Job"

    Write-Info "Fetching snapshot jobs..."
    $result = Invoke-GraphRequest -Method "GET" -Uri $script:SnapshotJobsEndpoint
    if (-not $result) {
        Pause-ForKey
        return
    }

    $jobs = $result.value
    if (-not $jobs -or $jobs.Count -eq 0) {
        Write-Info "No snapshot jobs found."
        Pause-ForKey
        return
    }

    Write-Host ""
    for ($i = 0; $i -lt $jobs.Count; $i++) {
        $job = $jobs[$i]
        $statusColor = switch ($job.status) {
            "succeeded"          { "Green" }
            "failed"             { "Red" }
            "running"            { "Yellow" }
            "notStarted"         { "DarkYellow" }
            "partiallySucceeded" { "Yellow" }
            default              { "Gray" }
        }
        $created = if ($job.createdDateTime) {
            try { ([datetime]$job.createdDateTime).ToString("yyyy-MM-dd HH:mm") } catch { "" }
        } else { "" }
        Write-Host "  [$($i + 1)] " -NoNewline -ForegroundColor DarkGray
        Write-Host "$($job.displayName)" -NoNewline -ForegroundColor White
        Write-Host " | $($job.status)" -NoNewline -ForegroundColor $statusColor
        Write-Host " | $created" -ForegroundColor DarkGray
    }
    Write-Host "  [0] Cancel" -ForegroundColor DarkGray

    $valid = @("0")
    for ($i = 1; $i -le $jobs.Count; $i++) { $valid += "$i" }

    $choice = Read-MenuChoice -Prompt "Select job to delete" -ValidChoices $valid
    if ($null -eq $choice -or $choice -eq "0") { return }

    $selectedJob = $jobs[[int]$choice - 1]

    Write-Host ""
    Write-Host "  About to delete:" -ForegroundColor Red
    Write-StatusLine "Job ID" $selectedJob.id "Cyan"
    Write-StatusLine "Display Name" $selectedJob.displayName
    Write-StatusLine "Status" $selectedJob.status
    Write-Host ""

    $confirm = Read-MenuChoice -Prompt "Are you sure you want to delete this job? (yes/no)" -ValidChoices @("yes","no")
    if ($confirm -ne "yes") {
        Write-Info "Cancelled."
        Pause-ForKey
        return
    }

    $deleteResult = Invoke-GraphRequest -Method "DELETE" -Uri "$script:SnapshotJobsEndpoint/$($selectedJob.id)"

    # DELETE typically returns 204 No Content on success, Invoke-MgGraphRequest may return $null
    Write-Host ""
    Write-Success "Snapshot job '$($selectedJob.displayName)' deleted successfully."
    Pause-ForKey
}

# ============================================================================
# Menu Action: UTCM Setup Helper
# ============================================================================

function Invoke-UTCMSetup {
    Write-Banner
    Write-MenuHeader "UTCM Setup Helper"

    Write-Host "  This helper walks through the prerequisites for using UTCM APIs." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Prerequisites:" -ForegroundColor DarkCyan
    Write-Host "    1. Microsoft 365 E3 or Microsoft Entra ID P1 license" -ForegroundColor Gray
    Write-Host "    2. Microsoft.Graph.Authentication PowerShell module" -ForegroundColor Gray
    Write-Host "    3. UTCM Service Principal registered in your tenant" -ForegroundColor Gray
    Write-Host "    4. Workload-specific permissions granted to the UTCM SP" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  UTCM Service Principal App ID:" -ForegroundColor DarkCyan
    Write-Host "    $script:UTCMAppId" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "  -------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  [1] Check/Install Microsoft.Graph.Authentication module" -ForegroundColor White
    Write-Host "  [2] Register UTCM Service Principal in tenant" -ForegroundColor White
    Write-Host "  [3] Grant permissions to UTCM Service Principal" -ForegroundColor White
    Write-Host "  [4] Show permission reference table" -ForegroundColor White
    Write-Host "  [0] Back to main menu" -ForegroundColor DarkGray

    $choice = Read-MenuChoice -Prompt "Select an option" -ValidChoices @("0","1","2","3","4")
    if ($null -eq $choice -or $choice -eq "0") { return }

    switch ($choice) {
        "1" { Install-GraphModule }
        "2" { Register-UTCMServicePrincipal }
        "3" { Grant-UTCMPermissions }
        "4" { Show-PermissionReference }
    }

    Pause-ForKey
}

function Install-GraphModule {
    Write-Host ""
    $modules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications")
    foreach ($mod in $modules) {
        if (Get-Module -ListAvailable -Name $mod) {
            Write-Success "$mod is already installed."
        }
        else {
            Write-Waiting "Installing $mod..."
            try {
                Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
                Write-Success "$mod installed successfully."
            }
            catch {
                Write-Failure "Failed to install ${mod}: $_"
            }
        }
    }
}

function Register-UTCMServicePrincipal {
    if (-not (Test-GraphConnection)) { return }

    Write-Host ""
    Write-Waiting "Checking if UTCM Service Principal exists..."

    Import-Module Microsoft.Graph.Applications -ErrorAction SilentlyContinue

    try {
        $existing = Get-MgServicePrincipal -Filter "AppId eq '$script:UTCMAppId'" -ErrorAction Stop
        if ($existing) {
            Write-Success "UTCM Service Principal already registered."
            Write-StatusLine "Object ID" $existing.Id
            Write-StatusLine "Display Name" $existing.DisplayName
            return
        }
    }
    catch { }

    Write-Waiting "Registering UTCM Service Principal..."
    try {
        $sp = New-MgServicePrincipal -AppId $script:UTCMAppId
        Write-Success "UTCM Service Principal registered successfully!"
        Write-StatusLine "Object ID" $sp.Id
        Write-StatusLine "Display Name" $sp.DisplayName
    }
    catch {
        Write-Failure "Failed to register: $_"
    }
}

function Grant-UTCMPermissions {
    if (-not (Test-GraphConnection)) { return }

    Import-Module Microsoft.Graph.Applications -ErrorAction SilentlyContinue

    Write-Host ""
    Write-Host "  Select permissions to grant:" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "  Recommended permissions by workload:" -ForegroundColor Gray
    Write-Host "    [1] Entra ID         - Policy.Read.All, Directory.Read.All, User.Read.All" -ForegroundColor White
    Write-Host "    [2] Exchange Online   - Exchange.ManageAsApp" -ForegroundColor White
    Write-Host "    [3] Intune            - DeviceManagementConfiguration.Read.All" -ForegroundColor White
    Write-Host "    [4] Teams             - Policy.Read.All, Directory.Read.All" -ForegroundColor White
    Write-Host "    [5] Security/Purview  - InformationProtectionPolicy.Read.All" -ForegroundColor White
    Write-Host "    [6] All of the above" -ForegroundColor White
    Write-Host "    [7] Custom - enter permissions manually" -ForegroundColor White
    Write-Host "    [0] Cancel" -ForegroundColor DarkGray

    $choice = Read-MenuChoice -Prompt "Select option" -ValidChoices @("0","1","2","3","4","5","6","7")
    if ($null -eq $choice -or $choice -eq "0") { return }

    $permissions = switch ($choice) {
        "1" { @("Policy.Read.All", "Directory.Read.All", "User.Read.All", "Policy.Read.ConditionalAccess") }
        "2" { @("Exchange.ManageAsApp") }
        "3" { @("DeviceManagementConfiguration.Read.All", "DeviceManagementApps.Read.All") }
        "4" { @("Policy.Read.All", "Directory.Read.All") }
        "5" { @("InformationProtectionPolicy.Read.All") }
        "6" {
            @(
                "Policy.Read.All", "Directory.Read.All", "User.Read.All",
                "Policy.Read.ConditionalAccess",
                "DeviceManagementConfiguration.Read.All", "DeviceManagementApps.Read.All",
                "InformationProtectionPolicy.Read.All"
            )
        }
        "7" {
            Write-Host "  Enter permissions (comma-separated)" -NoNewline -ForegroundColor White
            Write-Host ": " -NoNewline
            $raw = Read-Host
            $raw -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
    }

    if (-not $permissions -or $permissions.Count -eq 0) {
        Write-Failure "No permissions specified."
        return
    }

    Write-Host ""
    Write-Waiting "Granting $($permissions.Count) permission(s) to UTCM Service Principal..."

    try {
        $graphSP = Get-MgServicePrincipal -Filter "AppId eq '$script:GraphAppId'"
        $utcmSP  = Get-MgServicePrincipal -Filter "AppId eq '$script:UTCMAppId'"

        if (-not $graphSP) { Write-Failure "Microsoft Graph SP not found."; return }
        if (-not $utcmSP)  { Write-Failure "UTCM SP not found. Register it first."; return }

        foreach ($perm in $permissions) {
            $appRole = $graphSP.AppRoles | Where-Object { $_.Value -eq $perm }
            if (-not $appRole) {
                # Try Exchange Online SP for Exchange.ManageAsApp
                if ($perm -eq "Exchange.ManageAsApp") {
                    $exoAppId = "00000002-0000-0ff1-ce00-000000000000"
                    $exoSP = Get-MgServicePrincipal -Filter "AppId eq '$exoAppId'"
                    if ($exoSP) {
                        $appRole = $exoSP.AppRoles | Where-Object { $_.Value -eq $perm }
                        if ($appRole) {
                            $body = @{
                                AppRoleId   = $appRole.Id
                                ResourceId  = $exoSP.Id
                                PrincipalId = $utcmSP.Id
                            }
                            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $utcmSP.Id -BodyParameter $body
                            Write-Success "Granted: $perm (Exchange Online)"
                            continue
                        }
                    }
                }
                Write-Host "  [SKIP] Permission '$perm' not found in app roles." -ForegroundColor Yellow
                continue
            }

            $body = @{
                AppRoleId   = $appRole.Id
                ResourceId  = $graphSP.Id
                PrincipalId = $utcmSP.Id
            }

            try {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $utcmSP.Id -BodyParameter $body
                Write-Success "Granted: $perm"
            }
            catch {
                if ($_.Exception.Message -match "already exists") {
                    Write-Info "  Already granted: $perm"
                }
                else {
                    Write-Failure "Failed to grant ${perm}: $_"
                }
            }
        }
    }
    catch {
        Write-Failure "Error: $_"
    }
}

function Show-PermissionReference {
    Write-Host ""
    Write-MenuHeader "UTCM Permission Reference"

    $table = @(
        [PSCustomObject]@{ Workload = "Entra ID";          Permissions = "Policy.Read.All, Directory.Read.All, User.Read.All, Policy.Read.ConditionalAccess" }
        [PSCustomObject]@{ Workload = "Exchange Online";    Permissions = "Exchange.ManageAsApp (on Office 365 Exchange Online SP)" }
        [PSCustomObject]@{ Workload = "Intune";             Permissions = "DeviceManagementConfiguration.Read.All, DeviceManagementApps.Read.All" }
        [PSCustomObject]@{ Workload = "Teams";              Permissions = "Policy.Read.All, Directory.Read.All (+ Teams Admin role may be required)" }
        [PSCustomObject]@{ Workload = "Security/Purview";   Permissions = "InformationProtectionPolicy.Read.All" }
    )

    Write-Host "  {0,-22} {1}" -f "Workload", "Required Permissions" -ForegroundColor DarkCyan
    Write-Host "  $('-' * 80)" -ForegroundColor DarkGray
    foreach ($row in $table) {
        Write-Host "  {0,-22} {1}" -f $row.Workload, $row.Permissions -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "  Caller Permission: ConfigurationMonitoring.ReadWrite.All" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "  API Limits (Preview):" -ForegroundColor DarkCyan
    Write-Host "    - Max 12 visible snapshot jobs (delete old ones to create more)" -ForegroundColor Gray
    Write-Host "    - 20,000 resource extractions per tenant per month" -ForegroundColor Gray
    Write-Host "    - Snapshots auto-deleted after 7 days" -ForegroundColor Gray
    Write-Host "    - Max 30 monitors per tenant" -ForegroundColor Gray
}

# ============================================================================
# Main Menu
# ============================================================================

function Show-MainMenu {
    while ($true) {
        Write-Banner

        # Connection status
        $connStatus = if ($script:IsConnected) { "Connected" } else { "Not connected" }
        $connColor  = if ($script:IsConnected) { "Green" } else { "Red" }
        Write-Host "  Connection Status: " -NoNewline -ForegroundColor Gray
        Write-Host $connStatus -ForegroundColor $connColor

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            $script:IsConnected = $true
            Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor DarkGray
        }

        Write-Host ""
        Write-Host "  -------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "  CONNECT" -ForegroundColor DarkCyan
        Write-Host "    [1] Connect to Microsoft Graph" -ForegroundColor White
        Write-Host ""
        Write-Host "  CREATE SNAPSHOTS" -ForegroundColor DarkCyan
        Write-Host "    [2] Create snapshot by selected resource types" -ForegroundColor White
        Write-Host "    [3] Create snapshot of entire workload" -ForegroundColor White
        Write-Host "    [4] Create snapshot of all workloads" -ForegroundColor White
        Write-Host ""
        Write-Host "  MANAGE SNAPSHOTS" -ForegroundColor DarkCyan
        Write-Host "    [5] List snapshot jobs" -ForegroundColor White
        Write-Host "    [6] View snapshot job details" -ForegroundColor White
        Write-Host "    [7] Delete snapshot job" -ForegroundColor White
        Write-Host ""
        Write-Host "  SETUP" -ForegroundColor DarkCyan
        Write-Host "    [8] UTCM setup helper" -ForegroundColor White
        Write-Host ""
        Write-Host "  -------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "    [0] Exit" -ForegroundColor DarkGray

        $choice = Read-MenuChoice -Prompt "Select an option" -ValidChoices @("0","1","2","3","4","5","6","7","8")

        switch ($choice) {
            "0" {
                Write-Host ""
                Write-Host "  Goodbye!" -ForegroundColor Cyan
                Write-Host ""
                return
            }
            "1" { Connect-ToGraph }
            "2" { Invoke-CreateSnapshotByResourceTypes }
            "3" { Invoke-CreateSnapshotByWorkload }
            "4" { Invoke-CreateSnapshotAllWorkloads }
            "5" { Invoke-ListSnapshotJobs }
            "6" { Invoke-ViewSnapshotJobDetails }
            "7" { Invoke-DeleteSnapshotJob }
            "8" { Invoke-UTCMSetup }
            $null { }
        }
    }
}

# ============================================================================
# Entry Point
# ============================================================================

Show-MainMenu
