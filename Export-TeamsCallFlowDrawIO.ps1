<#
.SYNOPSIS
    Export-TeamsCallFlowDrawIO.ps1
    Exports all Microsoft Teams Auto Attendants and Call Queues as draw.io diagrams.

.DESCRIPTION
    Connects to Microsoft Teams PowerShell, retrieves all Auto Attendants, Call Queues,
    and Resource Accounts, then generates draw.io (.drawio) XML files showing the complete
    call flow including business hours, after hours, holidays, menu options, Call Queue
    routing, timeout/overflow/no-agents handling, queue settings, nested Auto Attendants
    (linked to their own page in the combined file), TTS/audio greetings, and business
    hours schedules.

    If a Microsoft Graph session is also active (Connect-MgGraph -Scopes Group.Read.All),
    shared voicemail targets are labelled with the Microsoft 365 group's name.

    Each Auto Attendant gets its own .drawio file, plus a combined _AllCallFlows.drawio
    is created with one page per Auto Attendant and a Legend page.
    An _ExportSummary.json is also written with per-diagram statistics.

.PARAMETER OutputPath
    The folder where .drawio files will be saved. Defaults to .\CallFlowDiagrams

.PARAMETER StylePreset
    Colour palette to use for the diagrams.
    - Default      : Vibrant Microsoft-themed colours (original look).
    - HighContrast : Bold colours with thick borders for accessibility / projectors.
    - Monochrome   : Greyscale palette for black-and-white printing.
    Defaults to Default.

.PARAMETER HideQueueSettings
    Omit the per-queue settings note (agent sources, alert time, presence-based
    routing, opt-out, conference mode, callback, greeting, music on hold) for a
    more compact diagram.

.EXAMPLE
    .\Export-TeamsCallFlowDrawIO.ps1
    .\Export-TeamsCallFlowDrawIO.ps1 -OutputPath "C:\Customers\Contoso\CallFlows"
    .\Export-TeamsCallFlowDrawIO.ps1 -StylePreset HighContrast
    .\Export-TeamsCallFlowDrawIO.ps1 -StylePreset Monochrome -OutputPath .\Printable
    .\Export-TeamsCallFlowDrawIO.ps1 -HideQueueSettings

.NOTES
    Author  : Kieran Lloyd
    Version : 1.6
    Date    : 2026-09-30

    Prerequisites:
      - MicrosoftTeams PowerShell module installed
      - Connected to Teams via Connect-MicrosoftTeams
      - Optional: Microsoft.Graph module + Connect-MgGraph -Scopes Group.Read.All
        to show shared voicemail group names
      - Sufficient admin permissions to read AA/CQ/Resource Account configuration
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\CallFlowDiagrams",

    [Parameter(Mandatory = $false)]
    [ValidateSet("Default", "HighContrast", "Monochrome")]
    [string]$StylePreset = "Default",

    [Parameter(Mandatory = $false)]
    [switch]$HideQueueSettings
)

# ============================================================================
# CONFIGURATION
# ============================================================================
$ErrorActionPreference = "Continue"

# Shared voicemail group-name lookup via Microsoft Graph. Enabled in MAIN SCRIPT
# only when a Graph session is already connected; otherwise targets are labelled
# plain "Shared Voicemail". Script-scoped so Resolve-CallTarget doesn't need
# another parameter threaded through every graph-building function.
$script:GraphGroupLookup = $false
$script:GroupNameCache   = @{}

# ============================================================================
# NODE STYLE DEFINITIONS  (palette driven by -StylePreset)
# ============================================================================

# Each style table key maps to a node type.  Fill/stroke colours are the only
# values that differ between presets; shape geometry stays constant.
switch ($StylePreset) {
    "HighContrast" {
        $NodeStyles = @{
            "AA"              = "rounded=1;whiteSpace=wrap;html=1;fillColor=#0050EF;fontColor=#FFFFFF;strokeColor=#002080;strokeWidth=3;fontSize=12;fontFamily=Segoe UI;fontStyle=1;"
            "CQ"              = "rounded=1;whiteSpace=wrap;html=1;fillColor=#007A00;fontColor=#FFFFFF;strokeColor=#003D00;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;arcSize=50;"
            "Menu"            = "rhombus;whiteSpace=wrap;html=1;fillColor=#D79B00;fontColor=#000000;strokeColor=#6D4E00;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;"
            "MenuAfterHours"  = "rhombus;whiteSpace=wrap;html=1;fillColor=#0078D4;fontColor=#FFFFFF;strokeColor=#003060;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;"
            "User"            = "rounded=1;whiteSpace=wrap;html=1;fillColor=#5B0099;fontColor=#FFFFFF;strokeColor=#2B0050;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;"
            "ExternalPstn"    = "rounded=1;whiteSpace=wrap;html=1;fillColor=#C84B00;fontColor=#FFFFFF;strokeColor=#602400;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;"
            "SharedVoicemail" = "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;fillColor=#444444;fontColor=#FFFFFF;strokeColor=#000000;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;"
            "Disconnect"      = "ellipse;whiteSpace=wrap;html=1;fillColor=#9B0000;fontColor=#FFFFFF;strokeColor=#4B0000;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;fontStyle=1;"
            "Holiday"         = "shape=hexagon;perimeter=hexagonPerimeter2;whiteSpace=wrap;html=1;fillColor=#7A5000;fontColor=#FFFFFF;strokeColor=#3D2800;strokeWidth=3;fontSize=11;fontFamily=Segoe UI;size=0.25;"
            "TimeoutOverflow" = "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;fillColor=#C84B00;fontColor=#FFFFFF;strokeColor=#602400;strokeWidth=3;fontSize=10;fontFamily=Segoe UI;"
            "Title"           = "text;html=1;align=center;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=14;fontFamily=Segoe UI;fontStyle=1;fontColor=#000000;"
            "Greeting"        = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#FFFACD;strokeColor=#A0800A;strokeWidth=2;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
            "Schedule"        = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#C8E6FF;strokeColor=#004C99;strokeWidth=2;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
            "QueueSettings"   = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#CCF2CC;strokeColor=#003D00;strokeWidth=2;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
        }
        $EdgeStyles = @{
            "BusinessHours"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#333333;strokeWidth=3;fontFamily=Segoe UI;fontSize=10;"
            "AfterHours"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#0050EF;strokeWidth=3;fontFamily=Segoe UI;fontSize=10;"
            "Holiday"          = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#7A5000;strokeWidth=3;dashed=1;fontFamily=Segoe UI;fontSize=10;"
            "MenuOption"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
            "TimeoutOverflow"  = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#C84B00;strokeWidth=2;dashed=1;fontFamily=Segoe UI;fontSize=10;"
            "Greeting"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#A0800A;strokeWidth=2;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
            "Schedule"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#004C99;strokeWidth=2;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
            "QueueSettings"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#003D00;strokeWidth=2;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
        }
    }
    "Monochrome" {
        $NodeStyles = @{
            "AA"              = "rounded=1;whiteSpace=wrap;html=1;fillColor=#1A1A1A;fontColor=#FFFFFF;strokeColor=#000000;fontSize=12;fontFamily=Segoe UI;fontStyle=1;"
            "CQ"              = "rounded=1;whiteSpace=wrap;html=1;fillColor=#404040;fontColor=#FFFFFF;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;arcSize=50;"
            "Menu"            = "rhombus;whiteSpace=wrap;html=1;fillColor=#606060;fontColor=#FFFFFF;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;"
            "MenuAfterHours"  = "rhombus;whiteSpace=wrap;html=1;fillColor=#7A7A7A;fontColor=#FFFFFF;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;"
            "User"            = "rounded=1;whiteSpace=wrap;html=1;fillColor=#555555;fontColor=#FFFFFF;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;"
            "ExternalPstn"    = "rounded=1;whiteSpace=wrap;html=1;fillColor=#888888;fontColor=#FFFFFF;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;"
            "SharedVoicemail" = "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;fillColor=#AAAAAA;fontColor=#000000;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;"
            "Disconnect"      = "ellipse;whiteSpace=wrap;html=1;fillColor=#2A2A2A;fontColor=#FFFFFF;strokeColor=#000000;fontSize=11;fontFamily=Segoe UI;fontStyle=1;"
            "Holiday"         = "shape=hexagon;perimeter=hexagonPerimeter2;whiteSpace=wrap;html=1;fillColor=#B0B0B0;fontColor=#000000;strokeColor=#1A1A1A;fontSize=11;fontFamily=Segoe UI;size=0.25;"
            "TimeoutOverflow" = "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;fillColor=#CCCCCC;fontColor=#000000;strokeColor=#1A1A1A;fontSize=10;fontFamily=Segoe UI;"
            "Title"           = "text;html=1;align=center;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=14;fontFamily=Segoe UI;fontStyle=1;fontColor=#000000;"
            "Greeting"        = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#F0F0F0;strokeColor=#555555;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
            "Schedule"        = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#E0E0E0;strokeColor=#333333;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
            "QueueSettings"   = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#F7F7F7;strokeColor=#404040;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
        }
        $EdgeStyles = @{
            "BusinessHours"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#000000;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
            "AfterHours"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#555555;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
            "Holiday"          = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#777777;strokeWidth=2;dashed=1;fontFamily=Segoe UI;fontSize=10;"
            "MenuOption"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#333333;strokeWidth=1;fontFamily=Segoe UI;fontSize=10;"
            "TimeoutOverflow"  = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#888888;strokeWidth=1;dashed=1;fontFamily=Segoe UI;fontSize=10;"
            "Greeting"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#888888;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
            "Schedule"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#555555;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
            "QueueSettings"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#404040;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
        }
    }
    default {
        # Default — vibrant Microsoft-themed colours (original palette)
        $NodeStyles = @{
            "AA"              = "rounded=1;whiteSpace=wrap;html=1;fillColor=#4472C4;fontColor=#FFFFFF;strokeColor=#2F5496;fontSize=12;fontFamily=Segoe UI;fontStyle=1;"
            "CQ"              = "rounded=1;whiteSpace=wrap;html=1;fillColor=#548235;fontColor=#FFFFFF;strokeColor=#375623;fontSize=11;fontFamily=Segoe UI;arcSize=50;"
            "Menu"            = "rhombus;whiteSpace=wrap;html=1;fillColor=#FFC000;fontColor=#000000;strokeColor=#BF8F00;fontSize=11;fontFamily=Segoe UI;"
            "MenuAfterHours"  = "rhombus;whiteSpace=wrap;html=1;fillColor=#2E75B6;fontColor=#FFFFFF;strokeColor=#1F4E79;fontSize=11;fontFamily=Segoe UI;"
            "User"            = "rounded=1;whiteSpace=wrap;html=1;fillColor=#7030A0;fontColor=#FFFFFF;strokeColor=#4B1D6B;fontSize=11;fontFamily=Segoe UI;"
            "ExternalPstn"    = "rounded=1;whiteSpace=wrap;html=1;fillColor=#ED7D31;fontColor=#FFFFFF;strokeColor=#C55A11;fontSize=11;fontFamily=Segoe UI;"
            "SharedVoicemail" = "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;fillColor=#A5A5A5;fontColor=#FFFFFF;strokeColor=#7B7B7B;fontSize=11;fontFamily=Segoe UI;"
            "Disconnect"      = "ellipse;whiteSpace=wrap;html=1;fillColor=#C00000;fontColor=#FFFFFF;strokeColor=#8B0000;fontSize=11;fontFamily=Segoe UI;fontStyle=1;"
            "Holiday"         = "shape=hexagon;perimeter=hexagonPerimeter2;whiteSpace=wrap;html=1;fillColor=#BF8F00;fontColor=#FFFFFF;strokeColor=#8C6900;fontSize=11;fontFamily=Segoe UI;size=0.25;"
            "TimeoutOverflow" = "shape=parallelogram;perimeter=parallelogramPerimeter;whiteSpace=wrap;html=1;fillColor=#ED7D31;fontColor=#FFFFFF;strokeColor=#C55A11;fontSize=10;fontFamily=Segoe UI;"
            "Title"           = "text;html=1;align=center;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=14;fontFamily=Segoe UI;fontStyle=1;fontColor=#333333;"
            "Greeting"        = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#FFF2CC;strokeColor=#D6B656;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
            "Schedule"        = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#DAE8FC;strokeColor=#6C8EBF;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
            "QueueSettings"   = "shape=note;whiteSpace=wrap;html=1;backgroundOutline=1;fillColor=#E2EFDA;strokeColor=#548235;fontSize=9;fontFamily=Segoe UI;align=left;verticalAlign=top;spacingLeft=5;spacingRight=5;spacingTop=5;"
        }
        $EdgeStyles = @{
            "BusinessHours"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#666666;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
            "AfterHours"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#2E75B6;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
            "Holiday"          = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#BF8F00;strokeWidth=2;dashed=1;fontFamily=Segoe UI;fontSize=10;"
            "MenuOption"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#333333;strokeWidth=1;fontFamily=Segoe UI;fontSize=10;"
            "TimeoutOverflow"  = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#ED7D31;strokeWidth=1;dashed=1;fontFamily=Segoe UI;fontSize=10;"
            "Greeting"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#D6B656;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
            "Schedule"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#6C8EBF;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
            "QueueSettings"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#548235;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
        }
    }
}

$NodeSizes = @{
    "AA"              = @{ Width = 220; Height = 60 }
    "CQ"              = @{ Width = 240; Height = 60 }
    "Menu"            = @{ Width = 200; Height = 120 }
    "MenuAfterHours"  = @{ Width = 200; Height = 120 }
    "User"            = @{ Width = 180; Height = 50 }
    "ExternalPstn"    = @{ Width = 200; Height = 50 }
    "SharedVoicemail" = @{ Width = 200; Height = 50 }
    "Disconnect"      = @{ Width = 120; Height = 60 }
    "Holiday"         = @{ Width = 200; Height = 60 }
    "TimeoutOverflow" = @{ Width = 200; Height = 50 }
    "Title"           = @{ Width = 400; Height = 30 }
    "Greeting"        = @{ Width = 280; Height = 80 }
    "Schedule"        = @{ Width = 220; Height = 160 }
    "QueueSettings"   = @{ Width = 240; Height = 140 }
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Sanitise-NodeId {
<#
.SYNOPSIS
    Converts a string into a safe node ID by replacing non-alphanumeric characters.
#>
    param([string]$Text)
    $id = $Text -replace '[^a-zA-Z0-9]', '_'
    if ($id -match '^\d') { $id = "n$id" }
    return $id
}

function Escape-XmlString {
<#
.SYNOPSIS
    XML-encodes a string for safe use in XML attribute values.
#>
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&apos;')
}

function Get-DtmfDisplayKey {
<#
.SYNOPSIS
    Converts a DTMF tone enum to a display-friendly key label.
#>
    param([string]$DtmfResponse)
    switch ($DtmfResponse) {
        "Tone0"     { return "0" }
        "Tone1"     { return "1" }
        "Tone2"     { return "2" }
        "Tone3"     { return "3" }
        "Tone4"     { return "4" }
        "Tone5"     { return "5" }
        "Tone6"     { return "6" }
        "Tone7"     { return "7" }
        "Tone8"     { return "8" }
        "Tone9"     { return "9" }
        "ToneStar"  { return "*" }
        "TonePound" { return "#" }
        "Automatic" { return "Voice" }
        default     { return $DtmfResponse }
    }
}

function Get-CallFlowGreeting {
<#
.SYNOPSIS
    Extracts greeting text from a call flow's Greetings collection.
    Returns a hashtable with Type (TTS/Audio) and Text, or $null if no greeting.
#>
    param([object]$CallFlow)

    if ($null -eq $CallFlow.Greetings -or $CallFlow.Greetings.Count -eq 0) {
        return $null
    }

    $greeting = $CallFlow.Greetings | Select-Object -First 1

    if ($greeting.TextToSpeechPrompt) {
        return @{
            Type = "TTS"
            Text = $greeting.TextToSpeechPrompt
        }
    }
    elseif ($greeting.AudioFilePrompt) {
        $fileName = "Audio File"
        if ($greeting.AudioFilePrompt.FileName) {
            $fileName = $greeting.AudioFilePrompt.FileName
        }
        return @{
            Type = "Audio"
            Text = $fileName
        }
    }

    return $null
}

function Get-UserDisplayName {
<#
.SYNOPSIS
    Resolves a user ObjectId to a display name, cached in $UserCache.
#>
    param([string]$UserId, [hashtable]$UserCache)

    if ($UserCache.ContainsKey($UserId)) { return $UserCache[$UserId] }
    try {
        $displayName = (Get-CsOnlineUser -Identity $UserId -ErrorAction Stop).DisplayName
    } catch {
        $displayName = "User ($UserId)"
    }
    $UserCache[$UserId] = $displayName
    return $displayName
}

function Get-SharedVoicemailLabel {
<#
.SYNOPSIS
    Returns "Shared Voicemail: <group name>" when the Microsoft 365 group can be
    resolved through an existing Microsoft Graph session, else "Shared Voicemail".
.NOTES
    A permissions failure (no Group.Read.All consent) disables further lookups
    for the rest of the run, so one missing scope costs one failed call, not one
    per voicemail target.
#>
    param([string]$GroupId)

    if (-not $GroupId) { return "Shared Voicemail" }
    if (-not $script:GroupNameCache.ContainsKey($GroupId)) {
        if (-not $script:GraphGroupLookup) { return "Shared Voicemail" }
        $name = $null
        try {
            $name = (Get-MgGroup -GroupId $GroupId -Property DisplayName -ErrorAction Stop).DisplayName
        } catch {
            if ($_.Exception.Message -match 'Authorization|Forbidden|Insufficient privileges|Access.*denied') {
                Write-Host "       [i] Graph can't read groups (missing Group.Read.All?); shared voicemail names won't be shown." -ForegroundColor DarkYellow
                $script:GraphGroupLookup = $false
            }
        }
        $script:GroupNameCache[$GroupId] = $name
    }
    $groupName = $script:GroupNameCache[$GroupId]
    if ($groupName) { return "Shared Voicemail: $groupName" }
    return "Shared Voicemail"
}

function Get-QueueSettingsSummary {
<#
.SYNOPSIS
    Builds the HTML label for a Call Queue's settings note: agent sources, alert
    time, presence-based routing, opt-out, conference mode, callback, greeting,
    and music on hold. Returns already-escaped HTML.
#>
    param([object]$CallQueue)

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add("<b>&#x2699; Queue Settings</b>")

    if ($CallQueue.ChannelId) {
        $agentSource = "Teams channel"
    } else {
        $parts = @()
        $userCount  = if ($CallQueue.Users) { @($CallQueue.Users).Count } else { 0 }
        $groupCount = if ($CallQueue.DistributionLists) { @($CallQueue.DistributionLists).Count } else { 0 }
        if ($userCount -gt 0)  { $parts += "$userCount user$(if ($userCount -ne 1) { 's' })" }
        if ($groupCount -gt 0) { $parts += "$groupCount group$(if ($groupCount -ne 1) { 's' })" }
        $agentSource = if ($parts) { $parts -join ', ' } else { "None" }
    }
    [void]$lines.Add("Agents from: $agentSource")
    if ($null -ne $CallQueue.AgentAlertTime) { [void]$lines.Add("Agent alert time: $($CallQueue.AgentAlertTime)s") }
    [void]$lines.Add("Presence-based routing: $(if ($CallQueue.PresenceBasedRouting -eq $true) { 'On' } else { 'Off' })")
    [void]$lines.Add("Agent opt-out: $(if ($CallQueue.AllowOptOut -eq $true) { 'Allowed' } else { 'Not allowed' })")
    [void]$lines.Add("Conference mode: $(if ($CallQueue.ConferenceMode -eq $true) { 'On' } else { 'Off' })")

    if ($CallQueue.IsCallbackEnabled -eq $true) {
        $cb = "Callback: Press $(Get-DtmfDisplayKey "$($CallQueue.CallbackRequestDtmf)")"
        $conds = @()
        if ($CallQueue.WaitTimeBeforeOfferingCallbackInSecond)          { $conds += "after $($CallQueue.WaitTimeBeforeOfferingCallbackInSecond)s" }
        if ($CallQueue.NumberOfCallsInQueueBeforeOfferingCallback)      { $conds += "$($CallQueue.NumberOfCallsInQueueBeforeOfferingCallback)+ queued" }
        if ($CallQueue.CallToAgentRatioThresholdBeforeOfferingCallback) { $conds += "ratio $($CallQueue.CallToAgentRatioThresholdBeforeOfferingCallback)" }
        if ($conds) { $cb += " ($($conds -join ', '))" }
        [void]$lines.Add($cb)
    } else {
        [void]$lines.Add("Callback: Off")
    }

    if ($CallQueue.WelcomeTextToSpeechPrompt) {
        $tts = [string]$CallQueue.WelcomeTextToSpeechPrompt
        if ($tts.Length -gt 60) { $tts = $tts.Substring(0, 57) + "..." }
        [void]$lines.Add("Greeting: <i>&quot;$(Escape-XmlString $tts)&quot;</i>")
    } elseif ($CallQueue.WelcomeMusicAudioFileId) {
        [void]$lines.Add("Greeting: Audio file")
    } else {
        [void]$lines.Add("Greeting: None")
    }

    $moh = if ($CallQueue.UseDefaultMusicOnHold -eq $false -and $CallQueue.MusicOnHoldAudioFileId) { "Custom audio file" } else { "Default" }
    [void]$lines.Add("Music on hold: $moh")

    return ($lines -join "<br/>")
}

function ConvertTo-SafeDateTime {
<#
.SYNOPSIS
    Normalizes a schedule Start/End value to a [DateTime], regardless of the
    underlying CLR type. Different Teams tenants/module versions have been
    observed returning WeeklyRecurrentSchedule/FixedSchedule time values as
    [DateTime], [TimeSpan], or a plain date/time [string] rather than a
    reliable single type - calling .ToString(<format>) directly on whichever
    one shows up can throw (e.g. TimeSpan doesn't support date custom format
    strings). Returns $null if the value can't be interpreted as a date/time,
    so callers can fall back to a safe placeholder instead of crashing.
#>
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTime]) { return $Value }
    if ($Value -is [TimeSpan]) { return [DateTime]::Today.Add($Value) }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse([string]$Value, [ref]$parsed)) { return $parsed }

    return $null
}

function Get-BusinessHoursSchedule {
<#
.SYNOPSIS
    Extracts the business hours schedule from an Auto Attendant's AfterHours
    call handling association and formats it as an HTML string for display.
    Returns $null if no schedule is configured.
#>
    param([object]$AutoAttendant)

    if ($null -eq $AutoAttendant.CallHandlingAssociations) { return $null }
    $afterHoursAssoc = $AutoAttendant.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq "AfterHours" }
    if (-not $afterHoursAssoc) { return $null }

    $scheduleId = $afterHoursAssoc.ScheduleId
    if ($null -eq $AutoAttendant.Schedules) { return $null }
    $schedule = $AutoAttendant.Schedules | Where-Object { $_.Id -eq $scheduleId } | Select-Object -First 1
    if (-not $schedule -or -not $schedule.WeeklyRecurrentSchedule) { return $null }

    $wrs = $schedule.WeeklyRecurrentSchedule
    $days = @("Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday")
    $lines = [System.Collections.Generic.List[string]]::new()

    foreach ($day in $days) {
        $hoursProperty = "${day}Hours"
        $dayHours = $wrs.$hoursProperty

        $abbrev = $day.Substring(0, 3)

        if ($dayHours -and $dayHours.Count -gt 0) {
            $ranges = [System.Collections.Generic.List[string]]::new()
            foreach ($tr in $dayHours) {
                $startDt = ConvertTo-SafeDateTime $tr.Start
                $endDt   = ConvertTo-SafeDateTime $tr.End
                if ($startDt -and $endDt) {
                    [void]$ranges.Add("$($startDt.ToString('HH\:mm')) - $($endDt.ToString('HH\:mm'))")
                } else {
                    [void]$ranges.Add("$(Escape-XmlString $tr.Start) - $(Escape-XmlString $tr.End)")
                }
            }
            [void]$lines.Add("${abbrev}: $($ranges -join ', ')")
        }
        else {
            [void]$lines.Add("${abbrev}: Closed")
        }
    }

    return ($lines -join "<br/>")
}

function Get-HolidayScheduleDates {
<#
.SYNOPSIS
    Formats the fixed date ranges for the Holiday schedule(s) linked to a given
    Holiday call flow, so the diagram can show *when* a Holiday branch fires
    rather than just its name. Returns $null if no fixed date ranges are found.

    The result is deliberately kept compact - it becomes an edge label, so a
    schedule with many dates (e.g. a full year of national public holidays) is
    collapsed to "<name> (N dates)" rather than listing every one, otherwise the
    label stretches across the page and overlaps the title/root node. Schedules
    with 3 or fewer ranges are listed inline. Each schedule name is shown once,
    not repeated per range. Output uses <br/> between schedules and is
    pre-escaped for use inside an html=1 label.
.PARAMETER HolidayAssociations
    All of the Auto Attendant's Holiday-type CallHandlingAssociations. Multiple
    holiday schedules can route to the same CallFlowId (e.g. "Christmas Day" and
    "Boxing Day" both closing with the same message), so every association whose
    CallFlowId matches is included rather than just the first one encountered.
#>
    param([object]$AutoAttendant, [array]$HolidayAssociations, [string]$CallFlowId)

    if ($null -eq $AutoAttendant.Schedules -or $null -eq $HolidayAssociations) { return $null }

    $matchingAssocs = $HolidayAssociations | Where-Object { $_.CallFlowId -eq $CallFlowId }
    $lines = [System.Collections.Generic.List[string]]::new()
    $seenScheduleIds = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($ha in $matchingAssocs) {
        if (-not $seenScheduleIds.Add($ha.ScheduleId)) { continue }
        $schedule = $AutoAttendant.Schedules | Where-Object { $_.Id -eq $ha.ScheduleId } | Select-Object -First 1
        if (-not $schedule -or -not $schedule.FixedSchedule -or -not $schedule.FixedSchedule.DateTimeRanges) { continue }

        $scheduleName = $schedule.Name; if (-not $scheduleName) { $scheduleName = "Holiday" }

        $dateStrs = [System.Collections.Generic.List[string]]::new()
        foreach ($range in $schedule.FixedSchedule.DateTimeRanges) {
            $startDt = ConvertTo-SafeDateTime $range.Start
            $endDt   = ConvertTo-SafeDateTime $range.End
            if ($startDt -and $endDt) {
                # Teams stores the end boundary as exclusive - midnight at the
                # start of the day AFTER the final holiday day - so a single-day
                # holiday comes back as Start=day, End=next day (confirmed from
                # live data: New Year's Day exports as 01 Jan -> 02 Jan). Step the
                # end back by one tick so it lands on the real last holiday day,
                # so a one-day holiday reads as one date.
                $effEnd = $endDt.AddTicks(-1)
                if ($effEnd -lt $startDt) { $effEnd = $startDt }

                if ($startDt.Date -eq $effEnd.Date) {
                    [void]$dateStrs.Add($startDt.ToString("dd MMM yyyy"))
                } elseif ($startDt.Year -eq $effEnd.Year) {
                    [void]$dateStrs.Add("$($startDt.ToString('dd MMM')) - $($effEnd.ToString('dd MMM yyyy'))")
                } else {
                    [void]$dateStrs.Add("$($startDt.ToString('dd MMM yyyy')) - $($effEnd.ToString('dd MMM yyyy'))")
                }
            } else {
                [void]$dateStrs.Add("$($range.Start)")
            }
        }

        if ($dateStrs.Count -eq 0) { continue }

        if ($dateStrs.Count -le 3) {
            [void]$lines.Add("$(Escape-XmlString $scheduleName): $(Escape-XmlString ($dateStrs -join ', '))")
        } else {
            [void]$lines.Add("$(Escape-XmlString $scheduleName) ($($dateStrs.Count) dates)")
        }
    }

    if ($lines.Count -eq 0) { return $null }
    return ($lines -join "<br/>")
}

function Resolve-CallTarget {
<#
.SYNOPSIS
    Resolves a call target to a display name, type, and linked entity.
    Handles ApplicationEndpoint, ConfigurationEndpoint, User, ExternalPstn,
    SharedVoicemail, and DisconnectCall actions.
    Falls back to direct AA/CQ identity lookup if resource account not found.

    Shared voicemail is reported inconsistently by Teams: Auto Attendant menu
    options surface it as a target Type of "SharedVoicemail", but a Call Queue
    timeout/overflow surfaces it as an Action of "SharedVoicemail"/"Voicemail"
    with a target Type of "Mailbox". Both forms are mapped to SharedVoicemail so
    a CQ voicemail target doesn't fall through to "Unknown (MailBox)".

    A CQ action of "Voicemail" (as opposed to "SharedVoicemail") is a *user's*
    personal voicemail and resolves to Type "Voicemail" with the user's name.
.OUTPUTS
    Hashtable with keys: DisplayName, Type, LinkedId
#>
    param(
        [object]$CallTarget,
        [string]$Action,
        [hashtable]$ResourceAccountLookup,
        [hashtable]$AALookup,
        [hashtable]$CQLookup,
        [hashtable]$UserCache,
        [hashtable]$AAByAppInstance,
        [hashtable]$CQByAppInstance
    )

    if ($Action -eq "DisconnectCall") {
        return @{ DisplayName = "Disconnect"; Type = "Disconnect"; LinkedId = $null }
    }

    # Call Queue "Voicemail" action = a user's personal voicemail (target is
    # the user). Only treated as personal when the target really is a user, so
    # anything unexpected still falls through to the shared-voicemail handling.
    if ($Action -eq 'Voicemail' -and $CallTarget -and "$($CallTarget.Type)" -eq 'User') {
        $userName = Get-UserDisplayName -UserId $CallTarget.Id -UserCache $UserCache
        return @{ DisplayName = "Voicemail: $userName"; Type = "Voicemail"; LinkedId = $CallTarget.Id }
    }

    # Call Queue timeout/overflow to shared voicemail arrives as an Action of
    # "SharedVoicemail" rather than a distinct target Type, so key off the
    # action before inspecting the target type.
    if ($Action -match 'voicemail') {
        $vmId = if ($CallTarget) { $CallTarget.Id } else { $null }
        return @{ DisplayName = (Get-SharedVoicemailLabel -GroupId $vmId); Type = "SharedVoicemail"; LinkedId = $vmId }
    }

    if ($null -eq $CallTarget) {
        if ($Action -eq "TransferCallToOperator") {
            return @{ DisplayName = "Operator (none set on this Auto Attendant)"; Type = "Unknown"; LinkedId = $null }
        }
        return @{ DisplayName = "No Target Configured"; Type = "Unknown"; LinkedId = $null }
    }

    $targetId = $CallTarget.Id
    $targetType = $CallTarget.Type.ToString()

    switch ($targetType) {
        { $_ -in @("ApplicationEndpoint", "ConfigurationEndpoint") } {
            if ($ResourceAccountLookup.ContainsKey($targetId)) {
                $ra = $ResourceAccountLookup[$targetId]
                $appId = $ra.ApplicationId

                if ($appId -eq "ce933385-9390-45d1-9512-c8d228074e07") {
                    $linkedAA = $AAByAppInstance[$targetId]

                    if ($linkedAA) {
                        return @{ DisplayName = $linkedAA.Name; Type = "AA"; LinkedId = $linkedAA.Identity }
                    }
                    return @{ DisplayName = $ra.DisplayName; Type = "AA"; LinkedId = $null }
                }
                elseif ($appId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07") {
                    $linkedCQ = $CQByAppInstance[$targetId]

                    if ($linkedCQ) {
                        return @{ DisplayName = $linkedCQ.Name; Type = "CQ"; LinkedId = $linkedCQ.Identity }
                    }
                    return @{ DisplayName = $ra.DisplayName; Type = "CQ"; LinkedId = $null }
                }
                else {
                    return @{ DisplayName = $ra.DisplayName; Type = "Unknown"; LinkedId = $null }
                }
            }

            # Resource account not found — try direct AA/CQ identity lookup
            if ($AALookup.ContainsKey($targetId)) {
                return @{ DisplayName = $AALookup[$targetId].Name; Type = "AA"; LinkedId = $targetId }
            }
            if ($CQLookup.ContainsKey($targetId)) {
                return @{ DisplayName = $CQLookup[$targetId].Name; Type = "CQ"; LinkedId = $targetId }
            }

            # Try matching by ApplicationInstances across all AAs and CQs
            $matchedAA = $AAByAppInstance[$targetId]
            if ($matchedAA) {
                return @{ DisplayName = $matchedAA.Name; Type = "AA"; LinkedId = $matchedAA.Identity }
            }

            $matchedCQ = $CQByAppInstance[$targetId]
            if ($matchedCQ) {
                return @{ DisplayName = $matchedCQ.Name; Type = "CQ"; LinkedId = $matchedCQ.Identity }
            }

            return @{ DisplayName = "Unknown Voice App ($targetId)"; Type = "Unknown"; LinkedId = $null }
        }
        "User" {
            $displayName = Get-UserDisplayName -UserId $targetId -UserCache $UserCache
            return @{ DisplayName = $displayName; Type = "User"; LinkedId = $targetId }
        }
        "ExternalPstn" {
            $phoneNumber = $targetId -replace 'tel:', ''
            return @{ DisplayName = "External $phoneNumber"; Type = "ExternalPstn"; LinkedId = $null }
        }
        { $_ -in @("SharedVoicemail", "Mailbox") } {
            return @{ DisplayName = (Get-SharedVoicemailLabel -GroupId $targetId); Type = "SharedVoicemail"; LinkedId = $targetId }
        }
        default {
            return @{ DisplayName = "Unknown ($targetType)"; Type = "Unknown"; LinkedId = $null }
        }
    }
}

function Add-DiagramNode {
<#
.SYNOPSIS
    Creates a node hashtable and adds it to the node collection. Returns the assigned CellId.
.PARAMETER Link
    Optional draw.io link (e.g. "data:page/id,<pageId>") making the node clickable.
.PARAMETER Tooltip
    Optional hover text, shown alongside a Link.
#>
    param(
        [string]$NodeId, [string]$Label, [string]$Type,
        [int]$Tier, [int]$BranchIndex, [int]$PositionInBranch, [string]$ParentNodeId,
        [ref]$Nodes, [ref]$NodeMap, [ref]$NextCellId, [ref]$DefinedNodes,
        [string]$Link, [string]$Tooltip
    )

    if ($DefinedNodes.Value.Contains($NodeId)) {
        if ($NodeMap.Value.ContainsKey($NodeId)) { return $NodeMap.Value[$NodeId] }
        return -1
    }

    $cellId = $NextCellId.Value
    $NextCellId.Value++

    $style = $NodeStyles[$Type]
    if (-not $style) { $style = $NodeStyles["AA"] }
    $size = $NodeSizes[$Type]
    if (-not $size) { $size = @{ Width = 200; Height = 60 } }

    $node = @{
        CellId = $cellId; NodeId = $NodeId; Label = $Label; Type = $Type
        Style = $style; Width = $size.Width; Height = $size.Height
        Tier = $Tier; BranchIndex = $BranchIndex; PositionInBranch = $PositionInBranch
        ParentNodeId = $ParentNodeId; X = 0; Y = 0
        Link = $Link; Tooltip = $Tooltip
    }

    [void]$Nodes.Value.Add($node)
    $NodeMap.Value[$NodeId] = $cellId
    [void]$DefinedNodes.Value.Add($NodeId)
    return $cellId
}

function Add-DiagramEdge {
<#
.SYNOPSIS
    Creates an edge hashtable and adds it to the edge collection. Returns the assigned CellId.
.PARAMETER StyleSuffix
    Extra mxGraph style (e.g. fixed exit/entry points) appended to the StyleKey style.
.PARAMETER RouteDown
    Route the edge out of the source's bottom and into the target's top, with
    its horizontal jog just above the target (waypoints are computed from the
    laid-out positions in Build-DiagramXml). Used for queue exception edges so
    they run below the queue's settings note instead of through it.
#>
    param(
        [string]$SourceNodeId, [string]$TargetNodeId, [string]$Label, [string]$StyleKey,
        [ref]$Edges, [ref]$NextCellId,
        [string]$StyleSuffix, [switch]$RouteDown
    )

    $cellId = $NextCellId.Value
    $NextCellId.Value++

    $style = $EdgeStyles[$StyleKey]
    if (-not $style) { $style = $EdgeStyles["BusinessHours"] }
    if ($StyleSuffix) { $style += $StyleSuffix }

    $edge = @{
        CellId = $cellId; SourceNodeId = $SourceNodeId; TargetNodeId = $TargetNodeId
        Label = $Label; Style = $style; RouteDown = [bool]$RouteDown
    }

    [void]$Edges.Value.Add($edge)
    return $cellId
}

function Resolve-AndAddTargetNode {
<#
.SYNOPSIS
    Given a resolved call target (from Resolve-CallTarget), ensures the appropriate
    diagram node exists (recursing into Build-CallQueueNodes for CQ targets so their
    own timeout/overflow children are built) and returns its NodeId.
    Centralises target-type handling so menu options, default (no-menu) actions, and
    Call Queue timeout/overflow routing all resolve AA/CQ/User/ExternalPstn/
    SharedVoicemail/Disconnect/Unknown targets the same way.
.PARAMETER DisambiguationKey
    Base key used to build a unique NodeId for target types that have no natural
    global identity (ExternalPstn, SharedVoicemail, Disconnect, Unknown). Ignored
    for AA/CQ/User targets, which are keyed by their linked identity plus
    BranchIndex: the same target reached from two menu options *within* one
    branch still shares a single node (so convergence within a branch still
    reads naturally), but each branch (Business Hours/After Hours/Holiday) gets
    its own copy, so no edge ever has to cross between branches to reach a
    shared target.
.PARAMETER CurrentAAIdentity
    Identity of the Auto Attendant whose diagram is being built. A target that
    routes back to this same AA (e.g. a "press 9 for the main menu" loop) is
    kept unsuffixed so it resolves to the diagram's existing root node instead
    of drawing a confusing duplicate root per branch.
.OUTPUTS
    The NodeId of the resolved target, or "" if the target type could not be resolved.
#>
    param(
        [hashtable]$Target, [string]$DisambiguationKey,
        [int]$Tier, [int]$BranchIndex, [int]$PositionInBranch, [string]$ParentNodeId,
        [string]$CurrentAAIdentity,
        [hashtable]$ResourceAccountLookup, [hashtable]$AALookup,
        [hashtable]$CQLookup, [hashtable]$UserCache,
        [hashtable]$AAByAppInstance, [hashtable]$CQByAppInstance,
        [ref]$Nodes, [ref]$Edges, [ref]$NodeMap, [ref]$NextCellId, [ref]$DefinedNodes
    )

    $targetNodeId = ""

    switch ($Target.Type) {
        "AA" {
            if ($Target.LinkedId -and $Target.LinkedId -eq $CurrentAAIdentity) {
                # Loop back to this diagram's own root (e.g. "press 9 for the main menu")
                # rather than drawing a confusing duplicate root per branch.
                $targetNodeId = "AA_$(Sanitise-NodeId $Target.LinkedId)"
            } else {
                $targetNodeId = "AA_$(Sanitise-NodeId $Target.LinkedId)_b$BranchIndex"
            }
            if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                # A nested AA links to its own page in _AllCallFlows.drawio
                # (page ids are "page_<sanitised identity>", see MAIN SCRIPT).
                $link = ""; $tooltip = ""; $label = "<b>$(Escape-XmlString $Target.DisplayName)</b>"
                if ($Target.LinkedId -and $Target.LinkedId -ne $CurrentAAIdentity) {
                    $link = "data:page/id,page_$(Sanitise-NodeId $Target.LinkedId)"
                    $tooltip = "Open the $($Target.DisplayName) call flow (in _AllCallFlows.drawio)"
                    $label += "<br/><font style=&quot;font-size:9px&quot;>&#x2197; open call flow</font>"
                }
                Add-DiagramNode -NodeId $targetNodeId -Label $label -Type "AA" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes `
                    -Link $link -Tooltip $tooltip | Out-Null
            }
        }
        "CQ" {
            $targetNodeId = "CQ_$(Sanitise-NodeId $Target.LinkedId)_b$BranchIndex"
            if ($Target.LinkedId -and $CQLookup.ContainsKey($Target.LinkedId)) {
                Build-CallQueueNodes -CallQueue $CQLookup[$Target.LinkedId] -CQNodeId $targetNodeId `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -CurrentAAIdentity $CurrentAAIdentity `
                    -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                    -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
                    -Nodes $Nodes -Edges $Edges -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes
            } elseif (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                Add-DiagramNode -NodeId $targetNodeId -Label "<b>$(Escape-XmlString $Target.DisplayName)</b>" -Type "CQ" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
            }
        }
        "User" {
            $targetNodeId = "User_$(Sanitise-NodeId $Target.LinkedId)_b$BranchIndex"
            if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $Target.DisplayName) -Type "User" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
            }
        }
        "ExternalPstn" {
            $targetNodeId = "${DisambiguationKey}_ext"
            if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $Target.DisplayName) -Type "ExternalPstn" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
            }
        }
        { $_ -in @("SharedVoicemail", "Voicemail") } {
            # Shared (group) and personal (user) voicemail share the voicemail shape;
            # the label says which.
            $targetNodeId = "${DisambiguationKey}_vm"
            if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $Target.DisplayName) -Type "SharedVoicemail" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
            }
        }
        "Disconnect" {
            $targetNodeId = "${DisambiguationKey}_disc"
            if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                Add-DiagramNode -NodeId $targetNodeId -Label "Disconnect" -Type "Disconnect" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
            }
        }
        default {
            $targetNodeId = "${DisambiguationKey}_unk"
            if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $Target.DisplayName) -Type "AA" `
                    -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch -ParentNodeId $ParentNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
            }
        }
    }

    return $targetNodeId
}

function Build-CallQueueNodes {
<#
.SYNOPSIS
    Builds draw.io nodes for a Call Queue including agents info and its three
    exception-handling rules: timeout, overflow, and no-agents.
.NOTES
    Each exception becomes a child node suffixed _timeout / _overflow / _noagent
    (pattern-matched by Calculate-NodePositions). "Disconnect"/"DisconnectWithBusy"
    is drawn as a terminal Disconnect node; the no-agents rule is only drawn when
    it does something other than the default "Queue" (keep the call in the queue).
#>
    param(
        [object]$CallQueue, [string]$CQNodeId,
        [int]$Tier, [int]$BranchIndex, [int]$PositionInBranch, [string]$ParentNodeId,
        [string]$CurrentAAIdentity,
        [hashtable]$ResourceAccountLookup, [hashtable]$AALookup,
        [hashtable]$CQLookup, [hashtable]$UserCache,
        [hashtable]$AAByAppInstance, [hashtable]$CQByAppInstance,
        [ref]$Nodes, [ref]$Edges, [ref]$NodeMap, [ref]$NextCellId, [ref]$DefinedNodes
    )

    if ($DefinedNodes.Value.Contains($CQNodeId)) { return }

    $agentCount = 0
    if ($CallQueue.Agents) { $agentCount = $CallQueue.Agents.Count }
    $routingMethod = $CallQueue.RoutingMethod.ToString()
    $cqLabel = "<b>$(Escape-XmlString $CallQueue.Name)</b><br/>$agentCount Agents | $routingMethod"

    Add-DiagramNode -NodeId $CQNodeId -Label $cqLabel -Type "CQ" `
        -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch $PositionInBranch `
        -ParentNodeId $ParentNodeId `
        -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null

    # Settings note, drawn to the left of the queue (greetings go on the right)
    if (-not $HideQueueSettings) {
        $settingsNodeId = "${CQNodeId}_settings"
        Add-DiagramNode -NodeId $settingsNodeId -Label (Get-QueueSettingsSummary -CallQueue $CallQueue) -Type "QueueSettings" `
            -Tier $Tier -BranchIndex $BranchIndex -PositionInBranch 98 -ParentNodeId $CQNodeId `
            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        Add-DiagramEdge -SourceNodeId $CQNodeId -TargetNodeId $settingsNodeId `
            -Label "" -StyleKey "QueueSettings" -Edges $Edges -NextCellId $NextCellId `
            -StyleSuffix "exitX=0;exitY=0.5;exitDx=0;exitDy=0;entryX=1;entryY=0;entryDx=0;entryDy=$([int]($NodeSizes['CQ'].Height / 2));" | Out-Null
    }

    # Thresholds of 0 are valid (e.g. overflow every call immediately), so only
    # a missing value is shown as N/A.
    $timeoutThreshold  = if ($null -ne $CallQueue.TimeoutThreshold)  { "$($CallQueue.TimeoutThreshold)s" } else { "N/A" }
    $overflowThreshold = if ($null -ne $CallQueue.OverflowThreshold) { "$($CallQueue.OverflowThreshold) calls" } else { "N/A" }
    $noAgentScope = switch ("$($CallQueue.NoAgentApplyTo)") {
        "NewCalls" { "New calls" }
        default    { "All calls" }
    }

    $exceptions = @(
        @{ Suffix = "timeout";  EdgeLabel = "Timeout";   Caption = "Timeout ($timeoutThreshold)"
           Action = $CallQueue.TimeoutAction;  Target = $CallQueue.TimeoutActionTarget }
        @{ Suffix = "overflow"; EdgeLabel = "Overflow";  Caption = "Overflow ($overflowThreshold)"
           Action = $CallQueue.OverflowAction; Target = $CallQueue.OverflowActionTarget }
        @{ Suffix = "noagent";  EdgeLabel = "No Agents"; Caption = "No Agents ($noAgentScope)"
           Action = $CallQueue.NoAgentAction;  Target = $CallQueue.NoAgentActionTarget }
    )

    $childTier = $Tier + 1
    $childPos  = 0

    foreach ($ex in $exceptions) {
        if (-not $ex.Action) { continue }
        $actionStr = $ex.Action.ToString()
        if ($actionStr -eq "Queue") { continue }   # no-agents default: call just waits in queue

        $childNodeId = "${CQNodeId}_$($ex.Suffix)"

        if ($actionStr -in @("Disconnect", "DisconnectWithBusy")) {
            Add-DiagramNode -NodeId $childNodeId -Label "$($ex.Caption)<br/>Disconnect" -Type "Disconnect" `
                -Tier $childTier -BranchIndex $BranchIndex -PositionInBranch $childPos `
                -ParentNodeId $CQNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        } else {
            $target = Resolve-CallTarget -CallTarget $ex.Target `
                -Action $actionStr -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance
            Add-DiagramNode -NodeId $childNodeId -Label "$($ex.Caption)<br/>$(Escape-XmlString $target.DisplayName)" -Type "TimeoutOverflow" `
                -Tier $childTier -BranchIndex $BranchIndex -PositionInBranch $childPos `
                -ParentNodeId $CQNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null

            # Route downstream to target
            $downstreamNodeId = Resolve-AndAddTargetNode -Target $target -DisambiguationKey $childNodeId `
                -Tier ($childTier + 1) -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $childNodeId `
                -CurrentAAIdentity $CurrentAAIdentity `
                -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
                -Nodes $Nodes -Edges $Edges -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes
            if ($downstreamNodeId -and $DefinedNodes.Value.Contains($downstreamNodeId)) {
                Add-DiagramEdge -SourceNodeId $childNodeId -TargetNodeId $downstreamNodeId `
                    -Label "" -StyleKey "TimeoutOverflow" -Edges $Edges -NextCellId $NextCellId -RouteDown | Out-Null
            }
        }
        Add-DiagramEdge -SourceNodeId $CQNodeId -TargetNodeId $childNodeId `
            -Label $ex.EdgeLabel -StyleKey "TimeoutOverflow" -Edges $Edges -NextCellId $NextCellId -RouteDown | Out-Null
        $childPos++
    }
}

function Build-CallFlowNodes {
<#
.SYNOPSIS
    Processes a call flow (business hours, after hours, or holiday) and generates
    draw.io nodes and edges, including greeting notes where present.
.PARAMETER Operator
    The Auto Attendant's Operator callable entity. A "TransferCallToOperator"
    option usually has no CallTarget of its own - Teams stores the operator
    once on the AA - so it resolves to this instead.
#>
    param(
        [object]$CallFlow, [string]$ParentNodeId, [string]$FlowType,
        [string]$LinkLabel, [string]$AAIdentity, [int]$BranchIndex,
        [object]$Operator,
        [hashtable]$ResourceAccountLookup, [hashtable]$AALookup,
        [hashtable]$CQLookup, [hashtable]$UserCache,
        [hashtable]$AAByAppInstance, [hashtable]$CQByAppInstance,
        [ref]$Nodes, [ref]$Edges, [ref]$NodeMap, [ref]$NextCellId, [ref]$DefinedNodes
    )

    if ($null -eq $CallFlow) { return }

    $flowPrefix = Sanitise-NodeId "$($AAIdentity)_$($CallFlow.Id)"

    $menuType = "Menu"; $edgeStyleKey = "BusinessHours"
    switch ($FlowType) {
        "AfterHours" { $menuType = "MenuAfterHours"; $edgeStyleKey = "AfterHours" }
        "Holiday"    { $menuType = "Holiday"; $edgeStyleKey = "Holiday" }
    }

    $menu = $CallFlow.Menu
    if ($null -eq $menu) { return }

    $menuOptions = $menu.MenuOptions
    if ($null -eq $menuOptions) { return }
    # A flow whose only option is "Automatic" is a straight transfer/disconnect
    # with no IVR, so it takes the direct-transfer path below rather than being
    # drawn as a menu with a meaningless "Press Voice" connector. @() so .Count
    # is reliable for a single match on Windows PowerShell 5.1.
    $hasIvrOptions = @($menuOptions | Where-Object { $_.DtmfResponse.ToString() -ne "Automatic" })

    if ($hasIvrOptions.Count -gt 0) {
        $menuNodeId = "${flowPrefix}_menu"

        $menuLabel = switch ($FlowType) {
            "BusinessHours" { "<b>$(Escape-XmlString $CallFlow.Name)</b><br/>Business Hours Menu" }
            "AfterHours"    { "<b>$(Escape-XmlString $CallFlow.Name)</b><br/>After Hours Menu" }
            "Holiday"       { "<b>$(Escape-XmlString $CallFlow.Name)</b><br/>Holiday Menu" }
            default         { "<b>$(Escape-XmlString $CallFlow.Name)</b><br/>Menu" }
        }

        if (-not $DefinedNodes.Value.Contains($menuNodeId)) {
            Add-DiagramNode -NodeId $menuNodeId -Label $menuLabel -Type $menuType `
                -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        }

        Add-DiagramEdge -SourceNodeId $ParentNodeId -TargetNodeId $menuNodeId `
            -Label $LinkLabel -StyleKey $edgeStyleKey -Edges $Edges -NextCellId $NextCellId | Out-Null

        # Add greeting note if present
        $greetingInfo = Get-CallFlowGreeting -CallFlow $CallFlow
        if ($greetingInfo) {
            $greetingNodeId = "${flowPrefix}_greeting"
            if (-not $DefinedNodes.Value.Contains($greetingNodeId)) {
                $greetingPrefix = if ($greetingInfo.Type -eq "TTS") { "&#x1f50a; TTS Greeting:" } else { "&#x1f3b5; Audio File:" }
                $greetingText = $greetingInfo.Text
                if ($greetingText.Length -gt 200) { $greetingText = $greetingText.Substring(0, 197) + "..." }
                $greetingLabel = "<b>$greetingPrefix</b><br/><i>$(Escape-XmlString $greetingText)</i>"

                Add-DiagramNode -NodeId $greetingNodeId -Label $greetingLabel -Type "Greeting" `
                    -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 99 -ParentNodeId $menuNodeId `
                    -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                Add-DiagramEdge -SourceNodeId $menuNodeId -TargetNodeId $greetingNodeId `
                    -Label "" -StyleKey "Greeting" -Edges $Edges -NextCellId $NextCellId | Out-Null
            }
        }

        $optionIndex = 0
        $targetOptionLabels = [ordered]@{}   # targetNodeId -> List[string] of option labels
        foreach ($option in $menuOptions) {
            $dtmfKey = Get-DtmfDisplayKey $option.DtmfResponse.ToString()
            $action = $option.Action.ToString()
            $callTarget = $option.CallTarget
            if ($action -eq "TransferCallToOperator" -and -not $callTarget) { $callTarget = $Operator }

            $target = Resolve-CallTarget -CallTarget $callTarget -Action $action `
                -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance

            $disambigKey = "${flowPrefix}_$(Sanitise-NodeId $dtmfKey)"
            $targetNodeId = Resolve-AndAddTargetNode -Target $target -DisambiguationKey $disambigKey `
                -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                -CurrentAAIdentity $AAIdentity `
                -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
                -Nodes $Nodes -Edges $Edges -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes

            if ($targetNodeId) {
                $optionLabel = "Press $dtmfKey"
                if ($option.VoiceResponses) {
                    $voicePrompt = ($option.VoiceResponses | Select-Object -First 1)
                    if ($voicePrompt) { $optionLabel = "Press $dtmfKey / Say $(Escape-XmlString $voicePrompt)" }
                }
                if ($action -eq "TransferCallToOperator") { $optionLabel += " (Operator)" }
                if (-not $targetOptionLabels.Contains($targetNodeId)) {
                    $targetOptionLabels[$targetNodeId] = [System.Collections.Generic.List[string]]::new()
                }
                [void]$targetOptionLabels[$targetNodeId].Add($optionLabel)
            }
            $optionIndex++
        }

        # Emit one connector per distinct target. When several menu options
        # resolve to the same node (e.g. Press 2 and Press 3 both route to one
        # call queue), branch-scoped dedup collapses them onto a single node, so
        # drawing an edge per option stacks multiple connectors and labels on top
        # of each other. Combine the keys into a single edge ("Press 2, Press 3")
        # instead.
        foreach ($tid in $targetOptionLabels.Keys) {
            $combinedLabel = ($targetOptionLabels[$tid]) -join ', '
            Add-DiagramEdge -SourceNodeId $menuNodeId -TargetNodeId $tid `
                -Label $combinedLabel -StyleKey "MenuOption" -Edges $Edges -NextCellId $NextCellId | Out-Null
        }
    }
    else {
        # No IVR menu — direct transfer
        $defaultAction = $menu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -eq "Automatic" } | Select-Object -First 1
        if ($defaultAction) {
            $defaultActionStr = $defaultAction.Action.ToString()
            $callTarget = $defaultAction.CallTarget
            if ($defaultActionStr -eq "TransferCallToOperator" -and -not $callTarget) { $callTarget = $Operator }
            $target = Resolve-CallTarget -CallTarget $callTarget `
                -Action $defaultActionStr -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance

            $disambigKey = "${flowPrefix}_default"
            $targetNodeId = Resolve-AndAddTargetNode -Target $target -DisambiguationKey $disambigKey `
                -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                -CurrentAAIdentity $AAIdentity `
                -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
                -Nodes $Nodes -Edges $Edges -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes

            if ($targetNodeId) {
                $flowLabel = switch ($FlowType) {
                    "BusinessHours" { "Business Hours" }; "AfterHours" { "After Hours" }
                    "Holiday" { $CallFlow.Name }; default { "" }
                }
                if ($LinkLabel) { $flowLabel = $LinkLabel }
                Add-DiagramEdge -SourceNodeId $ParentNodeId -TargetNodeId $targetNodeId `
                    -Label $flowLabel -StyleKey $edgeStyleKey -Edges $Edges -NextCellId $NextCellId | Out-Null

                # Add greeting note (direct transfer flow)
                $greetingInfo = Get-CallFlowGreeting -CallFlow $CallFlow
                if ($greetingInfo) {
                    $greetingNodeId = "${flowPrefix}_greeting"
                    if (-not $DefinedNodes.Value.Contains($greetingNodeId)) {
                        $greetingPrefix = if ($greetingInfo.Type -eq "TTS") { "&#x1f50a; TTS Greeting:" } else { "&#x1f3b5; Audio File:" }
                        $greetingText = $greetingInfo.Text
                        if ($greetingText.Length -gt 200) { $greetingText = $greetingText.Substring(0, 197) + "..." }
                        $greetingLabel = "<b>$greetingPrefix</b><br/><i>$(Escape-XmlString $greetingText)</i>"
                        Add-DiagramNode -NodeId $greetingNodeId -Label $greetingLabel -Type "Greeting" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 99 -ParentNodeId $targetNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                        Add-DiagramEdge -SourceNodeId $targetNodeId -TargetNodeId $greetingNodeId `
                            -Label "" -StyleKey "Greeting" -Edges $Edges -NextCellId $NextCellId | Out-Null
                    }
                }
            }
        }
    }
}

function Test-DiagramIntegrity {
<#
.SYNOPSIS
    Validates a generated diagram's node/edge collections before writing to disk.
    Logs warnings for: edges referencing undefined nodes, duplicate cell IDs,
    and empty-label non-title nodes. Returns $true if the diagram is clean.
#>
    param(
        [System.Collections.Generic.List[hashtable]]$Nodes,
        [System.Collections.Generic.List[hashtable]]$Edges,
        [hashtable]$NodeMap,
        [string]$DiagramName
    )

    $isClean = $true
    $prefix  = "  [VALIDATE] $DiagramName :"

    # --- Duplicate cell IDs ---
    $cellIds = $Nodes | ForEach-Object { $_.CellId }
    $dupes   = $cellIds | Group-Object | Where-Object { $_.Count -gt 1 }
    foreach ($d in $dupes) {
        Write-Warning "$prefix Duplicate cell ID $($d.Name) found."
        $isClean = $false
    }

    # --- Edges referencing missing nodes ---
    # Look up membership via $NodeMap.ContainsKey directly, exactly as
    # Build-DiagramXml does when it resolves each edge. Casting
    # $NodeMap.Keys to [HashSet[string]] is NOT equivalent: on Windows
    # PowerShell 5.1 that cast collapses all keys into a single
    # space-joined string, so every real key then reports as missing and
    # every edge is falsely flagged as dangling even when the emitted XML
    # is completely valid.
    foreach ($edge in $Edges) {
        if (-not $NodeMap.ContainsKey($edge.SourceNodeId)) {
            Write-Warning "$prefix Edge references undefined source node '$($edge.SourceNodeId)'."
            $isClean = $false
        }
        if (-not $NodeMap.ContainsKey($edge.TargetNodeId)) {
            Write-Warning "$prefix Edge references undefined target node '$($edge.TargetNodeId)'."
            $isClean = $false
        }
    }

    # --- Empty-label nodes (excluding title and note nodes which may be intentionally blank) ---
    $nonNoteTypes = @("AA","CQ","Menu","MenuAfterHours","User","ExternalPstn","SharedVoicemail","Disconnect","Holiday","TimeoutOverflow")
    foreach ($node in ($Nodes | Where-Object { $_.Type -in $nonNoteTypes })) {
        if ([string]::IsNullOrWhiteSpace($node.Label)) {
            Write-Warning "$prefix Node '$($node.NodeId)' (type $($node.Type)) has an empty label."
            $isClean = $false
        }
    }

    return $isClean
}

function Calculate-NodePositions {
<#
.SYNOPSIS
    Calculates X/Y positions for all nodes using a tier-based top-down hierarchical layout.
    Prevents overlaps dynamically using boundary-aware footprints for all nodes.
    Returns a hashtable with PageWidth and PageHeight.
#>
    param([ref]$Nodes)

    $tierYPositions = @{ 0 = 60; 1 = 200; 2 = 380; 3 = 540 }
    $branchGap = 160; $siblingGap = 50; $subtreeGap = 30; $subtreeYGap = 80

    $tier0 = [System.Collections.Generic.List[hashtable]]::new()
    $tier1 = [System.Collections.Generic.List[hashtable]]::new()
    $tier2 = [System.Collections.Generic.List[hashtable]]::new()
    $titleNodes = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($node in $Nodes.Value) {
        if ($node.Type -eq "Title") { [void]$titleNodes.Add($node); continue }
        if ($node.Type -in @("Greeting","Schedule","QueueSettings")) { continue }
        switch ($node.Tier) {
            0 { [void]$tier0.Add($node) }
            1 { [void]$tier1.Add($node) }
            2 { [void]$tier2.Add($node) }
        }
    }

    # Index nodes by ID (and greetings by parent) once up front so the
    # repeated by-ID/by-parent lookups below are O(1) instead of scanning
    # the whole node list on every iteration.
    $nodeById = @{}
    foreach ($n in $Nodes.Value) { $nodeById[$n.NodeId] = $n }

    $greetingByParent = @{}
    $settingsByParent = @{}
    foreach ($n in $Nodes.Value) {
        if ($n.Type -eq "Greeting" -and -not $greetingByParent.ContainsKey($n.ParentNodeId)) {
            $greetingByParent[$n.ParentNodeId] = $n
        }
        if ($n.Type -eq "QueueSettings" -and -not $settingsByParent.ContainsKey($n.ParentNodeId)) {
            $settingsByParent[$n.ParentNodeId] = $n
        }
    }
    $noteGap = 15

    $tier1Branches = @{}
    foreach ($n in $tier1) {
        if (-not $tier1Branches.ContainsKey($n.BranchIndex)) {
            $tier1Branches[$n.BranchIndex] = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void]$tier1Branches[$n.BranchIndex].Add($n)
    }

    # "Hanging" nodes are everything below a Call Queue: its exception children
    # (_timeout/_overflow/_noagent) and whatever those route to, however deep
    # the chain of nested queues goes. They aren't laid out in tier rows; each
    # hangs below its parent as a subtree (see Measure-Subtree/Place-Subtree).
    $isHanging = @{}
    foreach ($n in $Nodes.Value) {
        if ($n.Type -in @("Title","Greeting","Schedule","QueueSettings")) { continue }
        if ($n.Tier -ge 3 -or $n.NodeId -match '_(timeout|overflow|noagent)$') { $isHanging[$n.NodeId] = $true }
    }

    $hangingChildren = @{}
    foreach ($n in $Nodes.Value) {
        if (-not $isHanging.ContainsKey($n.NodeId) -or -not $n.ParentNodeId) { continue }
        if (-not $hangingChildren.ContainsKey($n.ParentNodeId)) {
            $hangingChildren[$n.ParentNodeId] = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void]$hangingChildren[$n.ParentNodeId].Add($n)
    }

    # Subtree extent relative to the node's own X: Left <= 0, Right >= Width.
    # Includes the node's own notes (queue settings on the left, greeting on
    # the right). Children sit side by side (in PositionInBranch order:
    # timeout, overflow, no-agents) with each child's full subtree width
    # reserved, and the row is centred under the parent, so nested queues
    # never overlap each other.
    $subtreeExtent = @{}
    $childRelX = @{}
    function Measure-Subtree([hashtable]$node) {
        if ($subtreeExtent.ContainsKey($node.NodeId)) { return $subtreeExtent[$node.NodeId] }
        $baseLeft = 0; $baseRight = $node.Width
        $settingsNote = $settingsByParent[$node.NodeId]
        if ($settingsNote) { $baseLeft = -($noteGap + $settingsNote.Width) }
        $greetingNote = $greetingByParent[$node.NodeId]
        if ($greetingNote) { $baseRight = $node.Width + $noteGap + $greetingNote.Width }
        $subtreeExtent[$node.NodeId] = @{ Left = $baseLeft; Right = $baseRight }   # recursion guard

        $kids = @()
        if ($hangingChildren.ContainsKey($node.NodeId)) {
            $kids = @($hangingChildren[$node.NodeId] | Sort-Object { $_.PositionInBranch })
        }
        if ($kids.Count -eq 0) { return $subtreeExtent[$node.NodeId] }

        $kidExt = @{}
        foreach ($k in $kids) { $kidExt[$k.NodeId] = Measure-Subtree $k }

        $rel = @{}
        for ($i = 0; $i -lt $kids.Count; $i++) {
            $k = $kids[$i]
            if ($i -eq 0) {
                $rel[$k.NodeId] = -$kidExt[$k.NodeId].Left
            } else {
                $p = $kids[$i-1]
                $rel[$k.NodeId] = $rel[$p.NodeId] + $kidExt[$p.NodeId].Right + $subtreeGap - $kidExt[$k.NodeId].Left
            }
        }

        # Centre the row of child boxes (first to last) under the parent's centre
        $first = $kids[0]; $last = $kids[-1]
        $rowMid = (($rel[$first.NodeId] + $first.Width / 2) + ($rel[$last.NodeId] + $last.Width / 2)) / 2
        $shift = ($node.Width / 2) - $rowMid

        $left = $baseLeft; $right = $baseRight
        foreach ($k in $kids) {
            $rel[$k.NodeId] += $shift
            $childRelX[$k.NodeId] = $rel[$k.NodeId]
            $l = $rel[$k.NodeId] + $kidExt[$k.NodeId].Left
            $r = $rel[$k.NodeId] + $kidExt[$k.NodeId].Right
            if ($l -lt $left)  { $left = $l }
            if ($r -gt $right) { $right = $r }
        }
        $subtreeExtent[$node.NodeId] = @{ Left = $left; Right = $right }
        return $subtreeExtent[$node.NodeId]
    }

    $placedSubtree = @{}
    function Place-Subtree([hashtable]$node) {
        if ($placedSubtree.ContainsKey($node.NodeId)) { return }
        $placedSubtree[$node.NodeId] = $true
        if (-not $hangingChildren.ContainsKey($node.NodeId)) { return }
        # Drop below the settings note too, which is taller than the queue
        $childY = $node.Y + $node.Height + $subtreeYGap
        $settingsNote = $settingsByParent[$node.NodeId]
        if ($settingsNote) { $childY = [Math]::Max($childY, $node.Y + $settingsNote.Height + 40) }
        foreach ($k in $hangingChildren[$node.NodeId]) {
            $k.X = [int]($node.X + $childRelX[$k.NodeId])
            $k.Y = [int]$childY
            Place-Subtree $k
        }
    }

    $tier2Branches = @{}
    foreach ($n in $tier2) {
        if ($isHanging.ContainsKey($n.NodeId)) { continue }
        if (-not $tier2Branches.ContainsKey($n.BranchIndex)) {
            $tier2Branches[$n.BranchIndex] = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void]$tier2Branches[$n.BranchIndex].Add($n)
    }

    # 1. Calculate footprints for all non-note, non-child nodes
    # Footprint defines LeftOffset and RightOffset relative to the X coordinate.
    $footprints = @{}
    foreach ($n in $Nodes.Value) {
        if ($n.Type -in @("Greeting","Schedule","Title","QueueSettings") -or $isHanging.ContainsKey($n.NodeId)) {
            continue
        }

        # Reserve room for everything hanging below (CQ exception chains)
        $ext = Measure-Subtree $n
        $leftOffset = $ext.Left
        $rightOffset = $ext.Right

        # Check for greeting note
        $greeting = $greetingByParent[$n.NodeId]
        if ($greeting) {
            $rightOffset = [Math]::Max($rightOffset, $n.Width + 15 + $greeting.Width)
        }

        $footprints[$n.NodeId] = @{
            Left = $leftOffset
            Right = $rightOffset
        }
    }

    # 2. Local branch layout
    $branchWidths = @{}
    $branchLocalPositions = @{}
    $branchIndices = @($tier1Branches.Keys + $tier2Branches.Keys | Sort-Object -Unique)

    foreach ($bi in $branchIndices) {
        $t1Nodes = @()
        if ($tier1Branches.ContainsKey($bi)) { $t1Nodes = $tier1Branches[$bi] }
        $t2Nodes = @()
        if ($tier2Branches.ContainsKey($bi)) { $t2Nodes = $tier2Branches[$bi] }

        # Lay out Tier 1 locally
        $localT1 = @{}
        $w1 = 0
        if ($t1Nodes.Count -gt 0) {
            $sortedT1 = @($t1Nodes | Sort-Object { $_.PositionInBranch })
            for ($i = 0; $i -lt $sortedT1.Count; $i++) {
                $node = $sortedT1[$i]
                $fp = $footprints[$node.NodeId]
                if ($i -eq 0) {
                    $nodeX = -$fp.Left
                } else {
                    $prevNode = $sortedT1[$i-1]
                    $prevFp = $footprints[$prevNode.NodeId]
                    $nodeX = $localT1[$prevNode.NodeId] + $prevFp.Right + $siblingGap - $fp.Left
                }
                $localT1[$node.NodeId] = $nodeX
            }
            $w1 = $localT1[$sortedT1[-1].NodeId] + $footprints[$sortedT1[-1].NodeId].Right
        }

        # Lay out Tier 2 locally
        $localT2 = @{}
        $w2 = 0
        if ($t2Nodes.Count -gt 0) {
            $sortedT2 = @($t2Nodes | Sort-Object { $_.PositionInBranch })
            for ($i = 0; $i -lt $sortedT2.Count; $i++) {
                $node = $sortedT2[$i]
                $fp = $footprints[$node.NodeId]
                if ($i -eq 0) {
                    $nodeX = -$fp.Left
                } else {
                    $prevNode = $sortedT2[$i-1]
                    $prevFp = $footprints[$prevNode.NodeId]
                    $nodeX = $localT2[$prevNode.NodeId] + $prevFp.Right + $siblingGap - $fp.Left
                }
                $localT2[$node.NodeId] = $nodeX
            }
            $w2 = $localT2[$sortedT2[-1].NodeId] + $footprints[$sortedT2[-1].NodeId].Right
        }

        $branchW = [Math]::Max($w1, $w2)
        if ($branchW -lt 250) { $branchW = 250 }
        $branchWidths[$bi] = $branchW

        # Center tiers locally within branch width
        $t1Offset = ($branchW - $w1) / 2
        foreach ($node in $t1Nodes) {
            $branchLocalPositions[$node.NodeId] = $localT1[$node.NodeId] + $t1Offset
        }
        $t2Offset = ($branchW - $w2) / 2
        foreach ($node in $t2Nodes) {
            $branchLocalPositions[$node.NodeId] = $localT2[$node.NodeId] + $t2Offset
        }
    }

    # 3. Position branches sequentially on page
    $startX = 100
    $branchStartXMap = @{}
    $currentX = $startX
    foreach ($bi in $branchIndices) {
        $branchStartXMap[$bi] = $currentX
        $currentX += $branchWidths[$bi] + $branchGap
    }

    # Apply absolute coordinates to Tier 1 and Tier 2 nodes
    foreach ($bi in $branchIndices) {
        $t1Nodes = @()
        if ($tier1Branches.ContainsKey($bi)) { $t1Nodes = $tier1Branches[$bi] }
        foreach ($node in $t1Nodes) {
            $node.X = [int]($branchStartXMap[$bi] + $branchLocalPositions[$node.NodeId])
            $node.Y = $tierYPositions[1]
        }

        $t2Nodes = @()
        if ($tier2Branches.ContainsKey($bi)) { $t2Nodes = $tier2Branches[$bi] }
        foreach ($node in $t2Nodes) {
            $node.X = [int]($branchStartXMap[$bi] + $branchLocalPositions[$node.NodeId])
            $node.Y = $tierYPositions[2]
        }
    }

    # 4. Center root AA (Tier 0)
    $totalWidth = $currentX - $branchGap - $startX
    foreach ($n in $tier0) {
        $tcx = $startX + ($totalWidth / 2)
        $n.X = [int]($tcx - ($n.Width / 2))
        $n.Y = $tierYPositions[0]
    }

    # 5. Hang CQ exception subtrees below their (now positioned) roots
    foreach ($n in $Nodes.Value) {
        if ($n.Type -in @("Greeting","Schedule","Title","QueueSettings") -or $isHanging.ContainsKey($n.NodeId)) { continue }
        Place-Subtree $n
    }
    # Orphans (parent missing) - shouldn't happen, but keep them on the page
    foreach ($n in $Nodes.Value) {
        if ($isHanging.ContainsKey($n.NodeId) -and -not $placedSubtree.ContainsKey($n.NodeId)) {
            $n.X = $startX; $n.Y = $tierYPositions[3]
            Place-Subtree $n
        }
    }

    # 6. Position schedule note on the far-right sidebar (after subtrees, so it clears them)
    $scheduleNodes = $Nodes.Value | Where-Object { $_.Type -eq "Schedule" }
    if ($scheduleNodes.Count -gt 0) {
        $flowMaxX = 0
        foreach ($node in $Nodes.Value) {
            if ($node.Type -in @("Title","Schedule","Greeting","QueueSettings")) {
                continue
            }
            $r = $node.X + $node.Width

            # Include attached greeting note footprints
            $greeting = $greetingByParent[$node.NodeId]
            if ($greeting) {
                $r += 15 + $greeting.Width
            }
            
            if ($r -gt $flowMaxX) {
                $flowMaxX = $r
            }
        }
        if ($flowMaxX -lt 800) { $flowMaxX = 800 }

        foreach ($sn in $scheduleNodes) {
            $sn.X = $flowMaxX + 80
            $sn.Y = 60 # Align vertically with the root AA node
        }
    }

    # 7. Position title nodes
    foreach ($n in $titleNodes) {
        if ($tier0.Count -gt 0) { $n.X = $tier0[0].X - 90 } else { $n.X = $startX }
        $n.Y = 10
    }

    # 8. Position queue settings notes (left of their queue, top-aligned)
    foreach ($sn in $settingsByParent.Values) {
        $pn = $nodeById[$sn.ParentNodeId]
        if ($pn) {
            $sn.X = [int]($pn.X - $noteGap - $sn.Width)
            $sn.Y = $pn.Y
        }
    }

    # 8b. Position greeting notes
    $greetingNodes = $Nodes.Value | Where-Object { $_.Type -eq "Greeting" }
    foreach ($gn in $greetingNodes) {
        $pn = $nodeById[$gn.ParentNodeId]
        if ($pn) {
            $gn.X = $pn.X + $pn.Width + 15
            $gn.Y = [int]($pn.Y + 10)
        } else {
            $gn.X = 600
            $gn.Y = $tierYPositions[1]
        }
    }

    # 9. Compute page bounds
    $maxX = 0; $maxY = 0
    foreach ($n in $Nodes.Value) {
        $r = $n.X + $n.Width
        $b = $n.Y + $n.Height
        if ($r -gt $maxX) { $maxX = $r }
        if ($b -gt $maxY) { $maxY = $b }
    }

    return @{ PageWidth = [int][Math]::Max(1169, $maxX + 100); PageHeight = [int][Math]::Max(827, $maxY + 100) }
}

function Get-VertexXml {
<#
.SYNOPSIS
    Renders one vertex hashtable (CellId, Label, Style, X, Y, Width, Height and
    optional Link/Tooltip) as mxGraph XML. A linked vertex is wrapped in a
    <UserObject>, which is how draw.io stores links and tooltips; the
    UserObject carries the cell id so edges still resolve.
#>
    param([hashtable]$Node, [string]$Indent = "          ")

    $geom = "$Indent  <mxGeometry x=""$($Node.X)"" y=""$($Node.Y)"" width=""$($Node.Width)"" height=""$($Node.Height)"" as=""geometry""/>"
    $label = Escape-XmlString $Node.Label
    if ($Node.Link) {
        $tip = if ($Node.Tooltip) { " tooltip=""$(Escape-XmlString $Node.Tooltip)""" } else { "" }
        return @(
            "$Indent<UserObject label=""$label"" link=""$(Escape-XmlString $Node.Link)""$tip id=""$($Node.CellId)"">"
            "$Indent  <mxCell style=""$($Node.Style)"" vertex=""1"" parent=""1"">"
            "  $geom"
            "$Indent  </mxCell>"
            "$Indent</UserObject>"
        ) -join [Environment]::NewLine
    }
    return @(
        "$Indent<mxCell id=""$($Node.CellId)"" value=""$label"" style=""$($Node.Style)"" vertex=""1"" parent=""1"">"
        $geom
        "$Indent</mxCell>"
    ) -join [Environment]::NewLine
}

function Build-DiagramXml {
<#
.SYNOPSIS
    Generates the mxGraphModel XML string from node and edge collections.
#>
    param(
        [System.Collections.Generic.List[hashtable]]$Nodes,
        [System.Collections.Generic.List[hashtable]]$Edges,
        [hashtable]$NodeMap, [int]$PageWidth, [int]$PageHeight
    )

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("      <mxGraphModel dx=""1422"" dy=""762"" grid=""1"" gridSize=""10"" guides=""1"" tooltips=""1"" connect=""1"" arrows=""1"" fold=""1"" page=""1"" pageScale=""1"" pageWidth=""$PageWidth"" pageHeight=""$PageHeight"" math=""0"" shadow=""0"">")
    [void]$sb.AppendLine("        <root>")
    [void]$sb.AppendLine("          <mxCell id=""0""/>")
    [void]$sb.AppendLine("          <mxCell id=""1"" parent=""0""/>")

    foreach ($node in $Nodes) {
        [void]$sb.AppendLine((Get-VertexXml -Node $node))
    }

    $nodeById = @{}
    foreach ($n in $Nodes) { $nodeById[$n.NodeId] = $n }

    foreach ($edge in $Edges) {
        $sc = ""; $tc = ""
        if ($NodeMap.ContainsKey($edge.SourceNodeId)) { $sc = $NodeMap[$edge.SourceNodeId] }
        if ($NodeMap.ContainsKey($edge.TargetNodeId)) { $tc = $NodeMap[$edge.TargetNodeId] }
        if ($sc -and $tc) {
            $sl = Escape-XmlString $edge.Label
            $style = $edge.Style
            $points = $null

            # Bottom-out / top-in with the horizontal jog 20px above the target,
            # so the connector clears anything beside the source (the queue
            # settings note). Only when the target really is below the source;
            # a loop back up the diagram keeps draw.io's automatic routing.
            $src = $nodeById[$edge.SourceNodeId]; $tgt = $nodeById[$edge.TargetNodeId]
            if ($edge.RouteDown -and $src -and $tgt -and $tgt.Y -gt ($src.Y + $src.Height + 30)) {
                $style += "exitX=0.5;exitY=1;exitDx=0;exitDy=0;entryX=0.5;entryY=0;entryDx=0;entryDy=0;"
                $jogY  = $tgt.Y - 20
                $srcCx = [int]($src.X + $src.Width / 2)
                $tgtCx = [int]($tgt.X + $tgt.Width / 2)
                $points = "<Array as=""points""><mxPoint x=""$srcCx"" y=""$jogY""/><mxPoint x=""$tgtCx"" y=""$jogY""/></Array>"
            }

            [void]$sb.AppendLine("          <mxCell id=""$($edge.CellId)"" value=""$sl"" style=""$style"" edge=""1"" source=""$sc"" target=""$tc"" parent=""1"">")
            if ($points) {
                [void]$sb.AppendLine("            <mxGeometry relative=""1"" as=""geometry"">$points</mxGeometry>")
            } else {
                [void]$sb.AppendLine("            <mxGeometry relative=""1"" as=""geometry""/>")
            }
            [void]$sb.AppendLine("          </mxCell>")
        }
    }

    [void]$sb.AppendLine("        </root>")
    [void]$sb.AppendLine("      </mxGraphModel>")
    return $sb.ToString()
}

function Build-LegendPage {
<#
.SYNOPSIS
    Generates the XML for a legend page showing all node types and edge styles.
#>
    param([string]$DiagramId = "legend")

    $legendItems = @(
        @{ Label = "Auto Attendant"; Type = "AA" }
        @{ Label = "Call Queue"; Type = "CQ" }
        @{ Label = "Menu (Business Hours)"; Type = "Menu" }
        @{ Label = "Menu (After Hours)"; Type = "MenuAfterHours" }
        @{ Label = "User"; Type = "User" }
        @{ Label = "External PSTN"; Type = "ExternalPstn" }
        @{ Label = "Voicemail (Shared / Personal)"; Type = "SharedVoicemail" }
        @{ Label = "Disconnect"; Type = "Disconnect" }
        @{ Label = "Holiday"; Type = "Holiday" }
        @{ Label = "Timeout / Overflow / No Agents"; Type = "TimeoutOverflow" }
        @{ Label = "TTS / Audio Greeting"; Type = "Greeting" }
        @{ Label = "Business Hours Schedule"; Type = "Schedule" }
        @{ Label = "Queue Settings"; Type = "QueueSettings" }
    )

    $cellId = 2; $nodes = [System.Collections.Generic.List[hashtable]]::new(); $y = 80

    [void]$nodes.Add(@{ CellId = $cellId; Label = "<b>Call Flow Diagram Legend</b>"; Style = $NodeStyles["Title"]; X = 50; Y = 20; Width = 400; Height = 40 })
    $cellId++

    foreach ($item in $legendItems) {
        $style = $NodeStyles[$item.Type]; $size = $NodeSizes[$item.Type]
        [void]$nodes.Add(@{ CellId = $cellId; Label = $item.Label; Style = $style; X = 80; Y = $y; Width = $size.Width; Height = $size.Height })
        $cellId++; $y += $size.Height + 20
    }

    $y += 20
    [void]$nodes.Add(@{ CellId = $cellId; Label = "<b>Edge Styles</b>"; Style = "text;html=1;align=left;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=12;fontFamily=Segoe UI;fontStyle=1;fontColor=#333333;"; X = 50; Y = $y; Width = 300; Height = 30 })
    $cellId++; $y += 40

    $edgeLegendItems = @(
        @{ Label = "Business Hours"; Color = "#666666"; Dashed = "0" }
        @{ Label = "After Hours"; Color = "#2E75B6"; Dashed = "0" }
        @{ Label = "Holiday"; Color = "#BF8F00"; Dashed = "1" }
        @{ Label = "Menu Option"; Color = "#333333"; Dashed = "0" }
        @{ Label = "Timeout / Overflow / No Agents"; Color = "#ED7D31"; Dashed = "1" }
        @{ Label = "Greeting"; Color = "#D6B656"; Dashed = "1" }
        @{ Label = "Schedule"; Color = "#6C8EBF"; Dashed = "1" }
    )

    foreach ($ei in $edgeLegendItems) {
        $srcId = $cellId; $cellId++
        [void]$nodes.Add(@{ CellId = $srcId; Label = ""; Style = "ellipse;whiteSpace=wrap;html=1;fillColor=$($ei.Color);strokeColor=$($ei.Color);fontSize=8;"; X = 80; Y = ($y + 10); Width = 20; Height = 20 })
        $tgtId = $cellId; $cellId++
        [void]$nodes.Add(@{ CellId = $tgtId; Label = ""; Style = "ellipse;whiteSpace=wrap;html=1;fillColor=$($ei.Color);strokeColor=$($ei.Color);fontSize=8;"; X = 250; Y = ($y + 10); Width = 20; Height = 20 })
        $lblId = $cellId; $cellId++
        [void]$nodes.Add(@{ CellId = $lblId; Label = $ei.Label; Style = "text;html=1;align=left;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=11;fontFamily=Segoe UI;fontColor=#333333;"; X = 290; Y = $y; Width = 200; Height = 40 })
        $ei.SrcId = $srcId
        $ei.TgtId = $tgtId
        $y += 50
    }

    $pageHeight = [Math]::Max(827, $y + 50)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("  <diagram id=""$DiagramId"" name=""Legend"">")
    [void]$sb.AppendLine("      <mxGraphModel dx=""1422"" dy=""762"" grid=""1"" gridSize=""10"" guides=""1"" tooltips=""1"" connect=""1"" arrows=""1"" fold=""1"" page=""1"" pageScale=""1"" pageWidth=""600"" pageHeight=""$pageHeight"" math=""0"" shadow=""0"">")
    [void]$sb.AppendLine("        <root>")
    [void]$sb.AppendLine("          <mxCell id=""0""/>")
    [void]$sb.AppendLine("          <mxCell id=""1"" parent=""0""/>")

    foreach ($node in $nodes) {
        [void]$sb.AppendLine("          <mxCell id=""$($node.CellId)"" value=""$(Escape-XmlString $node.Label)"" style=""$($node.Style)"" vertex=""1"" parent=""1"">")
        [void]$sb.AppendLine("            <mxGeometry x=""$($node.X)"" y=""$($node.Y)"" width=""$($node.Width)"" height=""$($node.Height)"" as=""geometry""/>")
        [void]$sb.AppendLine("          </mxCell>")
    }

    foreach ($ei in $edgeLegendItems) {
        $edgeId = $cellId; $cellId++
        $ds = ""; if ($ei.Dashed -eq "1") { $ds = "dashed=1;" }
        [void]$sb.AppendLine("          <mxCell id=""$edgeId"" value="""" style=""edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=$($ei.Color);strokeWidth=2;${ds}fontFamily=Segoe UI;fontSize=10;"" edge=""1"" source=""$($ei.SrcId)"" target=""$($ei.TgtId)"" parent=""1"">")
        [void]$sb.AppendLine("            <mxGeometry relative=""1"" as=""geometry""/>")
        [void]$sb.AppendLine("          </mxCell>")
    }

    [void]$sb.AppendLine("        </root>")
    [void]$sb.AppendLine("      </mxGraphModel>")
    [void]$sb.AppendLine("  </diagram>")
    return $sb.ToString()
}

function Build-IndexPage {
<#
.SYNOPSIS
    Generates the XML for a cover/index page listing every exported Auto
    Attendant with its phone number(s), sorted alphabetically, so a combined
    file with dozens of pages has a starting point instead of a stack of
    unrelated diagrams.
.PARAMETER Diagrams
    The list of per-AA result hashtables (from Export-AADiagram), each
    expected to have Name, PhoneNumbers and PageId. Each row links to its page.
#>
    param([string]$DiagramId = "index_page", [array]$Diagrams)

    $cellId = 2
    $nodes = [System.Collections.Generic.List[hashtable]]::new()
    $y = 90

    [void]$nodes.Add(@{ CellId = $cellId; Label = "<b>Auto Attendant Directory</b>"; Style = $NodeStyles["Title"]; X = 50; Y = 20; Width = 500; Height = 40 })
    $cellId++

    $countLabel = "$($Diagrams.Count) Auto Attendant$(if ($Diagrams.Count -ne 1) { 's' })"
    $subtitleStyle = "text;html=1;align=left;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=11;fontFamily=Segoe UI;fontColor=#666666;"
    [void]$nodes.Add(@{ CellId = $cellId; Label = $countLabel; Style = $subtitleStyle; X = 50; Y = 60; Width = 300; Height = 20 })
    $cellId++

    $rowStyle = "text;html=1;align=left;verticalAlign=middle;resizable=0;points=[];autosize=1;strokeColor=none;fillColor=none;fontSize=12;fontFamily=Segoe UI;fontColor=#333333;"
    $sortedDiagrams = $Diagrams | Sort-Object { $_.Name }

    foreach ($d in $sortedDiagrams) {
        $phoneText = "No number assigned"
        if ($d.PhoneNumbers -and $d.PhoneNumbers.Count -gt 0) { $phoneText = $d.PhoneNumbers -join ', ' }
        $rowLabel = "<b>$(Escape-XmlString $d.Name)</b>&#160;&#8212;&#160;$(Escape-XmlString $phoneText)"
        [void]$nodes.Add(@{ CellId = $cellId; Label = $rowLabel; Style = $rowStyle; X = 80; Y = $y; Width = 760; Height = 26
            Link = "data:page/id,$($d.PageId)"; Tooltip = "Open $($d.Name)" })
        $cellId++
        $y += 30
    }

    $pageHeight = [Math]::Max(827, $y + 50)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("  <diagram id=""$DiagramId"" name=""Index"">")
    [void]$sb.AppendLine("      <mxGraphModel dx=""1422"" dy=""762"" grid=""1"" gridSize=""10"" guides=""1"" tooltips=""1"" connect=""1"" arrows=""1"" fold=""1"" page=""1"" pageScale=""1"" pageWidth=""900"" pageHeight=""$pageHeight"" math=""0"" shadow=""0"">")
    [void]$sb.AppendLine("        <root>")
    [void]$sb.AppendLine("          <mxCell id=""0""/>")
    [void]$sb.AppendLine("          <mxCell id=""1"" parent=""0""/>")

    foreach ($node in $nodes) {
        [void]$sb.AppendLine((Get-VertexXml -Node $node))
    }

    [void]$sb.AppendLine("        </root>")
    [void]$sb.AppendLine("      </mxGraphModel>")
    [void]$sb.AppendLine("  </diagram>")
    return $sb.ToString()
}

function Export-AADiagram {
<#
.SYNOPSIS
    Orchestrates the diagram generation for a single Auto Attendant.
    Returns a hashtable with Name, DiagramXml, NodeCount, EdgeCount, IsValid, and PhoneNumbers.
#>
    param(
        [object]$AutoAttendant, [hashtable]$ResourceAccountLookup,
        [hashtable]$AALookup, [hashtable]$CQLookup,
        [hashtable]$UserCache, [hashtable]$RAPhoneNumbers,
        [hashtable]$AAByAppInstance, [hashtable]$CQByAppInstance
    )

    $nodes = [System.Collections.Generic.List[hashtable]]::new()
    $edges = [System.Collections.Generic.List[hashtable]]::new()
    $nodeMap = @{}; $nextCellId = 2
    $definedNodes = [System.Collections.Generic.HashSet[string]]::new()

    $aa = $AutoAttendant
    $aaNodeId = "AA_$(Sanitise-NodeId $aa.Identity)"

    $phoneNumbers = @()
    foreach ($appInstance in $aa.ApplicationInstances) {
        if ($RAPhoneNumbers.ContainsKey($appInstance)) { $phoneNumbers += $RAPhoneNumbers[$appInstance] }
    }

    $aaLabel = "<b>$(Escape-XmlString $aa.Name)</b>"
    if ($phoneNumbers.Count -gt 0) {
        $aaLabel = "<b>$(Escape-XmlString $aa.Name)</b><br/>$(Escape-XmlString ($phoneNumbers -join ', '))"
    }

    # Title node
    $titleLabel = "$(Escape-XmlString $aa.Name) - Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    Add-DiagramNode -NodeId "${aaNodeId}_title" -Label $titleLabel -Type "Title" `
        -Tier -1 -BranchIndex 0 -PositionInBranch 0 -ParentNodeId "" `
        -Nodes ([ref]$nodes) -NodeMap ([ref]$nodeMap) -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes) | Out-Null

    # AA root node
    Add-DiagramNode -NodeId $aaNodeId -Label $aaLabel -Type "AA" `
        -Tier 0 -BranchIndex 0 -PositionInBranch 0 -ParentNodeId "" `
        -Nodes ([ref]$nodes) -NodeMap ([ref]$nodeMap) -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes) | Out-Null

    # Business hours schedule note
    $scheduleText = Get-BusinessHoursSchedule -AutoAttendant $aa
    if ($scheduleText) {
        $scheduleNodeId = "${aaNodeId}_schedule"
        $scheduleLabel = "<b>&#x1f552; Business Hours</b><br/>$scheduleText"
        Add-DiagramNode -NodeId $scheduleNodeId -Label $scheduleLabel -Type "Schedule" `
            -Tier 0 -BranchIndex 0 -PositionInBranch 99 -ParentNodeId $aaNodeId `
            -Nodes ([ref]$nodes) -NodeMap ([ref]$nodeMap) -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes) | Out-Null

    }

    $branchCounter = 0

    if ($aa.DefaultCallFlow) {
        Build-CallFlowNodes -CallFlow $aa.DefaultCallFlow -ParentNodeId $aaNodeId `
            -FlowType "BusinessHours" -LinkLabel "Business Hours" -AAIdentity $aa.Identity -Operator $aa.Operator `
            -BranchIndex $branchCounter -ResourceAccountLookup $ResourceAccountLookup `
            -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
            -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
            -Nodes ([ref]$nodes) -Edges ([ref]$edges) -NodeMap ([ref]$nodeMap) `
            -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes)
        $branchCounter++
    }

    $ahAssoc = $aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq "AfterHours" }
    if ($ahAssoc) {
        $ahFlow = $aa.CallFlows | Where-Object { $_.Id -eq $ahAssoc.CallFlowId }
        if ($ahFlow) {
            Build-CallFlowNodes -CallFlow $ahFlow -ParentNodeId $aaNodeId `
                -FlowType "AfterHours" -LinkLabel "After Hours" -AAIdentity $aa.Identity -Operator $aa.Operator `
                -BranchIndex $branchCounter -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
                -Nodes ([ref]$nodes) -Edges ([ref]$edges) -NodeMap ([ref]$nodeMap) `
                -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes)
            $branchCounter++
        }
    }

    # Deduplicate holiday flows by CallFlowId
    $holAssocs  = $aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq "Holiday" }
    $seenFlowIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($ha in $holAssocs) {
        $hf = $aa.CallFlows | Where-Object { $_.Id -eq $ha.CallFlowId } | Select-Object -First 1
        if ($hf -and $seenFlowIds.Add($hf.Id)) {
            $hn = if ($hf.Name) { Escape-XmlString $hf.Name } else { "Holiday" }
            $holidayDates = Get-HolidayScheduleDates -AutoAttendant $aa -HolidayAssociations $holAssocs -CallFlowId $hf.Id
            if ($holidayDates) { $hn = "$hn<br/>$holidayDates" }
            Build-CallFlowNodes -CallFlow $hf -ParentNodeId $aaNodeId `
                -FlowType "Holiday" -LinkLabel $hn -AAIdentity $aa.Identity -Operator $aa.Operator `
                -BranchIndex $branchCounter -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance `
                -Nodes ([ref]$nodes) -Edges ([ref]$edges) -NodeMap ([ref]$nodeMap) `
                -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes)
            $branchCounter++
        }
    }

    $pageDims   = Calculate-NodePositions -Nodes ([ref]$nodes)
    $diagramXml = Build-DiagramXml -Nodes $nodes -Edges $edges -NodeMap $nodeMap `
        -PageWidth $pageDims.PageWidth -PageHeight $pageDims.PageHeight

    $isValid = Test-DiagramIntegrity -Nodes $nodes -Edges $edges -NodeMap $nodeMap -DiagramName $aa.Name

    return @{
        Name          = $aa.Name
        PageId        = "page_$(Sanitise-NodeId $aa.Identity)"
        DiagramXml    = $diagramXml
        NodeCount     = $nodes.Count
        EdgeCount     = $edges.Count
        IsValid       = $isValid
        PhoneNumbers  = $phoneNumbers
    }
}

# ============================================================================
# MAIN SCRIPT
# ============================================================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Teams Auto Attendant Call Flow Exporter"    -ForegroundColor Cyan
Write-Host " Draw.io Diagram Generator v1.6"             -ForegroundColor Cyan
Write-Host " Style Preset : $StylePreset"                -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
$_scriptStart = Get-Date

# Verify Teams connection before executing
Write-Host "Verifying Microsoft Teams connection..." -ForegroundColor Yellow
try {
    $null = Get-CsTenant -ErrorAction Stop
} catch {
    Write-Host "[!] Error: Not connected to Microsoft Teams." -ForegroundColor Red
    Write-Host "Please run 'Connect-MicrosoftTeams' first in your PowerShell session before executing this script." -ForegroundColor Yellow
    exit 1
}

# Optional: shared voicemail group names via an existing Microsoft Graph session.
# Never connects or imports on its own - it only uses a session the admin already has.
if ((Get-Command Get-MgContext -ErrorAction SilentlyContinue) -and (Get-Command Get-MgGroup -ErrorAction SilentlyContinue)) {
    $mgContext = $null
    try { $mgContext = Get-MgContext -ErrorAction Stop } catch { }
    if ($mgContext) {
        $script:GraphGroupLookup = $true
        Write-Host "[+] Microsoft Graph session found - shared voicemail group names will be shown." -ForegroundColor Green
    }
}
if (-not $script:GraphGroupLookup) {
    Write-Host "[i] No Microsoft Graph session - shared voicemail shows without a group name." -ForegroundColor DarkGray
    Write-Host "    (Optional: Connect-MgGraph -Scopes Group.Read.All before running to include them.)" -ForegroundColor DarkGray
}

$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "[+] Created output directory: $OutputPath" -ForegroundColor Green
}

Write-Host "[1/5] Retrieving Auto Attendants..." -ForegroundColor Yellow
$autoAttendants = [System.Collections.Generic.List[object]]::new()
$skip = 0; $batchSize = 100
try {
    do {
        $batch = Get-CsAutoAttendant -IncludeStatus -First $batchSize -Skip $skip -ErrorAction Stop
        if ($batch) { foreach ($item in $batch) { [void]$autoAttendants.Add($item) } }
        $skip += $batchSize
    } while ($batch -and $batch.Count -eq $batchSize)
} catch {
    Write-Host "[!] Error retrieving Auto Attendants: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "    Aborting export - continuing with partial data would produce an incomplete diagram set." -ForegroundColor Yellow
    exit 1
}
Write-Host "       Found $($autoAttendants.Count) Auto Attendant(s)" -ForegroundColor Gray

Write-Host "[2/5] Retrieving Call Queues..." -ForegroundColor Yellow
$callQueues = [System.Collections.Generic.List[object]]::new()
$skip = 0
try {
    do {
        $batch = Get-CsCallQueue -First $batchSize -Skip $skip -WarningAction SilentlyContinue -ErrorAction Stop
        if ($batch) { foreach ($item in $batch) { [void]$callQueues.Add($item) } }
        $skip += $batchSize
    } while ($batch -and $batch.Count -eq $batchSize)
} catch {
    Write-Host "[!] Error retrieving Call Queues: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "    Aborting export - continuing with partial data would produce an incomplete diagram set." -ForegroundColor Yellow
    exit 1
}
Write-Host "       Found $($callQueues.Count) Call Queue(s)" -ForegroundColor Gray

Write-Host "[3/5] Retrieving Resource Accounts..." -ForegroundColor Yellow
$resourceAccounts = [System.Collections.Generic.List[object]]::new()
$skip = 0
try {
    do {
        $batch = Get-CsOnlineApplicationInstance -ResultSize $batchSize -Skip $skip -ErrorAction Stop
        if ($batch) { foreach ($item in $batch) { [void]$resourceAccounts.Add($item) } }
        $skip += $batchSize
    } while ($batch -and $batch.Count -eq $batchSize)
} catch {
    Write-Host "[!] Error retrieving Resource Accounts: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "    Aborting export - continuing with partial data would produce an incomplete diagram set." -ForegroundColor Yellow
    exit 1
}
Write-Host "       Found $($resourceAccounts.Count) Resource Account(s)" -ForegroundColor Gray

Write-Host "[4/5] Building lookup tables..." -ForegroundColor Yellow

$ResourceAccountLookup = @{}
foreach ($ra in $resourceAccounts) { $ResourceAccountLookup[$ra.ObjectId] = $ra }

$AALookup = @{}
foreach ($aa in $autoAttendants) { $AALookup[$aa.Identity] = $aa }

$CQLookup = @{}
foreach ($cq in $callQueues) { $CQLookup[$cq.Identity] = $cq }

$UserCache = @{}

$RAPhoneNumbers = @{}
foreach ($ra in $resourceAccounts) {
    if ($ra.PhoneNumber) {
        $RAPhoneNumbers[$ra.ObjectId] = ($ra.PhoneNumber -replace 'tel:', '')
    }
}

# Reverse lookups (ApplicationInstance ID -> AA/CQ) so Resolve-CallTarget's
# "resource account missing/unlisted" fallback is O(1) instead of scanning
# every AA/CQ on every miss.
$AAByAppInstance = @{}
foreach ($item in $autoAttendants) {
    foreach ($appInstance in $item.ApplicationInstances) { $AAByAppInstance[$appInstance] = $item }
}

$CQByAppInstance = @{}
foreach ($item in $callQueues) {
    foreach ($appInstance in $item.ApplicationInstances) { $CQByAppInstance[$appInstance] = $item }
}

Write-Host "       Lookup tables ready" -ForegroundColor Gray

Write-Host "[5/5] Generating draw.io diagrams..." -ForegroundColor Yellow

$allDiagrams  = [System.Collections.Generic.List[hashtable]]::new()
$summaryRows  = [System.Collections.Generic.List[hashtable]]::new()
$aaCounter    = 0
$warnedCount  = 0
$failedCount  = 0

foreach ($aa in $autoAttendants) {
    $aaCounter++
    $aaStart = Get-Date
    Write-Host "       [$aaCounter/$($autoAttendants.Count)] Processing: $($aa.Name)" -ForegroundColor Gray

    try {
        $result = Export-AADiagram -AutoAttendant $aa `
            -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup `
            -CQLookup $CQLookup -UserCache $UserCache -RAPhoneNumbers $RAPhoneNumbers `
            -AAByAppInstance $AAByAppInstance -CQByAppInstance $CQByAppInstance

        if (-not $result.IsValid) { $warnedCount++ }

        $safeFileName = ($aa.Name -replace '[\\/\:\*\?"<>\|]', '_')
        $diagramId    = Sanitise-NodeId $aa.Identity
        $timestamp    = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.000Z"

        $fileContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="app.diagrams.net" modified="$timestamp" agent="Teams Call Flow Export" version="21.0.0" type="device">
  <diagram id="$diagramId" name="$(Escape-XmlString $aa.Name)">
$($result.DiagramXml)
  </diagram>
</mxfile>
"@

        $filePath  = Join-Path $OutputPath "$safeFileName.drawio"
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($filePath, $fileContent, $utf8NoBom)

        $fileSize = (Get-Item $filePath).Length
        $elapsed  = (Get-Date) - $aaStart
        $validMark = if ($result.IsValid) { "" } else { " [!WARNINGS]" }
        Write-Host "         Saved: $filePath  ($($result.NodeCount) nodes, $($result.EdgeCount) edges, $([Math]::Round($fileSize/1KB,1)) KB, $([Math]::Round($elapsed.TotalSeconds,1))s)$validMark" -ForegroundColor DarkGreen

        [void]$allDiagrams.Add($result)
        [void]$summaryRows.Add(@{
            AutoAttendant = $aa.Name
            File          = $filePath
            Nodes         = $result.NodeCount
            Edges         = $result.EdgeCount
            FileSizeBytes = $fileSize
            StylePreset   = $StylePreset
            Valid         = $result.IsValid
            GeneratedAt   = $timestamp
        })
    } catch {
        $failedCount++
        Write-Host "         [!] Failed to export '$($aa.Name)': $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "         Skipping this Auto Attendant and continuing with the rest." -ForegroundColor Yellow
        continue
    }
}

$timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.000Z"
$combinedSb = [System.Text.StringBuilder]::new()
[void]$combinedSb.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
[void]$combinedSb.AppendLine("<mxfile host=""app.diagrams.net"" modified=""$timestamp"" agent=""Teams Call Flow Export"" version=""21.0.0"" type=""device"">")

$indexXml = Build-IndexPage -DiagramId "index_page" -Diagrams $allDiagrams
[void]$combinedSb.AppendLine($indexXml)

$legendXml = Build-LegendPage -DiagramId "legend_page"
[void]$combinedSb.AppendLine($legendXml)

# Page ids are derived from the AA identity so nested-AA nodes and index rows
# can link to them ("data:page/id,page_<identity>").
foreach ($diagram in $allDiagrams) {
    $pageName = Escape-XmlString $diagram.Name
    [void]$combinedSb.AppendLine("  <diagram id=""$($diagram.PageId)"" name=""$pageName"">")
    [void]$combinedSb.AppendLine($diagram.DiagramXml)
    [void]$combinedSb.AppendLine("  </diagram>")
}

[void]$combinedSb.AppendLine("</mxfile>")

$combinedPath = Join-Path $OutputPath "_AllCallFlows.drawio"
$utf8NoBom    = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($combinedPath, $combinedSb.ToString(), $utf8NoBom)

# ---- Write _ExportSummary.json ----
$totalElapsed  = (Get-Date) - $_scriptStart
$totalNodesSum = 0; $totalEdgesSum = 0; $totalFileSizeSum = 0
foreach ($row in $summaryRows) {
    $totalNodesSum    += $row.Nodes
    $totalEdgesSum    += $row.Edges
    $totalFileSizeSum += $row.FileSizeBytes
}
$summaryObject = @{
    ExportedAt      = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
    StylePreset     = $StylePreset
    TotalDuration_s = [Math]::Round($totalElapsed.TotalSeconds, 1)
    AutoAttendants  = $summaryRows
    TotalNodes      = $totalNodesSum
    TotalEdges      = $totalEdgesSum
    TotalFileSizeKB = [Math]::Round($totalFileSizeSum / 1KB, 1)
    DiagramsWithWarnings = $warnedCount
    DiagramsFailed        = $failedCount
}
$summaryPath = Join-Path $OutputPath "_ExportSummary.json"
$summaryObject | ConvertTo-Json -Depth 5 | Set-Content -Path $summaryPath -Encoding UTF8

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " Export Complete!" -ForegroundColor Green
Write-Host " Files saved to : $OutputPath" -ForegroundColor Green
Write-Host " Individual files: $($allDiagrams.Count) of $($autoAttendants.Count) Auto Attendant(s)" -ForegroundColor Green
Write-Host " Combined file  : _AllCallFlows.drawio" -ForegroundColor Green
Write-Host " Summary JSON   : _ExportSummary.json" -ForegroundColor Green
Write-Host " Diagrams with validation warnings: $warnedCount" -ForegroundColor $(if ($warnedCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host " Diagrams that failed to export: $failedCount" -ForegroundColor $(if ($failedCount -gt 0) { 'Red' } else { 'Green' })
Write-Host " Total duration : $([Math]::Round($totalElapsed.TotalSeconds,1))s" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "TIP: Open the .drawio files in draw.io Desktop, diagrams.net," -ForegroundColor DarkYellow
Write-Host "     or VS Code with the draw.io extension to view and edit" -ForegroundColor DarkYellow
Write-Host "     the diagrams interactively." -ForegroundColor DarkYellow