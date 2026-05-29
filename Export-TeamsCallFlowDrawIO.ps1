<#
.SYNOPSIS
    Export-TeamsCallFlowDrawIO.ps1
    Exports all Microsoft Teams Auto Attendants and Call Queues as draw.io diagrams.

.DESCRIPTION
    Connects to Microsoft Teams PowerShell, retrieves all Auto Attendants, Call Queues,
    and Resource Accounts, then generates draw.io (.drawio) XML files showing the complete
    call flow including business hours, after hours, holidays, menu options, Call Queue
    routing, timeout/overflow handling, nested Auto Attendants, TTS/audio greetings,
    and business hours schedules.

    Each Auto Attendant gets its own .drawio file, plus a combined _AllCallFlows.drawio
    is created with one page per Auto Attendant and a Legend page.

.PARAMETER OutputPath
    The folder where .drawio files will be saved. Defaults to .\CallFlowDiagrams

.EXAMPLE
    .\Export-TeamsCallFlowDrawIO.ps1
    .\Export-TeamsCallFlowDrawIO.ps1 -OutputPath "C:\Customers\Contoso\CallFlows"

.NOTES
    Author  : Kieran Lloyd
    Version : 1.4
    Date    : 2026-05-28

    Prerequisites:
      - MicrosoftTeams PowerShell module installed
      - Connected to Teams via Connect-MicrosoftTeams
      - Sufficient admin permissions to read AA/CQ/Resource Account configuration

    Changelog:
    v1.4 - Added pagination loops for Get-CsAutoAttendant, Get-CsCallQueue, and
           Get-CsOnlineApplicationInstance to support tenants with more than 100
           Auto Attendants, Call Queues, or Resource Accounts.
    v1.3 - Improved node positioning for cleaner, more compact diagrams
         - Reduced branch gap, target gap, and CQ child offsets
         - Tightened greeting and schedule note positioning
         - Capped max branch width to prevent excessive horizontal spread
         - Adjusted tier Y positions for better vertical rhythm
    v1.2 - Added business hours schedule note beside each AA root node
         - Fixed voice app resolution when resource account not in lookup
           (falls back to direct AA/CQ identity lookup then ApplicationInstances scan)
    v1.1 - Added ConfigurationEndpoint support in Resolve-CallTarget
         - Added TTS and audio greeting notes on each call flow
    v1.0 - Initial draw.io version based on Mermaid v1.2
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\CallFlowDiagrams"
)

# ============================================================================
# CONFIGURATION
# ============================================================================
$ErrorActionPreference = "Continue"

# ============================================================================
# NODE STYLE DEFINITIONS
# ============================================================================
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
}

# ============================================================================
# EDGE STYLE DEFINITIONS
# ============================================================================
$EdgeStyles = @{
    "BusinessHours"    = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#666666;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
    "AfterHours"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#2E75B6;strokeWidth=2;fontFamily=Segoe UI;fontSize=10;"
    "Holiday"          = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#BF8F00;strokeWidth=2;dashed=1;fontFamily=Segoe UI;fontSize=10;"
    "MenuOption"       = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#333333;strokeWidth=1;fontFamily=Segoe UI;fontSize=10;"
    "TimeoutOverflow"  = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#ED7D31;strokeWidth=1;dashed=1;fontFamily=Segoe UI;fontSize=10;"
    "Greeting"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#D6B656;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
    "Schedule"         = "edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=#6C8EBF;strokeWidth=1;dashed=1;dashPattern=3 3;fontFamily=Segoe UI;fontSize=9;"
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
    Ampersand is replaced FIRST to avoid double-encoding.
#>
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text = $Text -replace '&', '&amp;'
    $Text = $Text -replace '<', '&lt;'
    $Text = $Text -replace '>', '&gt;'
    $Text = $Text -replace '"', '&quot;'
    $Text = $Text -replace "'", '&apos;'
    return $Text
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

function Get-BusinessHoursSchedule {
<#
.SYNOPSIS
    Extracts the business hours schedule from an Auto Attendant's AfterHours
    call handling association and formats it as an HTML string for display.
    Returns $null if no schedule is configured.
#>
    param([object]$AutoAttendant)

    $afterHoursAssoc = $AutoAttendant.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq "AfterHours" }
    if (-not $afterHoursAssoc) { return $null }

    $scheduleId = $afterHoursAssoc.ScheduleId
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
                $startTime = $tr.Start.ToString("hh\:mm")
                $endTime = $tr.End.ToString("hh\:mm")
                [void]$ranges.Add("$startTime - $endTime")
            }
            [void]$lines.Add("${abbrev}: $($ranges -join ', ')")
        }
        else {
            [void]$lines.Add("${abbrev}: Closed")
        }
    }

    return ($lines -join "<br/>")
}

function Resolve-CallTarget {
<#
.SYNOPSIS
    Resolves a call target to a display name, type, and linked entity.
    Handles ApplicationEndpoint, ConfigurationEndpoint, User, ExternalPstn,
    SharedVoicemail, and DisconnectCall actions.
    Falls back to direct AA/CQ identity lookup if resource account not found.
.OUTPUTS
    Hashtable with keys: DisplayName, Type, LinkedId
#>
    param(
        [object]$CallTarget,
        [string]$Action,
        [hashtable]$ResourceAccountLookup,
        [hashtable]$AALookup,
        [hashtable]$CQLookup,
        [hashtable]$UserCache
    )

    if ($Action -eq "DisconnectCall") {
        return @{ DisplayName = "Disconnect"; Type = "Disconnect"; LinkedId = $null }
    }

    if ($null -eq $CallTarget) {
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
                    $linkedAA = $AALookup.Values | Where-Object {
                        $_.ApplicationInstances -contains $targetId
                    } | Select-Object -First 1

                    if ($linkedAA) {
                        return @{ DisplayName = $linkedAA.Name; Type = "AA"; LinkedId = $linkedAA.Identity }
                    }
                    return @{ DisplayName = $ra.DisplayName; Type = "AA"; LinkedId = $null }
                }
                elseif ($appId -eq "11cd3e2e-fccb-42ad-ad00-878b93575e07") {
                    $linkedCQ = $CQLookup.Values | Where-Object {
                        $_.ApplicationInstances -contains $targetId
                    } | Select-Object -First 1

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
            $matchedAA = $AALookup.Values | Where-Object {
                $_.ApplicationInstances -contains $targetId
            } | Select-Object -First 1
            if ($matchedAA) {
                return @{ DisplayName = $matchedAA.Name; Type = "AA"; LinkedId = $matchedAA.Identity }
            }

            $matchedCQ = $CQLookup.Values | Where-Object {
                $_.ApplicationInstances -contains $targetId
            } | Select-Object -First 1
            if ($matchedCQ) {
                return @{ DisplayName = $matchedCQ.Name; Type = "CQ"; LinkedId = $matchedCQ.Identity }
            }

            return @{ DisplayName = "Unknown Voice App ($targetId)"; Type = "Unknown"; LinkedId = $null }
        }
        "User" {
            if ($UserCache.ContainsKey($targetId)) {
                $displayName = $UserCache[$targetId]
            }
            else {
                try {
                    $user = Get-CsOnlineUser -Identity $targetId -ErrorAction Stop
                    $displayName = $user.DisplayName
                    $UserCache[$targetId] = $displayName
                }
                catch {
                    $displayName = "User ($targetId)"
                    $UserCache[$targetId] = $displayName
                }
            }
            return @{ DisplayName = $displayName; Type = "User"; LinkedId = $targetId }
        }
        "ExternalPstn" {
            $phoneNumber = $targetId -replace 'tel:', ''
            return @{ DisplayName = "External $phoneNumber"; Type = "ExternalPstn"; LinkedId = $null }
        }
        "SharedVoicemail" {
            return @{ DisplayName = "Shared Voicemail"; Type = "SharedVoicemail"; LinkedId = $targetId }
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
#>
    param(
        [string]$NodeId, [string]$Label, [string]$Type,
        [int]$Tier, [int]$BranchIndex, [int]$PositionInBranch, [string]$ParentNodeId,
        [ref]$Nodes, [ref]$NodeMap, [ref]$NextCellId, [ref]$DefinedNodes
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
#>
    param(
        [string]$SourceNodeId, [string]$TargetNodeId, [string]$Label, [string]$StyleKey,
        [ref]$Edges, [ref]$NextCellId
    )

    $cellId = $NextCellId.Value
    $NextCellId.Value++

    $style = $EdgeStyles[$StyleKey]
    if (-not $style) { $style = $EdgeStyles["BusinessHours"] }

    $edge = @{
        CellId = $cellId; SourceNodeId = $SourceNodeId; TargetNodeId = $TargetNodeId
        Label = $Label; Style = $style
    }

    [void]$Edges.Value.Add($edge)
    return $cellId
}

function Build-CallQueueNodes {
<#
.SYNOPSIS
    Builds draw.io nodes for a Call Queue including agents info, timeout, and overflow.
#>
    param(
        [object]$CallQueue, [string]$CQNodeId,
        [int]$Tier, [int]$BranchIndex, [int]$PositionInBranch, [string]$ParentNodeId,
        [hashtable]$ResourceAccountLookup, [hashtable]$AALookup,
        [hashtable]$CQLookup, [hashtable]$UserCache,
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

    $childPos = 0
    if ($CallQueue.TimeoutAction) {
        $timeoutActionStr = $CallQueue.TimeoutAction.ToString()
        $timeoutNodeId = "${CQNodeId}_timeout"
        $timeoutThreshold = $CallQueue.TimeoutThreshold
        if (-not $timeoutThreshold) { $timeoutThreshold = "N/A" }

        if ($timeoutActionStr -ne "Disconnect") {
            $timeoutTarget = Resolve-CallTarget -CallTarget $CallQueue.TimeoutActionTarget `
                -Action $timeoutActionStr -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache
            $timeoutLabel = "Timeout (${timeoutThreshold}s)<br/>$(Escape-XmlString $timeoutTarget.DisplayName)"
            Add-DiagramNode -NodeId $timeoutNodeId -Label $timeoutLabel -Type "TimeoutOverflow" `
                -Tier ($Tier + 1) -BranchIndex $BranchIndex -PositionInBranch $childPos `
                -ParentNodeId $CQNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        } else {
            $timeoutLabel = "Timeout (${timeoutThreshold}s)<br/>Disconnect"
            Add-DiagramNode -NodeId $timeoutNodeId -Label $timeoutLabel -Type "Disconnect" `
                -Tier ($Tier + 1) -BranchIndex $BranchIndex -PositionInBranch $childPos `
                -ParentNodeId $CQNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        }
        Add-DiagramEdge -SourceNodeId $CQNodeId -TargetNodeId $timeoutNodeId `
            -Label "Timeout" -StyleKey "TimeoutOverflow" -Edges $Edges -NextCellId $NextCellId | Out-Null
        $childPos++
    }

    if ($CallQueue.OverflowAction) {
        $overflowActionStr = $CallQueue.OverflowAction.ToString()
        $overflowNodeId = "${CQNodeId}_overflow"
        $overflowThreshold = $CallQueue.OverflowThreshold
        if (-not $overflowThreshold) { $overflowThreshold = "N/A" }

        if ($overflowActionStr -ne "Disconnect") {
            $overflowTarget = Resolve-CallTarget -CallTarget $CallQueue.OverflowActionTarget `
                -Action $overflowActionStr -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache
            $overflowLabel = "Overflow ($overflowThreshold calls)<br/>$(Escape-XmlString $overflowTarget.DisplayName)"
            Add-DiagramNode -NodeId $overflowNodeId -Label $overflowLabel -Type "TimeoutOverflow" `
                -Tier ($Tier + 1) -BranchIndex $BranchIndex -PositionInBranch $childPos `
                -ParentNodeId $CQNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        } else {
            $overflowLabel = "Overflow ($overflowThreshold calls)<br/>Disconnect"
            Add-DiagramNode -NodeId $overflowNodeId -Label $overflowLabel -Type "Disconnect" `
                -Tier ($Tier + 1) -BranchIndex $BranchIndex -PositionInBranch $childPos `
                -ParentNodeId $CQNodeId `
                -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
        }
        Add-DiagramEdge -SourceNodeId $CQNodeId -TargetNodeId $overflowNodeId `
            -Label "Overflow" -StyleKey "TimeoutOverflow" -Edges $Edges -NextCellId $NextCellId | Out-Null
    }
}

function Build-CallFlowNodes {
<#
.SYNOPSIS
    Processes a call flow (business hours, after hours, or holiday) and generates
    draw.io nodes and edges, including greeting notes where present.
#>
    param(
        [object]$CallFlow, [string]$ParentNodeId, [string]$FlowType,
        [string]$LinkLabel, [string]$AAIdentity, [int]$BranchIndex,
        [hashtable]$ResourceAccountLookup, [hashtable]$AALookup,
        [hashtable]$CQLookup, [hashtable]$UserCache,
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
    $hasIvrOptions = $menuOptions | Where-Object { $_.DtmfResponse -ne "Automatic" -or $_.Action -ne "DisconnectCall" }

    if ($hasIvrOptions -and $hasIvrOptions.Count -gt 0) {
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
        foreach ($option in $menuOptions) {
            $dtmfKey = Get-DtmfDisplayKey $option.DtmfResponse.ToString()
            $action = $option.Action.ToString()

            $target = Resolve-CallTarget -CallTarget $option.CallTarget -Action $action `
                -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache

            $targetNodeId = ""; $targetType = $target.Type

            switch ($targetType) {
                "AA" {
                    $targetNodeId = "AA_$(Sanitise-NodeId $target.LinkedId)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "<b>$(Escape-XmlString $target.DisplayName)</b>" -Type "AA" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "CQ" {
                    $targetNodeId = "CQ_$(Sanitise-NodeId $target.LinkedId)"
                    if ($target.LinkedId -and $CQLookup.ContainsKey($target.LinkedId)) {
                        Build-CallQueueNodes -CallQueue $CQLookup[$target.LinkedId] -CQNodeId $targetNodeId `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                            -Nodes $Nodes -Edges $Edges -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes
                    } elseif (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "<b>$(Escape-XmlString $target.DisplayName)</b>" -Type "CQ" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "User" {
                    $targetNodeId = "User_$(Sanitise-NodeId $target.LinkedId)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $target.DisplayName) -Type "User" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "ExternalPstn" {
                    $targetNodeId = "${flowPrefix}_ext_$(Sanitise-NodeId $dtmfKey)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $target.DisplayName) -Type "ExternalPstn" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "SharedVoicemail" {
                    $targetNodeId = "${flowPrefix}_vm_$(Sanitise-NodeId $dtmfKey)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "Shared Voicemail" -Type "SharedVoicemail" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "Disconnect" {
                    $targetNodeId = "${flowPrefix}_disc_$(Sanitise-NodeId $dtmfKey)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "Disconnect" -Type "Disconnect" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                default {
                    $targetNodeId = "${flowPrefix}_unk_$(Sanitise-NodeId $dtmfKey)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $target.DisplayName) -Type "AA" `
                            -Tier 2 -BranchIndex $BranchIndex -PositionInBranch $optionIndex -ParentNodeId $menuNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
            }

            if ($targetNodeId) {
                $optionLabel = "Press $dtmfKey"
                if ($option.VoiceResponses) {
                    $voicePrompt = ($option.VoiceResponses | Select-Object -First 1)
                    if ($voicePrompt) { $optionLabel = "Press $dtmfKey / Say $(Escape-XmlString $voicePrompt)" }
                }
                Add-DiagramEdge -SourceNodeId $menuNodeId -TargetNodeId $targetNodeId `
                    -Label $optionLabel -StyleKey "MenuOption" -Edges $Edges -NextCellId $NextCellId | Out-Null
            }
            $optionIndex++
        }
    }
    else {
        # No IVR menu — direct transfer
        $defaultAction = $menu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -eq "Automatic" } | Select-Object -First 1
        if ($defaultAction) {
            $target = Resolve-CallTarget -CallTarget $defaultAction.CallTarget `
                -Action $defaultAction.Action.ToString() -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache

            $targetNodeId = ""
            switch ($target.Type) {
                "CQ" {
                    $targetNodeId = "CQ_$(Sanitise-NodeId $target.LinkedId)"
                    if ($target.LinkedId -and $CQLookup.ContainsKey($target.LinkedId)) {
                        Build-CallQueueNodes -CallQueue $CQLookup[$target.LinkedId] -CQNodeId $targetNodeId `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                            -Nodes $Nodes -Edges $Edges -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes
                    } elseif (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "<b>$(Escape-XmlString $target.DisplayName)</b>" -Type "CQ" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "AA" {
                    $targetNodeId = "AA_$(Sanitise-NodeId $target.LinkedId)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "<b>$(Escape-XmlString $target.DisplayName)</b>" -Type "AA" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "User" {
                    $targetNodeId = "User_$(Sanitise-NodeId $target.LinkedId)"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $target.DisplayName) -Type "User" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "ExternalPstn" {
                    $targetNodeId = "${flowPrefix}_ext_default"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label (Escape-XmlString $target.DisplayName) -Type "ExternalPstn" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "SharedVoicemail" {
                    $targetNodeId = "${flowPrefix}_vm_default"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "Shared Voicemail" -Type "SharedVoicemail" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
                "Disconnect" {
                    $targetNodeId = "${flowPrefix}_disc_default"
                    if (-not $DefinedNodes.Value.Contains($targetNodeId)) {
                        Add-DiagramNode -NodeId $targetNodeId -Label "Disconnect" -Type "Disconnect" `
                            -Tier 1 -BranchIndex $BranchIndex -PositionInBranch 0 -ParentNodeId $ParentNodeId `
                            -Nodes $Nodes -NodeMap $NodeMap -NextCellId $NextCellId -DefinedNodes $DefinedNodes | Out-Null
                    }
                }
            }

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

function Calculate-NodePositions {
<#
.SYNOPSIS
    Calculates X/Y positions for all nodes using a tier-based top-down hierarchical layout.
    Prevents overlaps dynamically using boundary-aware footprints for all nodes.
    Returns a hashtable with PageWidth and PageHeight.
#>
    param([ref]$Nodes)

    $tierYPositions = @{ 0 = 60; 1 = 200; 2 = 380; 3 = 540 }
    $branchGap = 160; $siblingGap = 50; $cqChildXOffset = 110; $cqChildYOffset = 140

    $tier0 = [System.Collections.Generic.List[hashtable]]::new()
    $tier1 = [System.Collections.Generic.List[hashtable]]::new()
    $tier2 = [System.Collections.Generic.List[hashtable]]::new()
    $tier3 = [System.Collections.Generic.List[hashtable]]::new()
    $titleNodes = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($node in $Nodes.Value) {
        if ($node.Type -eq "Title") { [void]$titleNodes.Add($node); continue }
        if ($node.Type -in @("Greeting","Schedule")) { continue }
        switch ($node.Tier) {
            0 { [void]$tier0.Add($node) }
            1 { [void]$tier1.Add($node) }
            2 { [void]$tier2.Add($node) }
            3 { [void]$tier3.Add($node) }
        }
    }

    $tier1Branches = @{}
    foreach ($n in $tier1) {
        if (-not $tier1Branches.ContainsKey($n.BranchIndex)) {
            $tier1Branches[$n.BranchIndex] = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void]$tier1Branches[$n.BranchIndex].Add($n)
    }

    $tier2Branches = @{}
    $cqChildNodes = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($n in $tier2) {
        if ($n.NodeId -match '_(timeout|overflow)$') { [void]$cqChildNodes.Add($n); continue }
        if (-not $tier2Branches.ContainsKey($n.BranchIndex)) {
            $tier2Branches[$n.BranchIndex] = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void]$tier2Branches[$n.BranchIndex].Add($n)
    }

    # 1. Calculate footprints for all non-note, non-child nodes
    # Footprint defines LeftOffset and RightOffset relative to the X coordinate.
    $footprints = @{}
    foreach ($n in $Nodes.Value) {
        if ($n.Type -in @("Greeting","Schedule","Title") -or $n.NodeId -match '_(timeout|overflow)$') {
            continue
        }

        $leftOffset = 0
        $rightOffset = $n.Width

        # Check for greeting note
        $greeting = $Nodes.Value | Where-Object { $_.Type -eq "Greeting" -and $_.ParentNodeId -eq $n.NodeId } | Select-Object -First 1
        if ($greeting) {
            $rightOffset = [Math]::Max($rightOffset, $n.Width + 15 + $greeting.Width)
        }

        # Check for CQ children (Timeout and Overflow)
        if ($n.Type -eq "CQ") {
            $timeout = $Nodes.Value | Where-Object { $_.NodeId -eq "$($n.NodeId)_timeout" } | Select-Object -First 1
            if ($timeout) {
                $tOffset = [int](($n.Width / 2) - $cqChildXOffset - ($timeout.Width / 2))
                if ($tOffset -lt $leftOffset) { $leftOffset = $tOffset }
            }
            $overflow = $Nodes.Value | Where-Object { $_.NodeId -eq "$($n.NodeId)_overflow" } | Select-Object -First 1
            if ($overflow) {
                $oOffset = [int](($n.Width / 2) + $cqChildXOffset + ($overflow.Width / 2))
                if ($oOffset -gt $rightOffset) { $rightOffset = $oOffset }
            }
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
            $sortedT1 = @($t1Nodes | Sort-Object PositionInBranch)
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
            $sortedT2 = @($t2Nodes | Sort-Object PositionInBranch)
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

    # 5. Position schedule note as a premium "Info Panel" on the far-right sidebar
    $scheduleNodes = $Nodes.Value | Where-Object { $_.Type -eq "Schedule" }
    if ($scheduleNodes.Count -gt 0) {
        $flowMaxX = 0
        foreach ($node in $Nodes.Value) {
            if ($node.Type -in @("Title","Schedule","Greeting")) {
                continue
            }
            $r = $node.X + $node.Width
            
            # Include attached greeting note footprints
            $greeting = $Nodes.Value | Where-Object { $_.Type -eq "Greeting" -and $_.ParentNodeId -eq $node.NodeId } | Select-Object -First 1
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
            $sn.Y = 60 # Perfectly aligned vertically with the root AA node
        }
    }

    # 6. Position CQ children (timeout and overflow)
    $allCqChildren = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($n in $tier3) { [void]$allCqChildren.Add($n) }
    foreach ($n in $cqChildNodes) { [void]$allCqChildren.Add($n) }

    foreach ($n in $allCqChildren) {
        $pn = $Nodes.Value | Where-Object { $_.NodeId -eq $n.ParentNodeId } | Select-Object -First 1
        if ($pn) {
            $bx = $pn.X + ($pn.Width / 2)
            if ($n.NodeId -match '_timeout$') {
                $n.X = [int]($bx - $cqChildXOffset - ($n.Width / 2))
            } else {
                $n.X = [int]($bx + $cqChildXOffset - ($n.Width / 2))
            }
            $n.Y = $pn.Y + $cqChildYOffset
        } else {
            $n.X = $startX
            $n.Y = $tierYPositions[3]
        }
    }

    # 7. Position title nodes
    foreach ($n in $titleNodes) {
        if ($tier0.Count -gt 0) { $n.X = $tier0[0].X - 90 } else { $n.X = $startX }
        $n.Y = 10
    }

    # 8. Position greeting notes
    $greetingNodes = $Nodes.Value | Where-Object { $_.Type -eq "Greeting" }
    foreach ($gn in $greetingNodes) {
        $pn = $Nodes.Value | Where-Object { $_.NodeId -eq $gn.ParentNodeId } | Select-Object -First 1
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
        $safeLabel = Escape-XmlString $node.Label
        [void]$sb.AppendLine("          <mxCell id=""$($node.CellId)"" value=""$safeLabel"" style=""$($node.Style)"" vertex=""1"" parent=""1"">")
        [void]$sb.AppendLine("            <mxGeometry x=""$($node.X)"" y=""$($node.Y)"" width=""$($node.Width)"" height=""$($node.Height)"" as=""geometry""/>")
        [void]$sb.AppendLine("          </mxCell>")
    }

    foreach ($edge in $Edges) {
        $sc = ""; $tc = ""
        if ($NodeMap.ContainsKey($edge.SourceNodeId)) { $sc = $NodeMap[$edge.SourceNodeId] }
        if ($NodeMap.ContainsKey($edge.TargetNodeId)) { $tc = $NodeMap[$edge.TargetNodeId] }
        if ($sc -and $tc) {
            $sl = Escape-XmlString $edge.Label
            [void]$sb.AppendLine("          <mxCell id=""$($edge.CellId)"" value=""$sl"" style=""$($edge.Style)"" edge=""1"" source=""$sc"" target=""$tc"" parent=""1"">")
            [void]$sb.AppendLine("            <mxGeometry relative=""1"" as=""geometry""/>")
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
        @{ Label = "Shared Voicemail"; Type = "SharedVoicemail" }
        @{ Label = "Disconnect"; Type = "Disconnect" }
        @{ Label = "Holiday"; Type = "Holiday" }
        @{ Label = "Timeout / Overflow"; Type = "TimeoutOverflow" }
        @{ Label = "TTS / Audio Greeting"; Type = "Greeting" }
        @{ Label = "Business Hours Schedule"; Type = "Schedule" }
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
        @{ Label = "Timeout / Overflow"; Color = "#ED7D31"; Dashed = "1" }
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

    $edgeIdx = 0
    foreach ($ei in $edgeLegendItems) {
        $edgeId = $cellId; $cellId++
        $ds = ""; if ($ei.Dashed -eq "1") { $ds = "dashed=1;" }
        $base = 3 + $legendItems.Count + ($edgeIdx * 3)
        [void]$sb.AppendLine("          <mxCell id=""$edgeId"" value="""" style=""edgeStyle=orthogonalEdgeStyle;rounded=1;orthogonalLoop=1;jettySize=auto;html=1;strokeColor=$($ei.Color);strokeWidth=2;${ds}fontFamily=Segoe UI;fontSize=10;"" edge=""1"" source=""$base"" target=""$($base + 1)"" parent=""1"">")
        [void]$sb.AppendLine("            <mxGeometry relative=""1"" as=""geometry""/>")
        [void]$sb.AppendLine("          </mxCell>")
        $edgeIdx++
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
    Returns a hashtable with Name and DiagramXml.
#>
    param(
        [object]$AutoAttendant, [hashtable]$ResourceAccountLookup,
        [hashtable]$AALookup, [hashtable]$CQLookup,
        [hashtable]$UserCache, [hashtable]$RAPhoneNumbers
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
            -FlowType "BusinessHours" -LinkLabel "Business Hours" -AAIdentity $aa.Identity `
            -BranchIndex $branchCounter -ResourceAccountLookup $ResourceAccountLookup `
            -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
            -Nodes ([ref]$nodes) -Edges ([ref]$edges) -NodeMap ([ref]$nodeMap) `
            -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes)
        $branchCounter++
    }

    $ahAssoc = $aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq "AfterHours" }
    if ($ahAssoc) {
        $ahFlow = $aa.CallFlows | Where-Object { $_.Id -eq $ahAssoc.CallFlowId }
        if ($ahFlow) {
            Build-CallFlowNodes -CallFlow $ahFlow -ParentNodeId $aaNodeId `
                -FlowType "AfterHours" -LinkLabel "After Hours" -AAIdentity $aa.Identity `
                -BranchIndex $branchCounter -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -Nodes ([ref]$nodes) -Edges ([ref]$edges) -NodeMap ([ref]$nodeMap) `
                -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes)
            $branchCounter++
        }
    }

    $holAssocs = $aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq "Holiday" }
    foreach ($ha in $holAssocs) {
        $hf = $aa.CallFlows | Where-Object { $_.Id -eq $ha.CallFlowId }
        if ($hf) {
            $hn = $hf.Name; if (-not $hn) { $hn = "Holiday" }
            Build-CallFlowNodes -CallFlow $hf -ParentNodeId $aaNodeId `
                -FlowType "Holiday" -LinkLabel $hn -AAIdentity $aa.Identity `
                -BranchIndex $branchCounter -ResourceAccountLookup $ResourceAccountLookup `
                -AALookup $AALookup -CQLookup $CQLookup -UserCache $UserCache `
                -Nodes ([ref]$nodes) -Edges ([ref]$edges) -NodeMap ([ref]$nodeMap) `
                -NextCellId ([ref]$nextCellId) -DefinedNodes ([ref]$definedNodes)
            $branchCounter++
        }
    }

    $pageDims = Calculate-NodePositions -Nodes ([ref]$nodes)
    $diagramXml = Build-DiagramXml -Nodes $nodes -Edges $edges -NodeMap $nodeMap `
        -PageWidth $pageDims.PageWidth -PageHeight $pageDims.PageHeight
    return @{ Name = $aa.Name; DiagramXml = $diagramXml }
}

# ============================================================================
# MAIN SCRIPT
# ============================================================================

Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Teams Auto Attendant Call Flow Exporter"    -ForegroundColor Cyan
Write-Host " Draw.io Diagram Generator v1.4"             -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Host "[+] Created output directory: $OutputPath" -ForegroundColor Green
}

Write-Host "[1/5] Retrieving Auto Attendants..." -ForegroundColor Yellow
$autoAttendants = [System.Collections.Generic.List[object]]::new()
$skip = 0; $batchSize = 100
do {
    $batch = Get-CsAutoAttendant -IncludeStatus -First $batchSize -Skip $skip
    if ($batch) { foreach ($item in $batch) { [void]$autoAttendants.Add($item) } }
    $skip += $batchSize
} while ($batch -and $batch.Count -eq $batchSize)
Write-Host "       Found $($autoAttendants.Count) Auto Attendant(s)" -ForegroundColor Gray

Write-Host "[2/5] Retrieving Call Queues..." -ForegroundColor Yellow
$callQueues = [System.Collections.Generic.List[object]]::new()
$skip = 0
do {
    $batch = Get-CsCallQueue -First $batchSize -Skip $skip -WarningAction SilentlyContinue
    if ($batch) { foreach ($item in $batch) { [void]$callQueues.Add($item) } }
    $skip += $batchSize
} while ($batch -and $batch.Count -eq $batchSize)
Write-Host "       Found $($callQueues.Count) Call Queue(s)" -ForegroundColor Gray

Write-Host "[3/5] Retrieving Resource Accounts..." -ForegroundColor Yellow
$resourceAccounts = [System.Collections.Generic.List[object]]::new()
$skip = 0
do {
    $batch = Get-CsOnlineApplicationInstance -ResultSize $batchSize -Skip $skip
    if ($batch) { foreach ($item in $batch) { [void]$resourceAccounts.Add($item) } }
    $skip += $batchSize
} while ($batch -and $batch.Count -eq $batchSize)
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

Write-Host "       Lookup tables ready" -ForegroundColor Gray

Write-Host "[5/5] Generating draw.io diagrams..." -ForegroundColor Yellow

$allDiagrams = [System.Collections.Generic.List[hashtable]]::new()
$aaCounter = 0

foreach ($aa in $autoAttendants) {
    $aaCounter++
    Write-Host "       [$aaCounter/$($autoAttendants.Count)] Processing: $($aa.Name)" -ForegroundColor Gray

    $result = Export-AADiagram -AutoAttendant $aa `
        -ResourceAccountLookup $ResourceAccountLookup -AALookup $AALookup `
        -CQLookup $CQLookup -UserCache $UserCache -RAPhoneNumbers $RAPhoneNumbers

    $safeFileName = ($aa.Name -replace '[\\\/\:\*\?\"\<\>\|]', '_')
    $diagramId = Sanitise-NodeId $aa.Identity
    $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.000Z"

    $fileContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<mxfile host="app.diagrams.net" modified="$timestamp" agent="Teams Call Flow Export" version="21.0.0" type="device">
  <diagram id="$diagramId" name="$(Escape-XmlString $aa.Name)">
$($result.DiagramXml)
  </diagram>
</mxfile>
"@

    $filePath = Join-Path $OutputPath "$safeFileName.drawio"
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($filePath, $fileContent, $utf8NoBom)
    Write-Host "         Saved: $filePath" -ForegroundColor DarkGreen

    [void]$allDiagrams.Add($result)
}

$timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.000Z"
$combinedSb = [System.Text.StringBuilder]::new()
[void]$combinedSb.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
[void]$combinedSb.AppendLine("<mxfile host=""app.diagrams.net"" modified=""$timestamp"" agent=""Teams Call Flow Export"" version=""21.0.0"" type=""device"">")

$legendXml = Build-LegendPage -DiagramId "legend_page"
[void]$combinedSb.AppendLine($legendXml)

$pageCounter = 0
foreach ($diagram in $allDiagrams) {
    $pageCounter++
    $pageName = Escape-XmlString $diagram.Name
    [void]$combinedSb.AppendLine("  <diagram id=""page_$pageCounter"" name=""$pageName"">")
    [void]$combinedSb.AppendLine($diagram.DiagramXml)
    [void]$combinedSb.AppendLine("  </diagram>")
}

[void]$combinedSb.AppendLine("</mxfile>")

$combinedPath = Join-Path $OutputPath "_AllCallFlows.drawio"
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($combinedPath, $combinedSb.ToString(), $utf8NoBom)

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host " Export Complete!" -ForegroundColor Green
Write-Host " Files saved to: $OutputPath" -ForegroundColor Green
Write-Host " Individual files: $($autoAttendants.Count) Auto Attendant(s)" -ForegroundColor Green
Write-Host " Combined file:   _AllCallFlows.drawio" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "TIP: Open the .drawio files in draw.io Desktop, diagrams.net," -ForegroundColor DarkYellow
Write-Host "     or VS Code with the draw.io extension to view and edit" -ForegroundColor DarkYellow
Write-Host "     the diagrams interactively." -ForegroundColor DarkYellow