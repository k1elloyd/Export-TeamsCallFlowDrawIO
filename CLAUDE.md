# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single self-contained PowerShell script, [Export-TeamsCallFlowDrawIO.ps1](Export-TeamsCallFlowDrawIO.ps1), that connects to a live Microsoft Teams tenant, discovers all Auto Attendants (AA) and Call Queues (CQ), and generates draw.io (`.drawio`) diagrams visualizing the call flow (business hours / after hours / holiday routing, IVR menus, queue timeout/overflow, nested AAs, TTS/audio greetings, schedules). There is no build system, package manifest, or test suite — it's one ~1600 line script plus a `CallFlowDiagrams/` folder of example/generated output.

## Running the script

Requires the `MicrosoftTeams` PowerShell module and an active `Connect-MicrosoftTeams` session (the script checks this via `Get-CsTenant` at startup and exits if not connected — there is no offline/mock mode, so most of the script cannot be exercised without live tenant access).

Target runtime is **Windows PowerShell 5.1** (what most Teams admins run), not just PowerShell 7. Several cmdlets treat `[hashtable]` lists differently between the two engines — see the hashtable-adapter note under "Key invariants" — so validate changes against 5.1 (`powershell.exe`), not only `pwsh`.

```powershell
Install-Module -Name MicrosoftTeams -Force -AllowClobber
Connect-MicrosoftTeams

.\Export-TeamsCallFlowDrawIO.ps1
.\Export-TeamsCallFlowDrawIO.ps1 -OutputPath "C:\Contoso\TelephonyDocs"
.\Export-TeamsCallFlowDrawIO.ps1 -StylePreset HighContrast   # or Monochrome
```

Syntax/parse-check without a Teams connection:
```powershell
powershell -NoProfile -Command "[System.Management.Automation.PSParser]::Tokenize((Get-Content -Raw .\Export-TeamsCallFlowDrawIO.ps1), [ref]$null) | Out-Null"
```
(or simply `Get-Command -Syntax .\Export-TeamsCallFlowDrawIO.ps1` to validate `param()` block parsing).

There are no linter/formatter configs or automated tests in this repo — validate changes by parsing the script and, where possible, running it against a real/test tenant and opening the resulting `.drawio` output in draw.io or VS Code's Draw.io Integration extension.

## Architecture

Script params: `-OutputPath`, `-StylePreset`, `-HideQueueSettings` (omit the per-CQ settings note). Optional Microsoft Graph integration: if `Get-MgContext` returns a session at startup, `$script:GraphGroupLookup` is enabled and `Get-SharedVoicemailLabel` resolves shared-voicemail group IDs to names via `Get-MgGroup` (cached in `$script:GroupNameCache`; a permissions error switches lookups off for the rest of the run). It never connects to Graph itself.

The script is organized top-to-bottom as one linear pipeline; there are no modules/classes, everything is functions plus a "MAIN SCRIPT" section at the bottom that calls them in sequence.

1. **Style tables** (top of file, driven by `-StylePreset`): `$NodeStyles`, `$EdgeStyles`, `$NodeSizes` are hashtables keyed by node type (`AA`, `CQ`, `Menu`, `MenuAfterHours`, `User`, `ExternalPstn`, `SharedVoicemail`, `Disconnect`, `Holiday`, `TimeoutOverflow`, `Title`, `Greeting`, `Schedule`). Adding a new node type means adding an entry to all three tables (for all three presets), plus the legend in `Build-LegendPage`. Note types (`Greeting`, `Schedule`, `QueueSettings`) are also excluded by name from the tier/footprint/hanging classification in `Calculate-NodePositions`, so a new note type must be added to those lists too.

2. **Data retrieval** (MAIN SCRIPT, `[1/5]`–`[4/5]`): paginated `Get-CsAutoAttendant`, `Get-CsCallQueue`, `Get-CsOnlineApplicationInstance` calls, then builds lookup hashtables (`$AALookup`, `$CQLookup`, `$ResourceAccountLookup`, `$RAPhoneNumbers`) keyed by identity/ObjectId for O(1) resolution during graph building.

3. **Target resolution** — `Resolve-CallTarget`: given a Teams call-flow action's target, determines whether it points to another AA, a CQ, a user, external PSTN, shared voicemail, or disconnect. Resource accounts are matched by `ApplicationId` GUID (`ce933385-...` = Auto Attendant app, `11cd3e2e-...` = Call Queue app) with fallback lookup via `$AAByAppInstance`/`$CQByAppInstance` (ApplicationInstance ID → AA/CQ, built once in MAIN SCRIPT) when the resource account lookup misses — this fallback exists specifically to handle voice apps with missing/unlisted resource accounts, and is O(1) rather than scanning every AA/CQ.

4. **Graph building** — `Build-CallFlowNodes` (walks an AA's `CallFlow` object: menu options, DTMF/voice triggers, targets; a `TransferCallToOperator` option normally has no `CallTarget` of its own, so it falls back to the AA-level `Operator` passed in via `-Operator`) and `Build-CallQueueNodes` (queue agents, routing method, a `QueueSettings` note from `Get-QueueSettingsSummary` placed to the queue's left unless `-HideQueueSettings`, and the three exception rules: timeout, overflow, no-agents — a no-agents action of `Queue` is the default and isn't drawn; `Disconnect`/`DisconnectWithBusy` become terminal Disconnect nodes; action `Voicemail` with a `User` target is that user's *personal* voicemail, distinct from `SharedVoicemail`) populate flat `$nodes` / `$edges` lists via `Add-DiagramNode` / `Add-DiagramEdge`. Both delegate per-target node creation to `Resolve-AndAddTargetNode`, a shared helper that takes a `Resolve-CallTarget` result and ensures the right node exists (AA/CQ/User/ExternalPstn/SharedVoicemail/Disconnect/Unknown), recursing into `Build-CallQueueNodes` for CQ targets so nested timeout/overflow chains get built too. This is the single place that maps target type → node type/label/ID convention — all four call sites (menu options, no-menu default action, CQ timeout, CQ overflow) route through it, so adding a new target type only needs to change one function. Nodes carry logical placement metadata (`Tier`, `BranchIndex`, `PositionInBranch`, `ParentNodeId`) but no coordinates yet — `Add-DiagramNode` dedupes by `NodeId`.

   Dedup is branch-scoped, not global: `Resolve-AndAddTargetNode` keys AA/CQ/User target node IDs by `(LinkedId, BranchIndex)` (e.g. `CQ_<identity>_b0`), so the same target reached twice *within* one branch still converges onto a single node, but each branch (Business Hours/After Hours/Holiday) draws its own copy — no edge crosses between branches to reach a shared target. The one exception is a target that routes back to the AA's own root (e.g. a "press 9 for the main menu" loop): that's kept unsuffixed via the `CurrentAAIdentity`/`AAIdentity` parameter threaded through `Build-CallFlowNodes` → `Resolve-AndAddTargetNode` → `Build-CallQueueNodes`, so it resolves to the existing root node instead of spawning a duplicate root per branch.

5. **Layout engine** — `Calculate-NodePositions`: a custom tier-based hierarchical layout (Tier 0 = AA root + schedule panel, Tier 1 = Business Hours/After Hours/Holiday flow roots, Tier 2 = IVR menus/options/CQs/direct targets). Everything below a CQ — its `_timeout`/`_overflow`/`_noagent` exception children and whatever they route to, including nested CQs with their own exceptions, to any depth — is a "hanging" subtree, not a tier row: `Measure-Subtree` recursively computes each subtree's left/right extent (including the node's own notes: queue settings on the left, greeting on the right) (children side by side in `PositionInBranch` order, centred under the parent) and `Place-Subtree` positions it below its root once tiers 0–2 are placed. Each tier 0–2 node's "footprint" (left/right pixel offsets) is its subtree extent plus any attached greeting note, which reserves exactly enough horizontal space to guarantee collision-free output regardless of call-flow complexity. The schedule panel is positioned after subtrees so it clears them. It treats every distinct `NodeId` as an independent node purely by `Tier`/`BranchIndex`/`PositionInBranch`/`ParentNodeId` — it has no notion of "the same underlying AA/CQ/User," so the branch-scoped duplication in step 4 needs no special handling here. Builds `$nodeById`/`$greetingByParent` index hashtables once up front; by-ID and by-parent lookups should go through those rather than re-scanning `$Nodes.Value` with `Where-Object`. This is the most complex/fragile part of the script — changes here have wide blast radius across every diagram shape.

6. **XML generation** — `Build-DiagramXml` renders the positioned nodes/edges into mxGraph XML, with every vertex going through `Get-VertexXml`. Edges added with `-RouteDown` (all queue exception edges) get fixed bottom-exit/top-entry points and two waypoints putting the horizontal jog 20px above the target, computed here from laid-out positions, so they pass below the queue settings note rather than through it; if the target isn't below the source, draw.io's auto-routing is kept. A node with a `Link` (optional `-Link`/`-Tooltip` on `Add-DiagramNode`) is wrapped in a `<UserObject>`, which is how draw.io stores links. Nested-AA target nodes and Index page rows link to `data:page/id,page_<Sanitise-NodeId identity>`, which is the page id each AA gets in `_AllCallFlows.drawio` (`PageId` in the `Export-AADiagram` result) — keep those two in sync; `Build-LegendPage` generates a standalone legend page reused across all diagrams.

7. **Validation** — `Test-DiagramIntegrity` sanity-checks the generated node/edge graph (e.g. dangling edges) and flags results as `IsValid`; the main loop tracks a `$warnedCount` across all AAs and reports it in the final summary.

8. **Orchestration** — `Export-AADiagram` ties one AA's title node, root node, schedule panel, and all three call-flow branches (business hours/after hours/holiday) together, then calls layout + XML generation + validation and returns a result hashtable (`Name`, `DiagramXml`, `NodeCount`, `EdgeCount`, `IsValid`, `PhoneNumbers`). The MAIN SCRIPT section calls this once per AA, writes one `.drawio` file per AA, then assembles all of them plus an index page (`Build-IndexPage`, an alphabetised AA/phone-number directory — the first page a reader sees) and the legend into a combined `_AllCallFlows.drawio` (one page per AA) and an `_ExportSummary.json` (per-diagram node/edge counts, file sizes, validation status).

Error handling policy: the three tenant data-retrieval loops (step 2) abort the whole script with a clear error if the underlying cmdlet fails — partial AA/CQ/resource-account data would silently produce an incomplete diagram set, so it fails fast instead. Per-AA export (step 8) is isolated in its own try/catch — one malformed AA is logged and skipped (counted in `$failedCount` / `DiagramsFailed` in the summary) rather than aborting the whole run.

### Key invariants to preserve when editing

- Every node type used anywhere in the graph-building functions must have a matching key in `$NodeStyles` and `$NodeSizes` (all three style presets), or it silently falls back to the `AA` style/a generic 200x60 size.
- `Add-DiagramNode` is idempotent per `NodeId` — call sites rely on this for dedup; don't bypass it with direct `$nodes.Add(...)`.
- Node IDs are sanitized via `Sanitise-NodeId` and suffixed conventions matter: `_timeout` / `_overflow` / `_noagent` suffixes on CQ exception child nodes are pattern-matched (`_(timeout|overflow|noagent)$`) by the layout engine to classify them as hanging subtree nodes — don't rename these without updating `Calculate-NodePositions`. Similarly, the `_b$BranchIndex` suffix on AA/CQ/User target NodeIds (see step 4) is how branch-scoped dedup works — don't strip or bypass it when adding new call sites into `Resolve-AndAddTargetNode`.
- All XML text content must go through `Escape-XmlString`.
- Nodes/edges are `[hashtable]`, not `[PSCustomObject]`. Windows PowerShell 5.1's hashtable adapter does **not** expose keys as member properties to several pipeline cmdlets, so operations that work in PowerShell 7 silently misbehave in 5.1 (the shipping target). Confirmed cases, all worked around in the code — do not reintroduce these forms on hashtable-list data:
  - `Sort-Object <BarePropertyName>` → sorts by the wrong order. Use `Sort-Object { $_.Prop }` (`Calculate-NodePositions`, `Build-IndexPage`).
  - `Measure-Object -Property <Name>` → throws `GenericMeasurePropertyNotFound`. Sum manually with a `foreach` loop (`_ExportSummary.json` totals in MAIN SCRIPT).
  - `[HashSet[string]]($hashtable.Keys)` → collapses all keys into a single space-joined string instead of a set of keys, so `.Contains(<key>)` is always false. Use `$hashtable.ContainsKey(...)` directly (`Test-DiagramIntegrity`).
