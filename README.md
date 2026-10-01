# Teams Call Flow Exporter to Draw.io

A powerful, robust PowerShell utility designed to connect to your Microsoft Teams tenant, discover all **Auto Attendants (AA)** and **Call Queues (CQ)**, resolve their configurations, and automatically construct clean, interactive, and perfectly-aligned **Draw.io (.drawio) diagrams**.

This tool solves the challenge of manually documenting complex cloud telephony setups by instantly exporting fully structured call flows. Each Auto Attendant receives its own diagram page, and a unified `_AllCallFlows.drawio` file is created containing a dedicated page per AA, complete with an interactive visual **Legend**.

---

## 🌟 Key Features

- **End-to-End Call Path Analysis:** Maps entire call routing including Business Hours, After Hours, Holiday calendars (with the actual holiday date ranges), IVR key presses (DTMF and voice triggers, including "transfer to operator" keys resolved to the Auto Attendant's operator and announcement keys), Call Queue agent counts, routing methods, and all three queue exception rules (timeout, overflow, no agents), including for queues nested behind other queues.
- **Dynamic Overlap-Free Layout:** Powered by an advanced, boundary-aware layout engine that calculates the precise footprint of nodes (including whole nested-queue subtrees, queue settings notes, schedule notes, and speech greetings), completely eliminating horizontal node collisions or overlapping branches. Multiple IVR keys that route to the same target are combined onto a single connector (e.g. `Press 2, Press 3`) rather than stacked on top of each other.
- **Selectable Colour Themes:** Ships with three palettes selected via `-StylePreset`: `Default` (vibrant Microsoft-themed), `HighContrast` (bold colours and thick borders for accessibility / projectors), and `Monochrome` (greyscale for black-and-white printing). Segoe UI typography, HSL-harmonized fills, clean rounded borders, and distinct shapes throughout.
- **Rich Greeting Note Integration:** Automatically extracts Text-to-Speech (TTS) prompts or audio greeting filenames and renders them as elegant, clean sticky notes floating next to their respective menu nodes.
- **Queue Settings Notes:** Each Call Queue gets a note beside it summarising where its agents come from (users / groups / Teams channel), agent alert time, presence-based routing, agent opt-out, conference mode, callback (key and conditions), greeting, and music on hold. Turn it off with `-HideQueueSettings` for a more compact diagram.
- **Clickable Nested Auto Attendants:** In `_AllCallFlows.drawio`, an Auto Attendant that routes to another Auto Attendant links to that AA's page, and every row on the Index page links to its diagram.
- **Integrated Schedule Panels:** Decodes complex weekly recurrent business hour schedules and displays them as clean, HTML-formatted calendar panels next to the root Auto Attendant.
- **Robust Target Resolution:** Intelligently maps targets (Application Accounts, Users, External PSTN numbers, shared voicemail — labelled with the Microsoft 365 group name when a Graph session is available — and a user's personal voicemail). Fallback logic ensures that even voice apps with missing/unlisted Resource Accounts are resolved via their `ApplicationInstances` across the tenant.
- **Fault-Tolerant Export:** A malformed Auto Attendant is logged and skipped rather than aborting the whole run, and an `_ExportSummary.json` records per-diagram statistics and any warnings/failures.

---

## 🎨 Diagram Styles & Legend

The exporter uses a clear, highly legible color hierarchy:

| Node Type | Shape | Color Theme | Description |
| :--- | :--- | :--- | :--- |
| **Auto Attendant (AA)** | Rounded Rectangle | Blue (`#4472C4`) | Main entry point or nested Auto Attendant |
| **Call Queue (CQ)** | Pill Rectangle | Green (`#548235`) | Routing container showing agent count & method |
| **Business Hours Menu** | Rhombus | Yellow (`#FFC000`) | IVR key options during business hours |
| **After Hours Menu** | Rhombus | Dark Blue (`#2E75B6`) | IVR key options during closed hours |
| **Holiday Menu** | Hexagon | Gold (`#BF8F00`) | IVR / routing during a holiday; the connector shows the holiday date(s) |
| **User** | Rounded Rectangle | Purple (`#7030A0`) | Call routed directly to a Teams user |
| **External PSTN** | Rounded Rectangle | Orange (`#ED7D31`) | Call routed out to an external phone number |
| **Voicemail** | Parallelogram | Gray (`#A5A5A5`) | Shared voicemail (with group name when available) or a user's personal voicemail |
| **Timeout / Overflow / No Agents** | Parallelogram | Light Orange (`#ED7D31`) | Queue exception handling: timeout, overflow cap, or no agents available |
| **Queue Settings** | Sticky Note | Light Green (`#E2EFDA`) | Agent sources, routing options, callback, greeting and music on hold for a Call Queue |
| **Greeting Note** | Sticky Note | Light Yellow (`#FFF2CC`) | Renders Text-to-Speech or audio greeting file details |
| **Schedule Panel** | Sticky Note | Light Blue (`#DAE8FC`) | Renders HTML calendar hours next to AA roots |
| **Disconnect** | Ellipse | Dark Red (`#C00000`) | Call termination points |
| **Announcement** | Dashed Rounded Rectangle | Light Orange (`#FCE4D6`) | A menu key that plays a message (TTS text or audio file name shown), then repeats the menu |

---

## 🛠️ Prerequisites

To run the script, ensure you have:

1. **PowerShell:** Windows PowerShell 5.1 or PowerShell Core (7+).
2. **Microsoft Teams PowerShell Module:** Installed on your machine.
   ```powershell
   Install-Module -Name MicrosoftTeams -Force -AllowClobber
   ```
3. **Privileges:** An account with at least **Teams Administrator** or **Teams Communications Support Engineer** (Reader) permissions.
4. **Active Session:** You must connect to Teams PowerShell prior to running the script.
   ```powershell
   Connect-MicrosoftTeams
   ```
5. **Optional — shared voicemail names:** Teams only returns the Microsoft 365 group ID for a shared voicemail target. If you also have the `Microsoft.Graph` module and connect with group read access before running, the diagrams show the group's name (e.g. `Shared Voicemail: Newcastle VM`). Without it — or without the permission — they just say `Shared Voicemail`.
   ```powershell
   Connect-MgGraph -Scopes Group.Read.All
   ```

---

## 🚀 Usage

### Simple Execution
Run the script to export all call flows to the default subdirectory (`.\CallFlowDiagrams`):
```powershell
.\Export-TeamsCallFlowDrawIO.ps1
```

### Custom Output Folder
Specify a custom directory for saving the generated files:
```powershell
.\Export-TeamsCallFlowDrawIO.ps1 -OutputPath "C:\Contoso\TelephonyDocs"
```

### Colour Theme
Choose a palette with `-StylePreset` (`Default`, `HighContrast`, or `Monochrome`):
```powershell
.\Export-TeamsCallFlowDrawIO.ps1 -StylePreset HighContrast
.\Export-TeamsCallFlowDrawIO.ps1 -StylePreset Monochrome -OutputPath .\Printable
```

### Compact Diagrams
Leave out the per-queue settings notes:
```powershell
.\Export-TeamsCallFlowDrawIO.ps1 -HideQueueSettings
```

---

## 📂 Outputs Generated

Inside your output folder, the script will generate:

1. **Individual `.drawio` files** (e.g. `Main_Line_AA.drawio`): A clean, single-page flow dedicated to each discovered Auto Attendant.
2. **`_AllCallFlows.drawio`**: A master drawing sheet. When opened, it displays:
   - An **Index** cover page — an alphabetised directory of every Auto Attendant with its phone number(s); click a row to jump to that diagram.
   - A visual **Legend** page detailing all node colors, shapes, and edge connector patterns.
   - Separate, named, high-fidelity **tabbed pages** for every Auto Attendant in your organization. Nested Auto Attendant nodes link to their own page.
3. **`_ExportSummary.json`**: A machine-readable run summary — per-diagram node/edge counts, file sizes, style preset, total duration, and counts of diagrams with validation warnings or that failed to export.

---

## 🖥️ How to View and Edit Diagrams

The generated `.drawio` files are standard XML and can be opened in several ways:
- **Draw.io Desktop App:** The recommended native application for Windows/macOS/Linux.
- **Web-based editor:** Visit [draw.io](https://app.diagrams.net/) (diagrams.net) and select "Open Existing Diagram".
- **VS Code:** Install the **Draw.io Integration** extension by Henning Dieterichs to view, edit, and export your telephony diagrams directly inside VS Code without leaving your editor.

---

## ⚙️ Layout Engine Logic (Behind the Scenes)

The script features a custom-built, tier-based hierarchical layout engine.
- **Tier 0:** Auto Attendant Root & Schedule Panel.
- **Tier 1:** Main Call Flows (Business Hours, After Hours, Holiday paths).
- **Tier 2:** IVR Menus, Menu Options, Call Queues, and Direct Targets.
- **Tier 3+:** Queue exception handling — Timeout, Overflow and No Agents — and their targets, hung below the parent CQ. Nested queues (e.g. a queue that overflows to another queue) get their own exception rules drawn too, however deep the chain goes, with room reserved so nothing overlaps.

Rather than placing nodes blindly or capping branch widths arbitrarily, the algorithm computes **boundary footprints** for each branch. By accounting for attached notes (greetings on the right, queue settings on the left) and the full width of every queue's exception subtree, it reserves exactly enough space horizontally and vertically, assuring **100% collision-free diagrams** regardless of your configuration's complexity.
