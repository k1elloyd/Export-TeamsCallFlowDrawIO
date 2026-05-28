# Teams Call Flow Exporter to Draw.io

A powerful, robust PowerShell utility designed to connect to your Microsoft Teams tenant, discover all **Auto Attendants (AA)** and **Call Queues (CQ)**, resolve their configurations, and automatically construct clean, interactive, and perfectly-aligned **Draw.io (.drawio) diagrams**.

This tool solves the challenge of manually documenting complex cloud telephony setups by instantly exporting fully structured call flows. Each Auto Attendant receives its own diagram page, and a unified `_AllCallFlows.drawio` file is created containing a dedicated page per AA, complete with an interactive visual **Legend**.

---

## 🌟 Key Features

- **End-to-End Call Path Analysis:** Maps entire call routing including Business Hours, After Hours, Holiday calendars, IVR key presses (DTMF and voice triggers), Call Queue agent counts, routing methods, timeouts, and overflow rules.
- **Dynamic Overlap-Free Layout:** Powered by an advanced, boundary-aware layout engine that calculates the precise footprint of nodes (including shifted timeout/overflow children, schedule notes, and speech greetings), completely eliminating horizontal node collisions or overlapping branches.
- **Vibrant Modern Styling:** Employs premium, professional color-coded themes (Segoe UI typography, HSL-harmonized fills, clean rounded borders, and distinct shapes) for instant visual scanning.
- **Rich Greeting Note Integration:** Automatically extracts Text-to-Speech (TTS) prompts or audio greeting filenames and renders them as elegant, clean sticky notes floating next to their respective menu nodes.
- **Integrated Schedule Panels:** Decodes complex weekly recurrent business hour schedules and displays them as clean, HTML-formatted calendar panels next to the root Auto Attendant.
- **Robust Target Resolution:** Intelligently maps targets (Application Accounts, Users, External PSTN numbers, and Shared Voicemails). Fallback logic ensures that even voice apps with missing/unlisted Resource Accounts are resolved by scanning `ApplicationInstances` across the tenant.

---

## 🎨 Diagram Styles & Legend

The exporter uses a clear, highly legible color hierarchy:

| Node Type | Shape | Color Theme | Description |
| :--- | :--- | :--- | :--- |
| **Auto Attendant (AA)** | Rounded Rectangle | Blue (`#4472C4`) | Main entry point or nested Auto Attendant |
| **Call Queue (CQ)** | Pill Rectangle | Green (`#548235`) | Routing container showing agent count & method |
| **Business Hours Menu** | Rhombus | Yellow (`#FFC000`) | IVR key options during business hours |
| **After Hours Menu** | Rhombus | Dark Blue (`#2E75B6`) | IVR key options during closed hours |
| **User** | Rounded Rectangle | Purple (`#7030A0`) | Call routed directly to a Teams user |
| **External PSTN** | Rounded Rectangle | Orange (`#ED7D31`) | Call routed out to an external phone number |
| **Shared Voicemail** | Parallelogram | Gray (`#A5A5A5`) | Call routed to a shared voicemail inbox |
| **Timeout / Overflow** | Parallelogram | Light Orange (`#ED7D31`) | Queue routing action on timeout or cap limit |
| **Greeting Note** | Sticky Note | Light Yellow (`#FFF2CC`) | Renders Text-to-Speech or audio greeting file details |
| **Schedule Panel** | Sticky Note | Light Blue (`#DAE8FC`) | Renders HTML calendar hours next to AA roots |
| **Disconnect** | Ellipse | Dark Red (`#C00000`) | Call termination points |

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

---

## 🚀 Usage

### Simple Execution
Run the script to export all call flows to the default subdirectory (`.\CallFlowDiagrams`):
```powershell
.\Export-TeamsCallFlowDrawIO_1.3.ps1
```

### Custom Output Folder
Specify a custom directory for saving the generated files:
```powershell
.\Export-TeamsCallFlowDrawIO_1.3.ps1 -OutputPath "C:\Contoso\TelephonyDocs"
```

---

## 📂 Outputs Generated

Inside your output folder, the script will generate:

1. **Individual `.drawio` files** (e.g. `Main_Line_AA.drawio`): A clean, single-page flow dedicated to each discovered Auto Attendant.
2. **`_AllCallFlows.drawio`**: A master drawing sheet. When opened, it displays:
   - A visual **Legend** page detailing all node colors, shapes, and edge connector patterns.
   - Separate, named, high-fidelity **tabbed pages** for every Auto Attendant in your organization.

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
- **Tier 3:** Queue Timeout and Overflow targets (positioned cleanly below the parent CQ).

Rather than placing nodes blindly or capping branch widths arbitrarily, the algorithm computes **boundary footprints** for each branch. By accounting for the size of attached speech greeting notes (+295px) and left/right shifts of Queue Timeout (-90px) and Overflow (+90px) actions, it reserves exactly enough space horizontally and vertically, assuring **100% collision-free diagrams** regardless of your configuration's complexity.
