# Outlook PST Mover Add-in

## Overview

The Outlook PST Mover Add-in allows you to move or copy items from selected folders and their subfolders to a specified PST file. It ensures all items, along with their metadata, are transferred accurately. The add-in is designed to handle exceptions gracefully, ensuring maximum items are processed even if some fail.

## Features

- Move or copy items from selected top-level folders and their subfolders to a PST file.
- Automatically creates the destination folder structure in the PST file.
- Skips source items when the destination already has an equally current copy; in move mode, skipped source items are left in place.
- Skips mail items that require Information Rights Management permissions and reports that count separately.
- Replaces older destination copies only after the source item has been written successfully, preventing duplicate copies without risking the existing destination item.
- Exports Calendar items to an iCalendar (`.ics`) file with category, appointment/meeting, and organizer filters.
- Exports Contacts items to a vCard (`.vcf`) file that can be imported into Google Contacts.
- Handles items that cannot be moved or copied by skipping them.
- Ensures all items are fully downloaded before processing.
- Displays a progress bar with statistics:
  - Items processed.
  - Time elapsed.
  - Estimated time remaining.

## Prerequisites

- Microsoft Outlook 2013 or later.
- .NET Framework 4.6.1 or later.
- Visual Studio 2017 or later with VSTO (Visual Studio Tools for Office) installed.

## Installation

1. **Download the Project**

   Clone or download the repository to your local machine.

2. **Build the Project**

   - Open the solution file `OutlookPstMover.sln` in Visual Studio.
   - Build the solution by selecting `Build > Build Solution` from the menu.

3. **Install the Add-in**

   - After a successful build, navigate to the `bin\Debug` or `bin\Release` folder of the project.
   - Run the `OutlookPstMover.vsto` installer file to install the add-in.

## Usage

1. **Launch Outlook**

   Open Microsoft Outlook. You should see a new tab called **OutBack** in the ribbon.

2. **Select a Folder**

   In the Outlook navigation pane, select the folder whose items you want to move or copy.

3. **Start the Operation**

   - Click on the **OutBack** tab.
   - Click the **Move/Copy** button.

4. **Configure Options**

   - In the dialog that appears, select a PST file and choose the source folders to process. Subfolders are included automatically.
   - Check the **Move Items** checkbox if you wish to move items. Leave it unchecked to copy items.
   - Click **Start** to begin the operation.

5. **Monitor Progress**

   - A progress window will display the number of items processed, time elapsed, and estimated time remaining.
   - Upon completion, a success message will appear.

6. **Export Contacts for Google Contacts**

   - Select a Contacts folder, or let the add-in use your default Contacts folder.
   - Click **Export Contacts**.
   - Save the `.vcf` file and import it from Google Contacts.

## How to Debug

This is a VSTO (Visual Studio Tools for Office) Outlook Add-in. Debugging requires Visual Studio to launch Outlook as the host application and attach its debugger.

### Prerequisites

- Visual Studio 2017 or later (Professional or higher recommended for full VSTO debugging support).
- Microsoft Outlook installed locally.
- The VSTO workload installed in Visual Studio (**Office/SharePoint development** workload via the Visual Studio Installer).

### Starting a Debug Session

1. Open `OutBack.sln` in Visual Studio.
2. Ensure the **Debug** configuration is selected in the toolbar (not Release).
3. Set breakpoints in any source file (e.g., `ItemMover.cs`, `MyRibbon.cs`, `PSTSelectionForm.cs`).
4. Press **F5** (or **Debug > Start Debugging**).
   - Visual Studio will build the add-in, register it temporarily, and launch Outlook automatically.
   - The debugger attaches to the Outlook process, and your breakpoints will be hit when the corresponding code executes.

### Useful Breakpoint Locations

| File | Where to Break | Why |
|------|----------------|-----|
| `MyRibbon.cs` | Button click handler | Entry point when the user clicks a ribbon command |
| `PSTSelectionForm.cs` | Form load / Start button | Inspect user-selected options before the operation begins |
| `ItemMover.cs` | Move/copy loop | Step through individual item processing |
| `ProgressForm.cs` | Progress update calls | Verify progress tracking and ETA calculations |
| `ThisAddIn.cs` | `ThisAddIn_Startup` | Verify the add-in initializes correctly |

### Debugging Techniques

- **Breakpoints (F9):** Click the left margin of a code line or press F9 to toggle a breakpoint. Execution will pause when that line is reached.
- **Conditional Breakpoints:** Right-click a breakpoint and choose **Conditions** to break only when a specific expression is true (e.g., `itemIndex > 100`).
- **Watch & Locals Windows:** Use **Debug > Windows > Watch** or **Locals** to inspect variable values while paused at a breakpoint.
- **Immediate Window (Ctrl+Alt+I):** Evaluate expressions and call methods at runtime while paused.
- **Output Window:** View debug trace messages in **View > Output** (select **Debug** in the "Show output from" dropdown).
- **Exception Settings (Ctrl+Alt+E):** Configure which exceptions break into the debugger. Enable **Common Language Runtime Exceptions** to catch all managed exceptions.

### Debugging COM / Outlook Interop Issues

- If Outlook items throw `COMException` or `RPC_E_CALL_REJECTED` errors, these are common in Office interop. Set the debugger to break on `System.Runtime.InteropServices.COMException` via **Debug > Windows > Exception Settings**.
- Use the **Immediate Window** to inspect Outlook object properties (e.g., `item.Subject`, `item.MessageClass`) while paused.

### Stopping a Debug Session

- Press **Shift+F5** (or **Debug > Stop Debugging**) to end the session. Visual Studio will close Outlook and unregister the temporary add-in.
- Alternatively, closing Outlook manually will also end the debug session.

### Troubleshooting

| Problem | Solution |
|---------|----------|
| Outlook does not launch on F5 | Verify Outlook is installed and the `OfficeApplication` property in the `.csproj` is set to `Outlook`. Check **Tools > Options > Office Tools > Project Debugging** settings. |
| Breakpoints show "No symbols loaded" | Ensure you are using the **Debug** configuration. Do a **Build > Clean Solution** followed by **Build > Rebuild Solution**. |
| Add-in does not appear in Outlook | Check **File > Options > Add-ins** in Outlook. If it is listed under **Disabled**, re-enable it. Also verify VSTO Runtime 4.0 is installed. |
| "The assembly could not be loaded" error | Run Visual Studio as Administrator, or re-sign the manifest by right-clicking the project > **Properties > Signing**. |

## Handling Errors

- The add-in is designed to skip items that cannot be moved or copied.
- In case of any unexpected errors, an error message will be displayed, and the operation will continue with the remaining items.

## Uninstallation

- Go to `Control Panel > Programs and Features`.
- Find **Outlook PST Mover Add-in** in the list.
- Right-click and select **Uninstall**.

## Support

For any issues or questions, please contact [Your Name] at [your.email@example.com].

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
