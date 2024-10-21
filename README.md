# Outlook PST Mover Add-in

## Overview

The Outlook PST Mover Add-in allows you to move or copy all items from the currently selected folder to a specified PST file. It ensures all items, along with their metadata, are transferred accurately. The add-in is designed to handle exceptions gracefully, ensuring maximum items are processed even if some fail.

## Features

- Move or copy items from the current folder to a PST file.
- Automatically creates the destination folder structure in the PST file.
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

   Open Microsoft Outlook. You should see a new tab called **PST Mover** in the ribbon.

2. **Select a Folder**

   In the Outlook navigation pane, select the folder whose items you want to move or copy.

3. **Start the Operation**

   - Click on the **PST Mover** tab.
   - Click the **Move/Copy to PST** button.

4. **Configure Options**

   - In the dialog that appears, click **Browse...** to select or create a PST file.
   - Check the **Move Items** checkbox if you wish to move items. Leave it unchecked to copy items.
   - Click **Start** to begin the operation.

5. **Monitor Progress**

   - A progress window will display the number of items processed, time elapsed, and estimated time remaining.
   - Upon completion, a success message will appear.

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
