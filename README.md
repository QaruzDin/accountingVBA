# DelCellAutomationV1
## Overview
DelCellAutomationV1 is a VBA (Visual Basic for Applications) script designed to automate the deletion of specific cells in an Excel worksheet based on certain conditions. This script is particularly useful for cleaning up datasets where certain columns need to be removed under specific circumstances without deleting the entire column.

## Features
- Automated Deletion: The script can automatically delete cells based on predefined conditions.
- Backup: Before making any changes, the script creates a backup of the original worksheet.
- User Confirmation: Prompts the user for confirmation before executing the deletion process.
## Requirements
- Microsoft Excel
- Basic knowledge of VBA to run the script
## Installation
1. Open your Excel workbook.
2. Press Alt + F11 to open the VBA editor.
3. Insert a new module by right-clicking on any existing module or the workbook name in the Project Explorer and selecting Insert > Module.
4. Copy and paste the contents of DelCellAutomationV1.bas into the new module.
5. Save the workbook as a macro-enabled workbook (.xlsm).
## Usage
1. Open the Excel workbook containing the dataset you wish to clean.
2. Press Alt + F8 to open the Macro dialog box.
3. Select Del_Cell_Automation from the list of macros and click Run.
4. Confirm the prompt to continue with the deletion process.
## Script Details
The script works as follows:

- Prompts the user for confirmation before proceeding.
- Creates a backup of the original worksheet.
- Identifies the active worksheet and the range of cells to process.
- Deletes cells based on predefined conditions (the conditions need to be defined within the script).
- The user can customize the conditions within the script to meet specific needs.
## Customization
To customize the script to fit your specific requirements:

1. Open the VBA editor (Alt + F11).
2. Locate the Del_Cell_Automation macro in the DelCellAutomationV1 module.
3. Modify the conditions within the script to suit your dataset. For example, adjust the ranges or criteria for deleting cells.

## Author
The script was created by Qaruz Din (Andi Artsam)

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request or open an Issue if you find a bug or have a feature request.
