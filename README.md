# Formula_Updater

## Overview
The **GSHEET Formula Beautifier** is a Google Sheets add-on that enhances the readability of formulas by displaying them in a multi-line format with added comments. This helps users better understand and debug complex formulas.

</br>
<img src="https://github.com/adegard/Formula_Updater/blob/main/Immagine 2025-06-09 141420.jpg"  align="center">

## Features
- **Formula Updater**: Opens a dialog box where users can visualize and modify formulas.
- **Retrieve Current Formula**: Extracts the formula from the selected cell.
- **Beautify Formula**: Formats formulas by structuring them into a readable, multi-line format.
- **Minify Formula**: Condenses formulas by removing unnecessary spaces.
- **Save Formula with Comments**: Allows users to annotate their formulas with comments and store backups.

## Installation & Usage
1. Install the add-on from the Google Sheets extension marketplace.
2. Upon installation, the add-on will integrate into the Google Sheets menu.
3. Navigate to the **Formula Updater** option to view and edit formulas.
4. Utilize the beautify or minify functions to reformat formulas.
5. Save your work with comments to track formula history.

## Code Architecture
The add-on is built using Google Apps Script and operates directly within Google Sheets. It uses UI dialogs to enhance user interaction and allows efficient formula modification.

### Core Functions:
- `getFormula()`: Retrieves the current formula.
- `setFormula(formula)`: Updates the cell with the formatted formula.
- `beautify()`: Structures formulas into a multi-line format.
- `minify()`: Condenses formulas for compact representation.
- `saveMyFormula(formula, comments)`: Saves formulas along with comments for easy tracking.

## Benefits
- Improved formula readability.
- Easier debugging and modification.
- Formula backup with annotation capabilities.

## License
This add-on is open-source and free to use under the MIT License.

## Additional Resources
For Notepad++ users interested in updating google sheet formulas: visit the [Google sheet Formula UDL (User Defined Language) for Notepad++](https://github.com/adegard/gsheet_notepad-plus-plus/tree/main)
