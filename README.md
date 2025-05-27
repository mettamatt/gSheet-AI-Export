# gSheet-AI-Export

**gSheet-AI-Export** is a Google Apps Script designed to export data and formulas from Google Sheets as a JSON file, optimized for use with AI models like ChatGPT and Claude. The script provides an easy way to extract all the sheet data and formulas with additional metadata for enhanced AI analysis.

## Features

- **Custom Menu Integration:** Adds an "Export Tools" menu to your Google Sheet, with the option "Export for AI (JSON)".
- **Cell-Level Data Mapping:** Exports each cell's data and formula along with its exact cell address (e.g., "A1"), allowing AI models to understand the precise location of each piece of content.
- **AI-Optimized Export:** Converts all spreadsheet data and formulas into a structured JSON format for optimal AI interaction.
- **Detailed Metadata:** Includes spreadsheet metadata such as sheet names, row/column counts, total formulas, and more.
- **Direct Download:** Creates downloadable JSON files without requiring Google Drive permissions.

## Installation Options

### New Installation
For new Google Sheets or sheets without existing Apps Script code:

1. **Create the Script:**
    - In your Google Sheet, click on `Extensions > Apps Script`.
    - Replace any existing code with the contents of `exportForAI.gs` from this repository.
    - Add these functions to create the menu and handle the UI:
    ```javascript
    function onOpen() {
      SpreadsheetApp.getActiveSpreadsheet()
        .addMenu('Export Tools', [{ name: 'Export for AI (JSON)', functionName: 'exportForAI' }]);
    }
    
    function exportForAI() {
      const result = AIExport.exportSpreadsheetAsJson();
      SpreadsheetApp.getUi().alert(
        'Export Complete', 
        `Download ready! File: ${result.filename} (${Math.round(result.size/1024)}KB)\n\nCopy this link to download:\n${result.dataUrl}`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    ```
    - Save the project with a meaningful name (e.g., "Export for AI").

2. **Refresh the Google Sheet:**
    - After saving the script, close the Apps Script editor and hard refresh your Google Sheet.
    - You will see a new menu called "Export Tools" with the "Export for AI (JSON)" option.

3. **Exporting Data:**
    - Once the custom menu appears, click on `Export Tools > Export for AI (JSON)` to export the spreadsheet data and formulas as a JSON file.
    - The script will generate a downloadable JSON file for immediate download.

### Integration with Existing Apps Script
If your Google Sheet already has Apps Script code:

1. **Copy the Code:**
    - Copy the contents of `exportForAI.gs` from this repository into your existing Apps Script project.

2. **Create a Wrapper Function:**
    - Since the export function returns download data, create a wrapper to handle the UI:
    ```javascript
    function exportForAI() {
      const result = AIExport.exportSpreadsheetAsJson();
      SpreadsheetApp.getUi().alert(
        'Export Complete', 
        `Download ready! File: ${result.filename} (${Math.round(result.size/1024)}KB)\n\nCopy this link to download:\n${result.dataUrl}`, 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    ```

3. **Add Menu Item:**
    - Add this line to your existing `onOpen()` function's menu array:
    ```javascript
    { name: 'Export for AI (JSON)', functionName: 'exportForAI' }
    ```

4. **Complete Example:**
    ```javascript
    function onOpen() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const menuItems = [
        // Your existing menu items
        { name: 'My Function', functionName: 'myFunction' },
        null, // separator
        { name: 'Export for AI (JSON)', functionName: 'exportForAI' }
      ];
      ss.addMenu('My Menu', menuItems);
    }
    ```

The `AIExport` namespace prevents any conflicts with your existing function names.

## Usage with AI Models (ChatGPT and Claude)

Once you've exported the JSON file, you can upload it to AI models like **ChatGPT** or **Claude** to perform various tasks, such as:

- Analyzing spreadsheet data trends.
- Explaining or improving formulas.
- Visualizing data or performing advanced analytics.

### Example Queries to Ask AI Models:

- "Can you explain the formulas in this spreadsheet?"
- "What insights can you draw from the data?"
- "Are there any improvements you recommend for the formulas?"
- "Help me generate graphs based on this data."

## JSON Output Example

```json
{
  "explanation": "Generated from Google Sheets. Only non-empty cells are included. Use \"sheet_name / cell_address\" to reference data programmatically.",
  "spreadsheet_metadata": {
    "spreadsheet_name": "Example Sheet",
    "export_timestamp": "2024-10-08T12:00:00Z",
    "total_sheets": 2,
    "total_rows": 100,
    "total_columns": 10,
    "total_cells": 1000,
    "total_formulas": 50
  },
  "sheets": [
    {
      "sheet_name": "Sheet1",
      "sheet_id": 123456789,
      "sheet_index": 1,
      "sheet_metadata": {
        "num_rows": 50,
        "num_columns": 5,
        "num_cells": 250,
        "num_formulas": 20,
        "visibility": "Visible"
      },
      "cells": [
        {
          "cell_address": "A1",
          "value": "Name"
        },
        {
          "cell_address": "B1",
          "value": "Age"
        },
        {
          "cell_address": "C1",
          "value": "Salary"
        },
        {
          "cell_address": "A2",
          "value": "John Doe"
        },
        {
          "cell_address": "B2",
          "value": 30
        },
        {
          "cell_address": "C2",
          "value": 50000,
          "formula": "=B2*1000"
        },
        {
          "cell_address": "A3",
          "value": "Jane Smith"
        },
        {
          "cell_address": "B3",
          "value": 25
        },
        {
          "cell_address": "C3",
          "value": 45000,
          "formula": "=B3*1800"
        }
      ]
    }
  ]
}
```

## Contributing

Contributions are welcome! If you'd like to enhance the script or fix any issues, feel free to fork the repository and submit a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
