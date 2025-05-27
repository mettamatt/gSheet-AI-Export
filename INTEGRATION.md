# Integration Guide: Adding AI Export to Existing Apps Script

This guide shows how to add the AI Export functionality to a Google Sheets project that already has Apps Script code.

## Quick Setup

### 1. Copy the Code
Copy the entire contents of `exportForAI-namespaced.gs` and paste it into your existing Apps Script project.

### 2. Add Menu Item
Add this line to your existing `onOpen()` function's menu array:
```javascript
{ name: 'Export for AI (JSON)', functionName: 'AIExport.exportSpreadsheetAsJson' }
```

### 3. Complete Example
```javascript
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const menuItems = [
    // Your existing menu items
    { name: 'My Function 1', functionName: 'myFunction1' },
    { name: 'My Function 2', functionName: 'myFunction2' },
    
    // Add a separator
    null,
    
    // Add the AI Export option
    { name: 'Export for AI (JSON)', functionName: 'AIExport.exportSpreadsheetAsJson' }
  ];
  
  ss.addMenu('My Menu', menuItems);
}
```

## Alternative Integration Methods

### Option 1: Separate Menu
Create a dedicated menu for export tools:
```javascript
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Your existing menu
  ss.addMenu('My Tools', [
    { name: 'Function 1', functionName: 'myFunction1' }
  ]);
  
  // Separate export menu
  ss.addMenu('Export Tools', [
    { name: 'Export for AI (JSON)', functionName: 'AIExport.exportSpreadsheetAsJson' }
  ]);
}
```

### Option 2: Programmatic Call
Call the export function from your own code:
```javascript
function myCustomExportFunction() {
  // Your custom logic here
  console.log('Starting export...');
  
  // Call the AI export
  AIExport.exportSpreadsheetAsJson();
  
  // Your post-export logic here
  console.log('Export completed!');
}
```

### Option 3: Trigger-Based
Set up a time-based or event-based trigger:
```javascript
function createTrigger() {
  ScriptApp.newTrigger('AIExport.exportSpreadsheetAsJson')
    .timeBased()
    .everyDays(1)
    .create();
}
```

## Important Notes

### Permissions
The first time you run the export, Google will request permission to:
- Access your Google Drive (to save the JSON file)
- Access your Google Sheets (to read the data)

### File Location
The exported JSON file is saved in the same Google Drive folder as your spreadsheet.

### Naming Conflicts
The namespace approach (`AIExport.functionName`) prevents any conflicts with your existing function names.

### Customization
You can rename the namespace if needed:
```javascript
// Change this line at the top:
const AIExport = {
// To:
const MyExporter = {

// Then update your menu item:
{ name: 'Export for AI', functionName: 'MyExporter.exportSpreadsheetAsJson' }
```

## Troubleshooting

### "Function not found" Error
- Make sure you copied the entire `AIExport` object
- Check that your menu item uses `AIExport.exportSpreadsheetAsJson` (with the dot)

### Permission Errors
- Run the function once manually from the Apps Script editor
- Accept all permission requests
- The menu item will work normally afterward

### No File Created
- Check your Google Drive's "Recent" section
- Ensure the spreadsheet is saved in Google Drive (not just downloaded/local)
- The file appears in the same folder as your spreadsheet

## Support

If you encounter issues:
1. Check the Apps Script execution log (View > Logs)
2. Ensure all permissions are granted
3. Try running `AIExport.exportSpreadsheetAsJson()` directly from the script editor first