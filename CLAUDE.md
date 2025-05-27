# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script that exports Google Sheets data and formulas as JSON files optimized for AI analysis. The script adds a custom menu to Google Sheets and allows users to export comprehensive spreadsheet data including cell values, formulas, metadata, comments, notes, data validation, merged ranges, and hyperlinks.

## Architecture

### Core Components

- **exportForAI.gs**: Single Google Apps Script file containing all functionality
  - `onOpen()`: Creates custom menu in Google Sheets UI
  - `saveAsJson()`: Main export function that processes all sheets and generates JSON
  - `columnToLetter()` & `getCellAddress()`: Utility functions for A1 notation conversion

### Key Features

- **Performance Optimization**: Uses bulk data fetching (`getDataRange().getValues()`, `getFormulas()`, etc.) to minimize API calls
- **Data Filtering**: Excludes empty cells and null fields to reduce file size
- **Comprehensive Metadata**: Captures sheet-level and spreadsheet-level statistics
- **Cell-Level Mapping**: Each cell includes its exact A1 notation address
- **Merged Range Handling**: Creates lookup map for efficient merged cell processing

### JSON Output Structure

The exported JSON follows this structure:
```
{
  "explanation": "...",
  "spreadsheet_metadata": { /* overall stats */ },
  "sheets": [
    {
      "sheet_name": "...",
      "sheet_metadata": { /* sheet stats */ },
      "cells": [
        {
          "cell_address": "A1",
          "value": "...",
          "formula": "...",
          // optional: comment, note, data_validation, merge_info, hyperlink
        }
      ]
    }
  ]
}
```

## Development Notes

- This is a Google Apps Script project - code must be deployed through Google Sheets Extensions > Apps Script
- No build process, linting, or testing infrastructure is present
- Single file architecture with all functions in one script
- Uses Google Apps Script APIs exclusively (SpreadsheetApp, HtmlService, Utilities)