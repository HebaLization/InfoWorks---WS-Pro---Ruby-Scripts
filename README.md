# InfoWorks WS Pro Ruby Tools

Practical Ruby scripts for automating repetitive tasks in Autodesk InfoWorks WS Pro.

## Tools

### 1. Scenario Change Analyzer (SCA)

**Purpose:** Compare two InfoWorks WS Pro scenarios and automatically create selection lists for:

- **New Objects** — new pipes and new nodes.
- **Modified Objects** — existing pipes with relevant network fields changed and existing nodes with changed demand.

**Status:** Add the tested `scenario_change_analyzer.rb` file to `scenario-change-analyzer/`.

> The Scenario Change Analyzer is based on a tested WS Pro workflow. Keep the tested script version unchanged when publishing it.

### 2. Excel → Selection Lists

**Purpose:** Create InfoWorks WS Pro Selection Lists from an Excel worksheet.

The script:

- Opens the first worksheet of an Excel workbook.
- Uses each Excel column header as the Selection List name.
- Reads the IDs below each header.
- Finds network objects by ID / Asset ID.
- Handles whitespace, non-breaking spaces, capitalization and common formatting differences in IDs.
- Creates or updates Selection Lists under a Selection List Group.
- Reports selected and missing IDs in the WS Pro script console.

## Excel format

Use the first row for Selection List names.

Example:

| Critical_Pipes | Renewal_2026 | Inspection |
|---|---|---|
| P-001 | P-101 | P-201 |
| P-002 | P-102 | P-202 |
| P-003 | P-103 | P-203 |

Each column becomes one Selection List.

## Requirements

- Autodesk InfoWorks WS Pro with Ruby scripting support.
- Microsoft Excel installed for the Excel → Selection Lists script.
- A Selection List Group in the InfoWorks database.

## Installation

1. Download the `.rb` script.
2. Open InfoWorks WS Pro.
3. Open the Ruby scripting environment.
4. Paste/open the script.
5. Follow the prompts.
6. For the Excel tool, provide the full path to the Excel workbook and select the target Selection List Group.

## Notes

These scripts are shared as practical engineering automation tools. Test them on a copy of your model/database before using them in production workflows.

The scripts are not an Autodesk product and are not affiliated with or endorsed by Autodesk.

## Contributing

If you improve a script, please consider sharing the change back so other hydraulic modelers can benefit from it.

## License

MIT License. See `LICENSE`.
