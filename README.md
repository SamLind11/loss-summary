# Loss Summary Macro for Google Sheets

The code in Code.js is a Javascript version of a script written for a macro in Google Sheets.  The purpose of the script is to read from another Google Spreadsheet called "Profit Margin", which tracks profit margins on various projects.  Any projects which have a profit margin at or below 10% are counted as losses and are recorded to a new Spreadsheet.  The macro uses `ScriptProperties` to track the last sheet and row it has reviewed, so that on each run, it only checks for new values in the Profit Margins sheet.  

## Running the Macro

![loss-summary](https://github.com/SamLind11/loss-summary/assets/131621692/c0cc682f-0761-4254-b140-efcdb551589e)

In the Loss Summary Spreadsheet, pressing a button runs the macro.  Once the script is added as a .gs file to a Spreadsheet, its functions can also be added as macros for the sheet, and the `lossSummary()` function can be run from the Extensions menu under Macros.

## Authorship

This macro was developed for internal use for Beechwood Metalworks Inc. in Burlington, NC, and was written by Sam Lind.
