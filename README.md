# GMT Research Accounting Screen Excel Spreadsheet

## Installation

- In the Excel file, press `Alt+F11`to open the VBA Editor
- Go to `Tools -> Reference`, enable the `Microsoft Scripting Runtime`
- Go to `File -> Import File...`, import all the `*.bas` files
  (For how to import multiple files at once: http://www.knowledgeworkerblog.com/2009/02/how-to-import-multiple-modules-into-vba.html)

## Usage

There are 4 "Macro"s available:

- `SaveJsonToFile`: Export the existing data on the worksheet to a JSON file `data.json` in the same folder
- `ClearWorksheetData`: Clear the existing data on the worksheets
- `DownloadJsonFile`: Download the JSON data file from a designated URL and save it as `data.json` in the same folder (overwrite)
- `LoadJsonFromFile`: Put the data from `data.json` to the worksheets

You can create buttons and assign them with the corresponding Macros.

## Configuration

- In the module `Config`, you can configure the URL of the JSON data file
  and the worksheet names for `SaveJsonToFile` and `ClearWorksheetData`
