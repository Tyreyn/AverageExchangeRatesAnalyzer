# Average Exchange Rates Analyzer
Console app that download, analyze and extract date to datasheet. Sample output is in "Currency_rate_report_24012025.xlsx"

## Requirements
Additional config arguments that could be added in appsettings.json
```c#
{
    "DestinationFolder": "C:\\Users\\athos\\source",
    "reportFolderName": "reports",
    "reportFileName": "Currency_rate_report",
    "SingleFile": false,
    "reportCleanup": false,
}
```
- DestinationFolder (default: current directory) - root destination folder where report will be stored.
- reportFolderName (default: "reports") - name of reports folder
- reportFileName (default: "Currency_rate_report") - name of exported file + .xsls
- SingleFile (default: false) - indicates whether reports should be in single file. If true previous reports will be deleted, otherwise files will be ordered in folders called <year>_<month> e.g. 2025_01
- reportCleanup (default: false) - indicates whether all reports should be deleted.


## Diagram
#### Main diagram
![Alt text](https://github.com/Tyreyn/AverageExchangeRatesAnalyzer/blob/main/MainDiagram.png "Main diagram")
#### Data download diagram
![Alt text](https://github.com/Tyreyn/AverageExchangeRatesAnalyzer/blob/main/DataDownloadDiagram.png "Data download diagram")
#### Prepare excel file diagram
![Alt text](https://github.com/Tyreyn/AverageExchangeRatesAnalyzer/blob/main/PrepareExcelFile.png "Prepare excel file diagram")
#### Perform management diagram
![Alt text](https://github.com/Tyreyn/AverageExchangeRatesAnalyzer/blob/main/PerformManagemenDiagram.png "Perform management diagram")

## Example excel file from 24.01.2025
![Alt text](https://github.com/Tyreyn/AverageExchangeRatesAnalyzer/blob/main/excel.png "This is admin panel")

## Using 
Download folder bin\Release\net9.0 and launch AverageExchangeRatesAnalyzer.exe or launch solution via IDE.

## Docs
HTML documentation can be find in Docs\html\index.html