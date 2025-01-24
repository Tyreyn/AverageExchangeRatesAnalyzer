using AverageExchangeRatesAnalyzer.DataObjects;
using AverageExchangeRatesAnalyzer.Helpers.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;

namespace AverageExchangeRatesAnalyzer.FileOperations.Exporters
{
    /// <summary>
    /// Responsible for exporting processed data to an Excel file.
    /// Provides methods for creating and saving Excel files with the given data.
    /// </summary>
    public class ExcelExporter
    {
        /// <summary>
        /// Logger class.
        /// </summary>
        private readonly ILogger? logger;

        /// <summary>
        /// Day data was downloaded.
        /// </summary>
        private readonly DateTime currentDate;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExporter"/> class.
        /// </summary>
        /// <param name="logger">
        /// Logger class.
        /// </param>
        /// <param name="currentDate">
        /// Day data was downloaded.
        /// </param>
        public ExcelExporter(ILogger logger, DateTime currentDate)
        {
            this.logger = logger;
            this.currentDate = currentDate;
            this.logger.LogInformation("Initializing ExcelExporter");
        }

        /// <summary>
        /// Prepare and export data to excel file.
        /// </summary>
        /// <param name="exchangeRatesData">
        /// Downloaded exchange rates data.
        /// </param>
        /// <param name="sortedAverageExchangeRatesDictionary">
        /// Sorted list of average exchange rates from specific time.
        /// </param>
        public void PrepareDataAndExport(
            List<ExchangeRatesTable> exchangeRatesData,
            List<string> sortedAverageExchangeRatesDictionary)
        {
            using (this.logger?.BeginScope("Export data"))
            {
                try
                {
                    this.logger?.LogInformation("Creating excel spreadsheet");
                    SpreadsheetDocument excel = SpreadsheetDocument.Create("tmp.xlsx", SpreadsheetDocumentType.Workbook);
                    WorkbookPart workbookpart = excel.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    SheetData data = this.CreateSheetData(exchangeRatesData, sortedAverageExchangeRatesDictionary);
                    Columns columns = ExcelExporterHelper.AutoSizeCells(data);

                    worksheetPart.Worksheet.Append(columns);
                    worksheetPart.Worksheet.Append(data);

                    var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = ExcelExporterHelper.CreateStyleSheet();
                    stylesPart.Stylesheet.Save();

                    Sheets sheets = excel?.WorkbookPart?.Workbook.AppendChild<Sheets>(new Sheets());

                    sheets?.Append(
                        new Sheet()
                        {
                            Id = excel?.WorkbookPart?.GetIdOfPart(worksheetPart),
                            SheetId = 1,
                            Name = "report",
                        });

                    // Save & close
                    workbookpart.Workbook.Save();
                    excel?.Dispose();
                }
                catch (IOException ioExc)
                {
                    this.logger?.LogCritical(ioExc.Message);
                }
            }
        }

        /// <summary>
        /// Create sheet data.
        /// </summary>
        /// <param name="exchangeRatesData">
        /// Downloaded exchange rates data.
        /// </param>
        /// <param name="sortedAverageExchangeRates">
        /// Sorted list of average exchange rates from specific time.
        /// </param>
        /// <returns>
        /// Created sheet data.
        /// </returns>
        private SheetData CreateSheetData(
            List<ExchangeRatesTable> exchangeRatesData,
            List<string> sortedAverageExchangeRates)
        {
            this.logger?.LogInformation("Creating sheet data");
            Dictionary<string, CustomRow> excelRowDictionary = new Dictionary<string, CustomRow>();
            excelRowDictionary.Add("Header", new CustomRow());
            SheetData data = new SheetData();
            for (DateTime currentDay = this.currentDate.AddDays(-7); currentDay <= this.currentDate; currentDay = currentDay.AddDays(1.0))
            {
                this.logger?.LogInformation("Filling column {0}", currentDay);
                if (excelRowDictionary.Count <= 1)
                {
                    this.PrepareFirstColumn(excelRowDictionary, exchangeRatesData[0], sortedAverageExchangeRates);
                }

                ExchangeRatesTable currentDayExchangeRatesData = exchangeRatesData.FirstOrDefault(root => root.EffectiveDate == currentDay.Date);
                if (currentDayExchangeRatesData == null)
                {
                    this.FillDayOffColumn(currentDay, excelRowDictionary, sortedAverageExchangeRates);
                }
                else
                {
                    this.InsertHeaderCell(
                        excelRowDictionary["Header"].Row,
                        currentDayExchangeRatesData.EffectiveDate.ToOADate().ToString(),
                        excelRowDictionary["Header"].CellPointer++);

                    foreach (Rate rate in currentDayExchangeRatesData.Rates)
                    {
                        this.InsertNumberCell(
                            excelRowDictionary[rate.Code].Row,
                            rate.Mid,
                            excelRowDictionary[rate.Code].CellPointer++,
                            ExcelExporterHelper.FindNumberCellStyleIndex(sortedAverageExchangeRates, rate.Code));
                    }
                }
            }

            this.PrepareLegendRows(excelRowDictionary);
            int rowId = 0;
            this.logger?.LogInformation("Inserting data to sheet");
            foreach (KeyValuePair<string, CustomRow> keyValuePair in excelRowDictionary)
            {
                data.InsertAt(keyValuePair.Value.Row, rowId++);
            }

            return data;
        }

        /// <summary>
        /// Prepare legend rows.
        /// </summary>
        /// <param name="excelRowDictionary">
        /// Key: rate code name
        /// Value: custom row.
        /// </param>
        private void PrepareLegendRows(Dictionary<string, CustomRow> excelRowDictionary)
        {
            excelRowDictionary.Add("Highest", new CustomRow());
            this.InsertTextCell(
                excelRowDictionary["Highest"].Row,
                "najwyższa średnia",
                excelRowDictionary["Highest"].CellPointer++,
                9);
            excelRowDictionary.Add("SHighest", new CustomRow());
            this.InsertTextCell(
                excelRowDictionary["SHighest"].Row,
                "druga najwyższa średnia",
                excelRowDictionary["SHighest"].CellPointer++,
                10);
            excelRowDictionary.Add("Lowest", new CustomRow());
            this.InsertTextCell(
                excelRowDictionary["Lowest"].Row,
                "najniższa średnia",
                excelRowDictionary["Lowest"].CellPointer++,
                11);
            excelRowDictionary.Add("SLowest", new CustomRow());
            this.InsertTextCell(
                excelRowDictionary["SLowest"].Row,
                "druga najniższa średnia",
                excelRowDictionary["SLowest"].CellPointer++,
                12);
        }

        /// <summary>
        /// Prepare fist Column.
        /// </summary>
        /// <param name="excelRowDictionary">
        /// Key: rate code name
        /// Value: custom row.
        /// </param>
        /// <param name="exchangeRatesData">
        /// Downloaded exchange rates data.
        /// </param>
        /// <param name="sortedAverageExchangeRates">
        /// Sorted list of average exchange rates from specific time.
        /// </param>
        private void PrepareFirstColumn(
            Dictionary<string, CustomRow> excelRowDictionary,
            ExchangeRatesTable exchangeRatesData,
            List<string> sortedAverageExchangeRates)
        {
            this.logger?.LogInformation("Preparing first column based on {0} data", exchangeRatesData.EffectiveDate);

            this.InsertTextCell(
                excelRowDictionary["Header"].Row,
                string.Empty,
                excelRowDictionary["Header"].CellPointer++);

            foreach (Rate rate in exchangeRatesData.Rates)
            {
                if (!excelRowDictionary.ContainsKey(rate.Code))
                {
                    excelRowDictionary.Add(rate.Code, new CustomRow());
                    this.InsertTextCell(
                        excelRowDictionary[rate.Code].Row,
                        string.Format("{0} ({1})", rate.Currency, rate.Code),
                        excelRowDictionary[rate.Code].CellPointer++,
                        ExcelExporterHelper.FindTextCellStyleIndex(
                            sortedAverageExchangeRates,
                            rate.Code));
                }
            }
        }

        /// <summary>
        /// Fill row with day off values.
        /// </summary>
        /// <param name="currentDate">
        /// Current date.
        /// </param>
        /// <param name="excelRowDictionary">
        /// Key: rate code name
        /// Value: custom row.
        /// </param>
        /// <param name="sortedAverageExchangeRates">
        /// Sorted list of average exchange rates from specific time.
        /// </param>
        private void FillDayOffColumn(
            DateTime currentDate,
            Dictionary<string, CustomRow> excelRowDictionary,
            List<string> sortedAverageExchangeRates)
        {
            uint headerdayOffStyleIndex = 4;
            this.logger?.LogInformation("{0} was day off", currentDate);
            foreach (KeyValuePair<string, CustomRow> keyValuePair in excelRowDictionary)
            {
                if (keyValuePair.Key == "Header")
                {
                    this.InsertHeaderCell(
                        excelRowDictionary[keyValuePair.Key].Row,
                        currentDate.ToOADate().ToString(),
                        excelRowDictionary[keyValuePair.Key].CellPointer++,
                        headerdayOffStyleIndex);
                }
                else
                {
                    this.InsertTextCell(
                        excelRowDictionary[keyValuePair.Key].Row,
                        "Dzień wolny",
                        excelRowDictionary[keyValuePair.Key].CellPointer++,
                        ExcelExporterHelper.FindTextCellStyleIndex(sortedAverageExchangeRates, keyValuePair.Key, true));
                }
            }
        }

        /// <summary>
        /// Insert header cell to row.
        /// </summary>
        /// <param name="row">
        /// Row to which  add a cell.
        /// </param>
        /// <param name="content">
        /// Cell value.
        /// </param>
        /// <param name="cellIndex">
        /// Cell index.
        /// </param>
        /// <param name="styleIndex">
        /// Cell style index.
        /// </param>
        private void InsertHeaderCell(Row row, string content, int cellIndex, uint styleIndex = 1)
        {
            row.InsertAt<Cell>(
                new Cell()
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(content),
                    StyleIndex = styleIndex,
                },
                cellIndex);
        }

        /// <summary>
        /// Insert text cell to row.
        /// </summary>
        /// <param name="row">
        /// Row to which  add a cell.
        /// </param>
        /// <param name="content">
        /// Cell value.
        /// </param>
        /// <param name="cellIndex">
        /// Cell index.
        /// </param>
        /// <param name="styleIndex">
        /// Cell style index.
        /// </param>
        private void InsertTextCell(Row row, string content, int cellIndex, uint styleIndex = 3)
        {
            row.InsertAt<Cell>(
                new Cell()
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(content),
                    StyleIndex = styleIndex,
                },
                cellIndex);
        }

        /// <summary>
        /// Insert number cell to row.
        /// </summary>
        /// <param name="row">
        /// Row to which  add a cell.
        /// </param>
        /// <param name="content">
        /// Cell value.
        /// </param>
        /// <param name="cellIndex">
        /// Cell index.
        /// </param>
        /// <param name="styleIndex">
        /// Cell style index.
        /// </param>
        private void InsertNumberCell(Row row, double content, int cellIndex, uint styleIndex = 2)
        {
            row.InsertAt<Cell>(
                new Cell()
                {
                    DataType = CellValues.Number,
                    CellValue = new CellValue(content),
                    StyleIndex = styleIndex,
                },
                cellIndex);
        }
    }
}
