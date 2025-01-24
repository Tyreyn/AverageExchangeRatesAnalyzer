using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using AverageExchangeRatesAnalyzer.DataObjects;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using static System.Runtime.InteropServices.JavaScript.JSType;
using AverageExchangeRatesAnalyzer.Helpers.Excel;

namespace AverageExchangeRatesAnalyzer.FileOperations.Exporters
{
    public class ExcelExporter
    {
        /// <summary>
        /// Logger class.
        /// </summary>
        private readonly ILogger? logger;

        /// <summary>
        /// Indicates whether reports should be in single file.
        /// </summary>
        private bool SingleFile = false;

        /// <summary>
        /// Indicates whether all reports are to be deleted.
        /// </summary>
        private bool RaportCleanup = false;

        /// <summary>
        /// Default name of exported file.
        /// </summary>
        private string RaportFileName = "Currency_rate_raport";

        /// <summary>
        /// Default name of raports folder.
        /// </summary>
        private string RaportFolderName = "Raports";

        /// <summary>
        /// Destination folder where raport will be stored.
        /// </summary>
        private static string DestinationFolder = Directory.GetCurrentDirectory();

        /// <summary>
        /// Configuration root.
        /// </summary>
        private IConfigurationRoot configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExporter"/> class.
        /// </summary>
        /// <param name="logger">
        /// Logger class.
        /// </param>
        public ExcelExporter(ILogger logger, IConfigurationRoot configuration)
        {
            this.logger = logger;
            this.configuration = configuration;
            this.logger.LogInformation("Initializing ExcelExporter");
            this.LoadSettings();
        }

        public void PrepareDataAndExport(
            List<ExchangeRatesTable> ExchangeRatesData,
            List<string> sortedAverageExchangeRatesDictionary)
        {
            using (this.logger.BeginScope("Export data"))
            {
                try
                {
                    this.logger.LogInformation("Creating excel spreadsheet");
                    SpreadsheetDocument excel = SpreadsheetDocument.Create("tmp.xlsx", SpreadsheetDocumentType.Workbook);
                    WorkbookPart workbookpart = excel.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    SheetData data = CreateSheetData(ExchangeRatesData, sortedAverageExchangeRatesDictionary);
                    Columns columns = ExcelExporterHelper.AutoSizeCells(data);

                    worksheetPart.Worksheet.Append(columns);
                    worksheetPart.Worksheet.Append(data);

                    var stylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = ExcelExporterHelper.CreateStyleSheet();
                    stylesPart.Stylesheet.Save();

                    Sheets sheets = excel.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                    sheets.Append(new Sheet()
                    {
                        Id = excel.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Raport"
                    });

                    //Save & close
                    workbookpart.Workbook.Save();
                    excel.Dispose();
                }
                catch (IOException ioExc)
                {
                    this.logger.LogCritical(ioExc.Message);
                }
            }
        }

        private SheetData CreateSheetData(
            List<ExchangeRatesTable> ExchangeRatesData,
            List<string> sortedAverageExchangeRates)
        {
            string highest = sortedAverageExchangeRates[0];
            this.logger.LogInformation("Creating sheet data");
            DateTime startDate = DateTime.UtcNow.AddDays(-7);
            Dictionary<string, CustomRow> excelRowDictionary = new Dictionary<string, CustomRow>();
            excelRowDictionary.Add("Header", new CustomRow());
            SheetData data = new SheetData();
            for (DateTime currentDate = DateTime.UtcNow.AddDays(-7); currentDate < DateTime.UtcNow; currentDate = currentDate.AddDays(1.0))
            {
                this.logger.LogInformation("Filling column {0}", currentDate);
                if (excelRowDictionary.Count <= 1) PrepareFirstColumn(excelRowDictionary, ExchangeRatesData[0], sortedAverageExchangeRates);
                ExchangeRatesTable currentDayExchangeRatesData = ExchangeRatesData.Where(root => root.EffectiveDate == currentDate.Date).FirstOrDefault();
                if (currentDayExchangeRatesData == null)
                {
                    FillDayOffColumn(currentDate, excelRowDictionary, sortedAverageExchangeRates);
                }
                else
                {
                    InsertHeaderCell(excelRowDictionary["Header"].Row,
                        currentDayExchangeRatesData.EffectiveDate.ToOADate().ToString(),
                        excelRowDictionary["Header"].CellPointer++);

                    foreach (Rate rate in currentDayExchangeRatesData.Rates)
                    {
                        InsertNumberCell(
                            excelRowDictionary[rate.Code].Row,
                            rate.Mid,
                            excelRowDictionary[rate.Code].CellPointer++,
                            ExcelExporterHelper.FindNumberCellStyleIndex(sortedAverageExchangeRates, rate.Code));
                    }
                }
            }

            int rowId = 0;
            this.logger.LogInformation("Inserting data to sheet");
            foreach (KeyValuePair<string, CustomRow> keyValuePair in excelRowDictionary)
            {
                data.InsertAt(keyValuePair.Value.Row, rowId++);
            }

            return data;
        }


        private void PrepareFirstColumn(
            Dictionary<string, CustomRow> excelRowDictionary,
            ExchangeRatesTable ExchangeRatesData,
            List<string> sortedAverageExchangeRates)
        {
            this.logger.LogInformation("Preparing first column based on {0} data", ExchangeRatesData.EffectiveDate);
            InsertTextCell(excelRowDictionary["Header"].Row, string.Empty, excelRowDictionary["Header"].CellPointer++);

            foreach (Rate rate in ExchangeRatesData.Rates)
            {
                if (!excelRowDictionary.ContainsKey(rate.Code))
                {
                    excelRowDictionary.Add(rate.Code, new CustomRow());
                    InsertTextCell(excelRowDictionary[rate.Code].Row,
                        string.Format("{0} ({1})", rate.Currency, rate.Code),
                        excelRowDictionary[rate.Code].CellPointer++,
                        ExcelExporterHelper.FindTextCellStyleIndex(sortedAverageExchangeRates, rate.Code));
                }
            }
        }

        private void FillDayOffColumn(
            DateTime currentDate,
            Dictionary<string, CustomRow> excelRowDictionary,
            List<string> sortedAverageExchangeRates)
        {
            uint HeaderdayOffStyleIndex = 4;
            this.logger.LogInformation("{0} was day off", currentDate);
            foreach (KeyValuePair<string, CustomRow> keyValuePair in excelRowDictionary)
            {
                if (keyValuePair.Key == "Header")
                {
                    InsertHeaderCell(
                        excelRowDictionary[keyValuePair.Key].Row,
                        currentDate.ToOADate().ToString(),
                        excelRowDictionary[keyValuePair.Key].CellPointer++,
                        HeaderdayOffStyleIndex);
                }
                else
                {
                    InsertTextCell(excelRowDictionary[keyValuePair.Key].Row,
                        "Dzień wolny",
                        excelRowDictionary[keyValuePair.Key].CellPointer++,
                        ExcelExporterHelper.FindTextCellStyleIndex(sortedAverageExchangeRates, keyValuePair.Key, true));
                }
            }
        }

        private void InsertHeaderCell(Row row, string content, int cellIndex, uint styleIndex = 1)
        {
            row.InsertAt<Cell>(new Cell()
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(content),
                StyleIndex = styleIndex
            }, cellIndex);
        }

        private void InsertTextCell(Row row, string content, int cellIndex, uint styleIndex = 3)
        {
            row.InsertAt<Cell>(new Cell()
            { 
                DataType = CellValues.String,
                CellValue = new CellValue(content),
                StyleIndex = styleIndex 
            }, cellIndex);
        }

        private void InsertNumberCell(Row row, double content, int cellIndex, uint styleIndex = 2)
        {
            row.InsertAt<Cell>(new Cell() 
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(content),
                StyleIndex = styleIndex 
            }, cellIndex);
        }

        private void LoadSettings()
        {
            string tmpRaportFolderName = configuration.GetValue<string>("RaportFolderName");
            string tmpRaportFileName = configuration.GetValue<string>("RaportFileName");
            string tmpDestinationFolder = configuration.GetValue<string>("DestinationFolder");
            bool tmpSingleFile = configuration.GetValue<bool>("SingleFile");
            bool tmpRaportCleanup = configuration.GetValue<bool>("RaportCleanup");
            if (tmpDestinationFolder != null) DestinationFolder = tmpDestinationFolder;
            if (tmpSingleFile != null) SingleFile = tmpSingleFile;
            if (tmpRaportCleanup != null) RaportCleanup = tmpRaportCleanup;
            if (tmpRaportFileName != null) RaportFileName = tmpRaportFileName;
            if (tmpRaportFolderName != null) RaportFolderName = tmpRaportFolderName;

            this.logger.LogInformation("Files will be saved in {destination}/<month_name>", Path.Combine(DestinationFolder, RaportFolderName));
            this.logger.LogInformation("File name will be {fileName}_<end_date>", RaportFileName);
            if (SingleFile) this.logger.LogInformation("Single File Mode", DestinationFolder);
            if (RaportCleanup) this.logger.LogInformation("All raports will be deleted", DestinationFolder);
        }
    }
}
