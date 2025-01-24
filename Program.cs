namespace AverageExchangeRatesAnalyzer
{
    using AverageExchangeRatesAnalyzer.Business;
    using AverageExchangeRatesAnalyzer.DataAccess;
    using AverageExchangeRatesAnalyzer.DataObjects;
    using AverageExchangeRatesAnalyzer.FileOperations.Exporters;
    using AverageExchangeRatesAnalyzer.FileOperations.FileManagement;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Logging;
    using Newtonsoft.Json;

    /// <summary>
    /// Main class.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Logger class.
        /// </summary>
        private static ILogger? logger;

        /// <summary>
        /// HttpNbpApiClient class.
        /// </summary>
        private static HttpNbpApiClient? httpNbpApiClient;

        /// <summary>
        /// Configuration root.
        /// </summary>
        private static IConfigurationRoot? configuration;

        /// <summary>
        /// Main method.
        /// </summary>
        /// <param name="args">
        /// Launch arguments.
        /// </param>
        private static void Main(string[] args)
        {
            Startup();
            DateTime currentDate = DateTime.UtcNow;

            if (httpNbpApiClient != null && logger != null && configuration != null)
            {
                (List<ExchangeRatesTable>?, string) response = httpNbpApiClient.GetExchangeRates(
                    currentDate.AddDays(-7).ToString("yyyy-MM-dd"),
                    currentDate.ToString("yyyy-MM-dd")).Result;
                if (response.Item1 == null)
                {
                    logger.LogCritical("Something went wrong!");
                }
                else
                {
                    List<string> sortedAverageExchangeRatesDictionary = ExchangeRatesDataService.GetSortedAverageExchangeRatesDictionary(logger, response.Item1);

                    ExcelExporter excelExporter = new ExcelExporter(logger, currentDate);
                    excelExporter.PrepareDataAndExport(response.Item1, sortedAverageExchangeRatesDictionary);

                    ReportFileManager reportFileManager = new ReportFileManager(logger, configuration, currentDate);
                    reportFileManager.PerformManagementOfNewReport();
                }
            }
            else
            {
                logger?.LogCritical("Something went wrong!");
            }
        }

        /// <summary>
        /// Initialize required objects.
        /// </summary>
        private static void Startup()
        {
            configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile($"appsettings.json", false, true)
                .Build();

            var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder
                    .AddConfiguration(configuration.GetSection("Logging"))
                    .AddSimpleConsole();
            });
            logger = loggerFactory.CreateLogger("Program");
            logger.LogInformation("Starting script");
            httpNbpApiClient = new HttpNbpApiClient(logger);
        }
    }
}