using AverageExchangeRatesAnalyzer.Business;
using AverageExchangeRatesAnalyzer.DataAccess;
using AverageExchangeRatesAnalyzer.DataObjects;
using AverageExchangeRatesAnalyzer.FileOperations.Exporters;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace AverageExchangeRatesAnalyzer
{
    static class Program
    {
        /// <summary>
        /// Logger class.
        /// </summary>
        private static ILogger? logger;

        /// <summary>
        /// HttpNbpApiClient class.
        /// </summary>
        private static HttpNbpApiClient httpNbpApiClient;

        /// <summary>
        /// Configuration root.
        /// </summary>
        private static IConfigurationRoot configuration;

        static void Main(string[] args)
        {
            Startup();
            logger.LogInformation("Starting script");
            DateTime currentDate = DateTime.UtcNow;

            //var response = httpNbpApiClient.GetExchangeRates(currentDate.AddDays(-7).ToString("yyyy-MM-dd"), currentDate.ToString("yyyy-MM-dd")).Result;
            var response = JsonConvert.DeserializeObject<List<ExchangeRatesTable>>(Samples.SampleApiOutput.SampleXmlOut2025_01_15_2025_01_22);
            List<string> sortedAverageExchangeRatesDictionary = ExchangeRatesDataService.GetSortedAverageExchangeRatesDictionary(logger, response);
            ExcelExporter excelExporter = new ExcelExporter(logger, configuration);
            excelExporter.PrepareDataAndExport(response, sortedAverageExchangeRatesDictionary);

        }

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
            httpNbpApiClient = new HttpNbpApiClient(logger);
        }
    }
}