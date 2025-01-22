using AverageExchangeRatesAnalyzer.Business;
using AverageExchangeRatesAnalyzer.DataAccess;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

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
        /// ExchangeRatesDataService service.
        /// </summary>
        private static ExchangeRatesDataService exchangeRatesDataService;

        /// <summary>
        /// Destination folder where raport will be stored.
        /// </summary>
        private static string DestinationFolder = Directory.GetCurrentDirectory();

        /// <summary>
        /// Indicates whether reports should be in single file.
        /// </summary>
        private static bool SingleFile = false;

        /// <summary>
        /// Indicates whether all reports are to be deleted.
        /// </summary>
        private static bool RaportCleanup = false;

        static void Main(string[] args)
        {
            Startup();
            logger.LogInformation("Starting script");
            DateTime currentDate = DateTime.UtcNow;

            var response = httpNbpApiClient.GetExchangeRates(currentDate.AddDays(-7).ToString("yyyy-MM-dd"), currentDate.ToString("yyyy-MM-dd")).Result;

            exchangeRatesDataService = new ExchangeRatesDataService(logger, response.Item1);
            var sorted = exchangeRatesDataService.GetSortedAverageExchangeRatesDictionary();

        }

        private static void Startup()
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile($"appsettings.json", false, true)
                .Build();

            var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder
                    .AddConfiguration(configuration.GetSection("Logging"))
                    .AddSimpleConsole();
            });
            string tmpDestinationFolder = configuration.GetValue<string>("DestinationFolder");
            bool tmpSingleFile = configuration.GetValue<bool>("SingleFile");
            bool tmpRaportCleanup = configuration.GetValue<bool>("RaportCleanup");
            if (tmpDestinationFolder != null) DestinationFolder = tmpDestinationFolder;
            if (tmpSingleFile != null) SingleFile = tmpSingleFile;
            if (tmpRaportCleanup != null) RaportCleanup = tmpRaportCleanup;
            logger = loggerFactory.CreateLogger("Program");
            logger.LogInformation("Files will be saved in {destination}", DestinationFolder);
            httpNbpApiClient = new HttpNbpApiClient(logger);
        }
    }
}