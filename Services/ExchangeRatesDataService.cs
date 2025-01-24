using AverageExchangeRatesAnalyzer.DataObjects;
using Microsoft.Extensions.Logging;

namespace AverageExchangeRatesAnalyzer.Business
{
    /// <summary>
    /// Responsible for processing and transforming data retrieved from the NPB API.
    /// </summary>
    public static class ExchangeRatesDataService
    {
        /// <summary>
        /// Sort average exchange rates dictionary.
        /// </summary>
        /// <param name="logger">
        /// Logger class.
        /// </param>
        /// <param name="exchangeRatesTables">
        /// Downloaded exchange rates data.
        /// </param>
        /// <returns>
        /// Sorted average exchange rates dictionary.
        /// Key: currency name.
        /// Value: average currency value for specific time.
        /// </returns>
        public static List<string> GetSortedAverageExchangeRatesDictionary(
            ILogger logger,
            List<ExchangeRatesTable> exchangeRatesTables)
        {
            Dictionary<string, double> averageExchangeRatesDictionary = new();
            using (logger.BeginScope("Get sorted average exchange rates dictionary"))
            {
                Dictionary<string, List<double>> collectedExchangeCurrency = CollectCurrencyToDictionary(exchangeRatesTables, logger);
                averageExchangeRatesDictionary = CalculateAverageRates(collectedExchangeCurrency, logger);
            }

            return averageExchangeRatesDictionary
                .OrderByDescending(currency => currency.Value)
                .ToDictionary()
                .Keys
                .ToList();
        }

        /// <summary>
        /// Calculate average value of currency from specific time.
        /// </summary>
        /// <param name="collectedExchangeCurrency">
        /// Collected Exchange Rates Dictionary.
        /// Key: Currency name.
        /// Value: List of currency values.
        /// </param>
        /// <returns>
        /// Average exchange Rates Dictionary.
        /// Key: Currency name.
        /// Value: List of currency values.
        /// </returns>
        private static Dictionary<string, double> CalculateAverageRates(
            Dictionary<string, List<double>> collectedExchangeCurrency,
            ILogger logger)
        {
            logger.LogInformation("Starting calculating currency to dictionary");
            Dictionary<string, double> averageExchangeRatesDictionary = new();
            foreach (var rate in collectedExchangeCurrency)
            {
                // SDR (MFW) XDR - Special drawing rights
                if (rate.Key == "XDR")
                {
                    continue;
                }

                double average = rate.Value.Sum(x => x) / rate.Value.Count;
                averageExchangeRatesDictionary[rate.Key] = average;
                logger.LogDebug("Currency [{CurrencyName}] - average rate {CurrencyValue} ", rate.Key, average);
            }

            return averageExchangeRatesDictionary;
        }

        /// <summary>
        /// Collect and sort requested data from specific time.
        /// </summary>
        /// <param name="exchangeRatesTables">
        /// Downloaded data.
        /// </param>
        /// <returns>
        /// Exchange Rates Dictionary.
        /// Key: Currency name.
        /// Value: List of currency values.
        /// </returns>
        private static Dictionary<string, List<double>> CollectCurrencyToDictionary(
            List<ExchangeRatesTable> exchangeRatesTables,
            ILogger logger)
        {
            Dictionary<string, List<double>> tmpCollectedExchangeCurrency = new Dictionary<string, List<double>>();
            logger.LogInformation("Starting colleting currency to dictionary");
            foreach (ExchangeRatesTable exchangeRatesTable in exchangeRatesTables)
            {
                foreach (Rate rate in exchangeRatesTable.Rates)
                {
                    if (!tmpCollectedExchangeCurrency.ContainsKey(rate.Code))
                    {
                        tmpCollectedExchangeCurrency[rate.Code] = new List<double> { rate.Mid };
                    }
                    else
                    {
                        tmpCollectedExchangeCurrency[rate.Code].Add(rate.Mid);
                    }
                }
            }

            return tmpCollectedExchangeCurrency;
        }
    }
}
