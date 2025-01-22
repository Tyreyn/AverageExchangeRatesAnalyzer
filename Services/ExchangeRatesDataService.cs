using AverageExchangeRatesAnalyzer.DataObjects;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace AverageExchangeRatesAnalyzer.Business
{

    /// <summary>
    /// Responsible for processing and transforming data retrieved from the NPB API.
    /// </summary>
    public class ExchangeRatesDataService
    {
        /// <summary>
        /// Logger class.
        /// </summary>
        private readonly ILogger? logger;

        /// <summary>
        /// Average exchange Rates Dictionary.
        /// Key: Currency name.
        /// Value: List of currency values.
        /// </summary>
        private Dictionary<string, double> AverageExchangeRatesDictionary = new Dictionary<string, double>();

        public ExchangeRatesDataService(ILogger logger, List<ExchangeRatesTable> exchangeRatesTables)
        {
            this.logger = logger;
            using (this.logger.BeginScope("ExchangeRatesDataService"))
            {
                Dictionary<string, List<double>> collectedExchangeCurrency = CollectCurrencyToDictionary(exchangeRatesTables);
                this.CalculateAverageRates(collectedExchangeCurrency);
            }
        }

        /// <summary>
        /// Sort average exchange rates dictionary.
        /// </summary>
        /// <returns>
        /// Sorted average exchange rates dictionary.
        /// Key: currency name.
        /// Value: average currency value for specific time.
        /// </returns>
        public Dictionary<string, double> GetSortedAverageExchangeRatesDictionary()
        {
            return this.AverageExchangeRatesDictionary.OrderByDescending(currency => currency.Value).ToDictionary();
        }

        /// <summary>
        /// Calculate average value of currency from specific time.
        /// </summary>
        /// <param name="collectedExchangeCurrency">
        /// Collected Exchange Rates Dictionary.
        /// Key: Currency name.
        /// Value: List of currency values.
        /// </param>
        private void CalculateAverageRates(Dictionary<string, List<double>> collectedExchangeCurrency)
        {
            this.logger.LogInformation("Starting calculating currency to dictionary");
            foreach (var rate in collectedExchangeCurrency)
            {
                double average = rate.Value.Sum(x => x) / rate.Value.Count();
                AverageExchangeRatesDictionary[rate.Key] = average;
                this.logger.LogDebug("Currency [{currencyName}] - average rate {currencyValue} ", rate.Key, average);
            }
        }

        /// <summary>
        /// Collect and sort requested data from specific time.
        /// </summary>
        /// <param name="exchangeRatesTables"></param>
        /// <returns>
        /// Exchange Rates Dictionary.
        /// Key: Currency name.
        /// Value: List of currency values.
        /// </returns>
        private Dictionary<string, List<double>> CollectCurrencyToDictionary(List<ExchangeRatesTable> exchangeRatesTables)
        {
            Dictionary<string, List<double>> tmpCollectedExchangeCurrency = new Dictionary<string, List<double>>();
            this.logger.LogInformation("Starting colleting currency to dictionary");
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
