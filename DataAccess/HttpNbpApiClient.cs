using System.Net.Http.Headers;
using System.Net.Http.Json;
using AverageExchangeRatesAnalyzer.DataObjects;
using Microsoft.Extensions.Logging;

namespace AverageExchangeRatesAnalyzer.DataAccess
{
    /// <summary>
    /// A client responsible for handling HTTP requests to the NBP API.
    /// Provides methods for retrieving from the API endpoints.
    /// </summary>
    public class HttpNbpApiClient
    {
        /// <summary>
        /// NBP Api uri string.
        /// </summary>
        private const string UriString = "https://api.nbp.pl/api/exchangerates/tables/a/";

        /// <summary>
        /// Logger class.
        /// </summary>
        private readonly ILogger? logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="HttpNbpApiClient"/> class.
        /// </summary>
        /// <param name="logger">
        /// Logger class.
        /// </param>
        public HttpNbpApiClient(ILogger logger)
        {
            this.logger = logger;
            this.logger.LogInformation("Starting Http Nbp Api Client");
        }

        /// <summary>
        /// Get data from NBP Api.
        /// </summary>
        /// <param name="startDate">
        /// Start date in string.
        /// </param>
        /// <param name="endDate">
        /// End date in string.
        /// </param>
        /// <returns>
        /// List of exchange rates from specific date.
        /// </returns>
        public async Task<(List<ExchangeRatesTable>?, string)> GetExchangeRates(string startDate, string endDate)
        {
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(UriString);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                this.logger.LogInformation("Sending request {0}", UriString + $"{startDate}/{endDate}");
                HttpResponseMessage response = await client.GetAsync($"{startDate}/{endDate}");

                if (response.IsSuccessStatusCode)
                {
                    try
                    {
                        List<ExchangeRatesTable>? requestExchangeRatesTable;
                        requestExchangeRatesTable = await response.Content.ReadFromJsonAsync<List<ExchangeRatesTable>>();
                        return (requestExchangeRatesTable, "Exchange rates downloaded successfully");
                    }
                    catch (Exception ex)
                    {
                        return (null, ex.ToString());
                    }
                }
                else
                {
                    this.logger?.LogCritical(response.StatusCode.ToString());
                    return (null, response.StatusCode.ToString());
                }
            }
        }
    }
}
