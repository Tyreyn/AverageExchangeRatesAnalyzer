using AverageExchangeRatesAnalyzer.DataObjects;

namespace AverageExchangeRatesAnalyzer.DataAccess
{
    public interface IHttpNbpApiClient
    {
        Task<(List<ExchangeRatesTable>, string)> GetExchangeRates(string startDate, string endDate);
    }
}