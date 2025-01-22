using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AverageExchangeRatesAnalyzer.DataObjects
{
    /// <summary>
    /// Exchange rates table DTO.
    /// </summary>
    public class ExchangeRatesTable
    {
        /// <summary>
        /// Table type.
        /// </summary>
        public required string Table { get; set; }

        /// <summary>
        /// Table number.
        /// </summary>
        public required string No { get; set; }

        /// <summary>
        /// Publication date.
        /// </summary>
        public required DateTime EffectiveDate { get; set; }

        /// <summary>
        /// List of rates.
        /// </summary>
        public virtual required IList<Rate> Rates {get; set;}
    }
}
