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
        /// Gets or sets table type.
        /// </summary>
        public required string Table { get; set; }

        /// <summary>
        /// Gets or sets table number.
        /// </summary>
        public required string No { get; set; }

        /// <summary>
        /// Gets or sets publication date.
        /// </summary>
        public required DateTime EffectiveDate { get; set; }

        /// <summary>
        /// Gets or sets list of rates.
        /// </summary>
        public virtual required IList<Rate> Rates { get; set; }
    }
}
