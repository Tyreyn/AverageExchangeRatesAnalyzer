using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AverageExchangeRatesAnalyzer.DataObjects
{
    /// <summary>
    /// Exchange rate of individual currencies in the table.
    /// </summary>
    public class Rate
    {
        /// <summary>
        /// Gets or sets currency name.
        /// </summary>
        public required string Currency { get; set; }

        /// <summary>
        /// Gets or sets currency code.
        /// </summary>
        public required string Code { get; set; }

        /// <summary>
        /// Gets or sets converted average currency rate.
        /// </summary>
        public required float Mid { get; set; }
    }
}
