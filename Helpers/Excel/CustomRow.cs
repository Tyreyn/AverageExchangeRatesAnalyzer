using DocumentFormat.OpenXml.Spreadsheet;

namespace AverageExchangeRatesAnalyzer.Helpers.Excel
{
    /// <summary>
    /// Custom row class.
    /// </summary>
    public class CustomRow()
    {
        /// <summary>
        /// Gets or sets row class.
        /// </summary>
        public Row Row { get; set; } = new();

        /// <summary>
        /// Gets or sets current cell in row pointer.
        /// </summary>
        public int CellPointer { get; set; } = 0;
    }
}
