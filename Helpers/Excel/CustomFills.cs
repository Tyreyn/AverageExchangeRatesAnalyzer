using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AverageExchangeRatesAnalyzer.Helpers.Excel
{
    /// <summary>
    /// Custom fills class.
    /// </summary>
    public static class CustomFills
    {
        /// <summary>
        /// Get custom fills.
        /// </summary>
        /// <returns>
        /// Custom fills.
        /// </returns>
        public static Fills GetCustomFills()
        {
            Fills fills = new Fills();
            fills.Append(
                new Fill() // Fill index 0
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.None,
                    },
                });
            fills.Append(
                new Fill() // Fill index 1
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 2
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.LightBlue),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 3
                {
                    PatternFill = new PatternFill
                    {
                        BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.LightBlue).Rgb },
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                    },
                });
            fills.Append(
                new Fill() // Fill index 4
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Green),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.Green).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 5
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.LightGreen),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.LightGreen).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 6
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.DarkRed),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.DarkRed).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 7
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Solid,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Red),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.Red).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 8
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.Green).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 9
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.LightGreen).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 10
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                        BackgroundColor = new BackgroundColor
                        {
                            Rgb = TranslateForeground(System.Drawing.Color.DarkRed).Rgb,
                        },
                    },
                });
            fills.Append(
                new Fill() // Fill index 11
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                        BackgroundColor = new BackgroundColor { Rgb = TranslateForeground(System.Drawing.Color.Red).Rgb },
                    },
                });
            fills.Append(
                new Fill() // Fill index 12
                {
                    PatternFill = new PatternFill
                    {
                        PatternType = PatternValues.Gray0625,
                        ForegroundColor = TranslateForeground(System.Drawing.Color.Black),
                    },
                });
            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            return fills;
        }

        private static ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            return new ForegroundColor()
            {
                Rgb = new HexBinaryValue()
                {
                    Value =
                              System.Drawing.ColorTranslator.ToHtml(
                              System.Drawing.Color.FromArgb(
                                  fillColor.A,
                                  fillColor.R,
                                  fillColor.G,
                                  fillColor.B)).Replace("#", string.Empty),
                },
            };
        }
    }
}
