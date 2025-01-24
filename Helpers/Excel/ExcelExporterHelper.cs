using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AverageExchangeRatesAnalyzer.Helpers.Excel
{
    public static class ExcelExporterHelper
    {
        public static Columns AutoSizeCells(SheetData sheetData)
        {
            var maxColWidth = GetMaxCharacterWidth(sheetData);

            Columns columns = new Columns();
            //this is the width of my font - yours may be different
            double maxWidth = 7;
            foreach (var item in maxColWidth)
            {
                //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
                double width = Math.Truncate((item.Value * maxWidth + 5) / maxWidth * 256) / 256;
                Column col = new Column() { BestFit = true, Min = (UInt32)(item.Key + 1), Max = (UInt32)(item.Key + 1), CustomWidth = true, Width = (DoubleValue)width };
                columns.Append(col);
            }

            return columns;
        }

        public static Dictionary<int, int> GetMaxCharacterWidth(SheetData sheetData)
        {
            Dictionary<int, int> maxColWidth = new Dictionary<int, int>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 1, 5, 6, 7, 8, 14 };
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8, 14 };
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? cell.InnerText : cell.CellValue.InnerText;
                    var cellTextLength = cellValue.Length;

                    if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                    {
                        int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                        cellTextLength += (3 + thousandCount);
                    }

                    if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                    {
                        cellTextLength += 1;
                    }

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }

        public static Stylesheet CreateStyleSheet()
        {
            Stylesheet stylesheet = new Stylesheet();

            #region Number format
            uint DIGITS8_FORMAT = 164;
            var numberingFormats = new NumberingFormats();
            numberingFormats.Append(new NumberingFormat // Datetime format
            {
                NumberFormatId = UInt32Value.FromUInt32(DIGITS8_FORMAT),
                FormatCode = "0.00000000"
            });
            numberingFormats.Count = UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
            #endregion

            #region Fonts
            var fonts = new Fonts();
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font()  // Font index 0 - default
            {
                FontName = new FontName { Val = StringValue.FromString("Calibri") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) }
            });
            fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font()  // Font index 1
            {
                FontName = new FontName { Val = StringValue.FromString("Arial") },
                FontSize = new FontSize { Val = DoubleValue.FromDouble(11) },
                Bold = new Bold()
            });
            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
            #endregion

            #region Borders
            var borders = new Borders();
            borders.Append(new Border   // Border index 0: no border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Append(new Border    //Boarder Index 1: All
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            });
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            #endregion

            #region Cell Style Format
            var cellStyleFormats = new CellStyleFormats();
            cellStyleFormats.Append(new CellFormat  // Cell style format index 0: no format
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            });
            cellStyleFormats.Count = UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
            #endregion

            #region Cell format

            #endregion

            stylesheet.Append(fonts);
            stylesheet.Append(CustomFills.GetCustomFills());
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(CustomCellFormats.GetCustomCellFormats(DIGITS8_FORMAT));

            #region Cell styles
            var css = new CellStyles();
            css.Append(new CellStyle
            {
                Name = StringValue.FromString("Normal"),
                FormatId = 0,
                BuiltinId = 0
            });
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            stylesheet.Append(css);
            #endregion

            var dfs = new DifferentialFormats { Count = 0 };
            stylesheet.Append(dfs);
            var tss = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            };
            stylesheet.Append(tss);

            return stylesheet;
        }
        public static uint FindNumberCellStyleIndex(List<string> sortedAverageExchangeRates, string rateCode)
        {
            if (sortedAverageExchangeRates[0] == rateCode)
            {
                return 5;
            }
            else if (sortedAverageExchangeRates[1] == rateCode)
            {
                return 6;
            }
            else if (sortedAverageExchangeRates[sortedAverageExchangeRates.Count - 1] == rateCode)
            {
                return 7;
            }
            else if (sortedAverageExchangeRates[sortedAverageExchangeRates.Count - 2] == rateCode)
            {
                return 8;
            }
            else
            {
                return 2;
            }
        }

        public static uint FindTextCellStyleIndex(List<string> sortedAverageExchangeRates, string rateCode, bool dayOff = false)
        {
            if (sortedAverageExchangeRates[0] == rateCode)
            {
                return dayOff ? UInt32Value.FromUInt32(13) : 9;
            }
            else if (sortedAverageExchangeRates[1] == rateCode)
            {
                return dayOff ? UInt32Value.FromUInt32(14) : 10;
            }
            else if (sortedAverageExchangeRates[sortedAverageExchangeRates.Count - 1] == rateCode)
            {
                return dayOff ? UInt32Value.FromUInt32(15) : 11;
            }
            else if (sortedAverageExchangeRates[sortedAverageExchangeRates.Count - 2] == rateCode)
            {
                return dayOff ? UInt32Value.FromUInt32(16) : 12;
            }
            else
            {
                return dayOff ? UInt32Value.FromUInt32(17) : 3;
            }
        }

    }
}
