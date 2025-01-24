using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AverageExchangeRatesAnalyzer.Helpers.Excel
{
    public static class CustomCellFormats
    {
        public static CellFormats GetCustomCellFormats(uint DIGITS8_FORMAT)
        {
            CellFormats cellFormats = new CellFormats();
            cellFormats.Append(new CellFormat());    // Cell format index 0
            cellFormats.Append(new CellFormat   // CellFormat index 1
            {
                NumberFormatId = 14,        // 14 = 'mm-dd-yy'. Standard Date format;
                FontId = 1,
                FillId = 2,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 2
            {
                NumberFormatId = DIGITS8_FORMAT,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 3
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 3,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat   // CellFormat index 4 if day off
            {
                NumberFormatId = 14,        // 14 = 'mm-dd-yy'. Standard Date format;
                FontId = 1,
                FillId = 3,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 5 if highest average
            {
                NumberFormatId = DIGITS8_FORMAT,
                FontId = 0,
                FillId = 4,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 6 if second highest average
            {
                NumberFormatId = DIGITS8_FORMAT,
                FontId = 0,
                FillId = 5,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 7 if worst average
            {
                NumberFormatId = DIGITS8_FORMAT,
                FontId = 0,
                FillId = 6,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 8 if second worst average
            {
                NumberFormatId = DIGITS8_FORMAT,
                FontId = 0,
                FillId = 7,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 9 if highest average
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 4,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 10 if second highest average
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 5,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 11 if worst average
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 6,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 12 if second worst average
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 7,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 13 if highest average day off
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 8,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 14 if second highest average day off
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 9,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 15 if worst average day off
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 10,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 16 if second worst average day off
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 11,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Append(new CellFormat    // Cell format index 17 if day off
            {
                NumberFormatId = 0, // 0 overall
                FontId = 0,
                FillId = 12,
                BorderId = 1,
                FormatId = 0,
                Alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center },
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            });
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);

            return cellFormats;
        }
    }
}
