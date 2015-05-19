using System.Collections.Generic;
using System.IO;
using ClosedXML;
using ClosedXML.Excel;

namespace Meridium.ExcelExport {
    /// <summary>
    /// Wrapper of the SpreadsheetGear Excel library.
    /// </summary>
    /// <typeparam name="TData">The type that represents a row of data</typeparam>
    public class ExcelExporter<TData> {
        /// <summary>
        /// Initializes a new instance of the exporter.
        /// </summary>
        /// <param name="isPoco">Set to true to output all public properties of the data objects</param>
        public ExcelExporter(bool isPoco = false) {
            _rowSpec = new RowSpec<TData>(isPoco);
        }

        /// <summary>
        /// Export the specified collection of data items to an excel file.
        /// </summary>
        /// <param name="data">The data to export</param>
        /// <returns>An array of bytes representing the Excel file.</returns>
        public byte[] Export(IEnumerable<TData> data) {
            var book = new XLWorkbook();
            var sheet = book.Worksheets.Add("Sheet1");

            for (var i = 0; i < _rowSpec.ColCount; i++) {
                sheet.Cell(1, i+1).Value = _rowSpec.Headings[i];
            }

            var row = 2;
            foreach (var item in data) {
                var cells = _rowSpec.GetCells(item);
                for (var col = 0; col < _rowSpec.ColCount; col++) {
                    if (cells[col].IsText) {
                        sheet.Cell(row, col+1).Style.NumberFormat.Format = "@";
                    }
                    sheet.Cell(row, col+1).Value = cells[col].Value;
                }
                ++row;
            }

            sheet.Columns().AdjustToContents();

            var saveStream = new MemoryStream();
            book.SaveAs(saveStream);

            return saveStream.ToArray();
        }

        public const string ContentType = "application/vnd.ms-excel";

        private readonly RowSpec<TData> _rowSpec;
    }
}