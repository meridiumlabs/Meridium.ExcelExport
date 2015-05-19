using System.Collections.Generic;
using System.IO;

namespace Meridium.ExcelExport {
    /// <summary>
    /// Convenience methods for excel exports.
    /// </summary>
    public static class  ExcelExtensions {

        /// <summary>
        /// Exports the sequence of data to an excel file.
        /// </summary>
        /// <typeparam name="TData">The type that represents a data row.</typeparam>
        /// <param name="self">The sequence of data to export</param>
        /// <param name="isPoco">When true: export all public properties.</param>
        /// <returns>The contents of an excel file.</returns>
        public static byte[] ToExcel<TData>(this IEnumerable<TData> self, bool isPoco = false) {
            var exporter = new ExcelExporter<TData>(isPoco);
            return exporter.Export(self);
        }

        /// <summary>
        /// Write the sequence of data to the specified stream.
        /// </summary>
        /// <typeparam name="TData">The type that represents a data row.</typeparam>
        /// <param name="self">The sequence of data to export</param>
        /// <param name="output">The output stream to write to</param>
        /// <param name="isPoco">When true: export all public properties.</param>
        public static void WriteExcel<TData>(this IEnumerable<TData> self, Stream output, bool isPoco = false) {
            var exporter = new ExcelExporter<TData>(isPoco);
            var data = exporter.Export(self);
            output.Write(data, 0, data.Length);
        }
    }
}
