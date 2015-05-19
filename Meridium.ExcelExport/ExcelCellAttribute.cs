using System;

namespace Meridium.ExcelExport {
    /// <summary>
    /// Defines how the decorated property will be represented in an Ecxel cell.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelCellAttribute : Attribute {
        /// <summary>
        /// Initializes the attribute with default values. 
        /// </summary>
        public ExcelCellAttribute() {
            Column = int.MaxValue;
        }

        /// <summary>
        /// The order of the column where the property will be output.
        /// </summary>
        public int Column { get; set; }

        /// <summary>
        /// The column heading, defaults to the property name.
        /// </summary>
        public string Heading { get; set; }

        /// <summary>
        /// A format string to apply to the property value. 
        /// Think <c>string.Format("{0:Format}", PropertyValue)</c> 
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Value should always be treated as text by Excel, even if it looks
        /// like a number or a date or whatever.
        /// </summary>
        public bool TreatAsText { get; set; }
    }
}