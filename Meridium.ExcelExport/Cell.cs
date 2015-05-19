namespace Meridium.ExcelExport {
    /// <summary>
    /// Represents the data in an Excel cell.
    /// </summary>
    internal class Cell {
        
        /// <summary>
        /// The value in the cell.
        /// </summary>
        public object Value    { get; set; }

        /// <summary>
        /// Indicates whether the value in the cell always should be interpreted as text.
        /// </summary>
        public bool   IsText   { get; set; }
    }
}