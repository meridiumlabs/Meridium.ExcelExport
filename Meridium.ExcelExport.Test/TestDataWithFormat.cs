using System;

namespace Meridium.ExcelExport.Test {
    public class TestDataWithFormat {
        [ExcelCell(Column = 1, Format = "yyyy-MM-dd")]
        public DateTime Date { get; set; }

        [ExcelCell(Column = 2, Format = "0.##" )]
        public double Number { get; set; }
    }
}