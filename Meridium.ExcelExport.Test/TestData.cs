using System;

namespace Meridium.ExcelExport.Test {
    public class TestData {
        [ExcelCell(TreatAsText = true)]
        public string Foo { get; set; }

        [ExcelCell(Column = 1,  Heading = "My col heading")]
        public int Bar { get; set; }

        [ExcelCell(Column = 2)]
        public DateTime Quxx { get; set; }

        public string Baz { get; set; }
    }
}