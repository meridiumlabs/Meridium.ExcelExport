using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NFluent;
using Ploeh.AutoFixture;
using Xunit;

namespace Meridium.ExcelExport.Test {
    public class ExcelExtensionsTest {
        public class ToExcel_method {
            private Fixture fixture;
            public ToExcel_method() {
                fixture = new Fixture();
            }

            [Fact]
            public void should_generate_excel_data_as_the_ExcelExporter_Export_method() {
                var data = fixture.CreateMany<MyTestData>().ToArray();

                var exporterExcel = new ExcelExporter<MyTestData>().Export(data);

                var toExcelData = data.ToExcel();

                Check.That(toExcelData).HasSize(exporterExcel.Length);
            }

            [Fact]
            public void should_treat_the_class_as_poco_when_flag_is_set() {
                var data = fixture.CreateMany<MyTestData>();

                var exporterExcel = new ExcelExporter<MyTestData>(true).Export(data);

                var toExcelData = data.ToExcel(isPoco: true);

                Check.That(toExcelData).HasSize(exporterExcel.Length);
            }
        }

        public class WriteExcel_method {
            private Fixture fixture;
            public WriteExcel_method() {
                fixture = new Fixture();
            }

            [Fact]
            public void should_write_the_excel_data_to_the_specified_stream() {
                var data = fixture.CreateMany<MyTestData>();

                var exporterExcel = new ExcelExporter<MyTestData>().Export(data);

                var stream = new MemoryStream();

                data.WriteExcel(stream);

                Check.That(stream.ToArray()).HasSize(exporterExcel.Length);
            }

            [Fact]
            public void should_treat_the_object_as_a_poco_when_flag_is_set() {
                var data = fixture.CreateMany<MyTestData>();

                var exporterExcel = new ExcelExporter<MyTestData>(true).Export(data);

                var stream = new MemoryStream();

                data.WriteExcel(stream, isPoco: true);

                Check.That(stream.ToArray()).HasSize(exporterExcel.Length);
                
            }
        }

        class MyTestData {
            [ExcelCell(Column = 1, Heading = "Namn")]
            public string Name { get; set; }
            [ExcelCell(Column = 2, Heading = "Datum för händelse", Format = "D")]
            public DateTime Date { get; set; }
            [ExcelCell(Column = 3, Heading = "Lön", Format = "0.00")]
            public decimal SomeNumber { get; set; }
            [ExcelCell(Column = 4, Heading = "Produkt-id", TreatAsText = true)]
            public string ProductId { get; set; }

        }
    }
}
