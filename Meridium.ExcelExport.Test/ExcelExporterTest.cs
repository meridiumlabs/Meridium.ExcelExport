using NFluent;
using Ploeh.AutoFixture;
using Xunit;

namespace Meridium.ExcelExport.Test {
    public class ExcelExporterTest {

        public class ExportData_method {
            private Fixture fixture;
            public ExportData_method() {
                fixture = new Fixture();
            }

            [Fact]
            public void should_treat_a_class_different_depending_on_the_isPoco_flag() {
                var testdata = fixture.Create<TestData>();

                var exporter1 = new ExcelExporter<TestData>(false);
                var excel1 = exporter1.Export( new []{testdata});
                var exporter2 = new ExcelExporter<TestData>(true);
                var excel2 = exporter2.Export( new []{testdata});

                Check.That(excel1).Not.ContainsExactly(excel2);
            }

            [Fact]
            public void should_write_serialize_data_as_an_excel_file() {
                var exporter = new ExcelExporter<TestData>();
                var data = fixture.CreateMany<TestData>();

                var exceldata = exporter.Export(data);

                Check.That(exceldata).Not.IsEmpty();
            }
        }
    }
}