
using System;
using System.Globalization;
using System.Linq;
using NFluent;
using Xunit;

namespace Meridium.ExcelExport.Test {
    public class RowSpecTest {
        public class Contstructor {
            [Fact]
            public void should_create_a_new_RowSpec_instance_using_attributes_by_default() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.DataType).IsEqualTo(typeof (TestData));
                Check.That(spec.ColCount).Equals(3);
            }

            [Fact]
            public void should_create_a_new_RowSpec_instance_using_public_properties_when_specified() {
                var spec = new RowSpec<UndecoratedTestData>(isPoco: true);
                
                Check.That(spec.DataType).IsEqualTo(typeof (UndecoratedTestData));
                Check.That(spec.ColCount).Equals(3);
            }
        }



        public class Heading_property {
            [Fact]
            public void should_use_property_name_when_heading_is_not_set_in_cell_attribute() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.Headings).Contains("Foo");
            }

            [Fact]
            public void should_use_specified_heading_name_when_set_in_cell_attribute() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.Headings).Contains("My col heading");
            }

            [Fact]
            public void should_skip_properties_without_ExcelCell_attribute() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.Headings).Not.Contains("Baz");
            }

            [Fact]
            public void should_order_headings_by_oridinal_attribute_value_putting_unspecified_last() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.Headings).ContainsExactly("My col heading", "Quxx", "Foo");
            }

            [Fact]
            public void should_have_the_same_length_as_the_ColCount_property() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.Headings.Length).Equals(spec.ColCount);
            }

            [Fact]
            public void should_contain_splitted_properties_when_isPoco_is_defined() {
                var spec = new RowSpec<UndecoratedTestData>(isPoco: true);

                Check.That(spec.Headings).Contains("Foo Bar", "Data ID", "ISO 3434 Date");
            }
        }

        public class ColCount_property {
            [Fact]
            public void should_return_the_number_of_properties_that_will_be_excel_cells() {
                var spec = new RowSpec<TestData>();

                Check.That(spec.ColCount).Equals(3);
            }
        }

        public class GetCells_method {
            public GetCells_method() {
                System.Threading.Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            }
            [Fact]
            public void should_get_the_cells_representing_the_property_values() {
                var date = new DateTime(1985, 4, 3);
                var spec = new RowSpec<TestData>();

                var data = new TestData {
                    Foo = "Foo bar baz",
                    Bar = 42,
                    Baz = "Barabbas",
                    Quxx = date
                };

                var cells = spec.GetCells(data);

                Check.That(cells.Properties("Value")).ContainsExactly<object>(42, date, "Foo bar baz");
            }

            [Fact]
            public void should_return_a_collection_of_same_length_as_the_ColCount_property() {
                var data = new TestData();
                var spec = new RowSpec<TestData>();

                var result = spec.GetCells(data);

                Check.That(result.Length).Equals(spec.ColCount);
            }

            [Fact]
            public void should_try_to_format_the_value_as_a_string_when_format_is_specified() {
                var date = new DateTime(1989, 3, 2);
                var data = new TestDataWithFormat {Date = date, Number = 43.3939};

                var spec = new RowSpec<TestDataWithFormat>();

                var cells = spec.GetCells(data);

                Check.That(cells.Properties("Value")).ContainsExactly("1989-03-02", "43.39");

            }

            [Fact]
            public void should_set_the_text_flag_of_the_cell_when_TreatAsText_property_is_set() {
                var testdata = new TestData {Foo = "0005"};

                var spec = new RowSpec<TestData>();

                var cells = spec.GetCells(testdata);

                Check.That(cells.Properties("IsText")).Contains(true).Once();
            }
        }
    }
}
