using NFluent;
using Xunit;

namespace Meridium.ExcelExport.Test {
    public class DeCamellizerTest {
        public class Split_method {
            [Fact]
            public void should_return_the_empty_string_when_arg_is_null() {
                var result = DeCamelizer.Split(null);
                Check.That(result).IsEmpty();
            }

            [Fact]
            public void should_return_the_empty_string_when_arg_is_empty() {
                var result = DeCamelizer.Split("");
                Check.That(result).IsEmpty();
            }

            [Fact]
            public void should_return_the_empty_string_when_arg_is_only_ws() {
                var result = DeCamelizer.Split(" \r\n\t");
                Check.That(result).IsEmpty();
            }

            [Fact]
            public void should_split_on_first_uppercase_after_lowercase() {
                var result = DeCamelizer.Split("FooBarBaz");
                Check.That(result).Equals("Foo Bar Baz");
            }

            [Fact]
            public void should_treat_ranges_of_uppercase_as_abbrevations_inside_word() {
                var result = DeCamelizer.Split("DefaultHTMLParser");
                Check.That(result).Equals("Default HTML Parser");
            }

            [Fact]
            public void should_treat_ranges_of_uppercase_as_abbrevations_at_start_of_word() {
                var result = DeCamelizer.Split("SQLDatabase");
                Check.That(result).Equals("SQL Database");
            }

            [Fact]
            public void should_treat_ranges_of_uppercase_as_abbrevation_at_end_of_word() {
                var result = DeCamelizer.Split("MainDB");
                Check.That(result).Equals("Main DB");
            }

            [Fact]
            public void should_split_on_first_non_letter_after_letter() {
                var result = DeCamelizer.Split("Foo42Bar");
                Check.That(result).Equals("Foo 42 Bar");
            }

            [Fact]
            public void should_not_split_on_uppercase_before_non_letter() {
                var result = DeCamelizer.Split("ISO8953");
                Check.That(result).Equals("ISO 8953");
            }
        }
    }
}