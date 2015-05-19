using System;
using System.Linq;
using System.Reflection;

namespace Meridium.ExcelExport {

    internal class RowSpec<TType> {
        public Type DataType { get; private set; }

        public string[] Headings { get; private set; }

        public int ColCount { get { return _properties.Length; } }

        public RowSpec(bool isPoco = false) {
            DataType = typeof(TType);

            if (isPoco) {
                PopulateUsingPublicProperties();
            }
            else {
                PopulateUsingAttributes();
            }

            Headings = _properties.Select(p => p.Attribute.Heading ?? p.Property.Name).ToArray();
        }

        public Cell[] GetCells(TType data) {
            const string formatTemplate = @"{{0:{0}}}";
            return _properties
                .Select(p => {
                    var value = p.Property.GetValue(data, BindingFlags.Default, null, null, null);
                    if (p.Attribute.Format != null) {
                        var format = string.Format(formatTemplate, p.Attribute.Format);
                        value = string.Format(format, p.Property.GetValue(data, BindingFlags.Default, null, null, null));
                    }
                    return new Cell {
                        Value = value,
                        IsText = p.Attribute.TreatAsText
                    };
                })
                .ToArray();
        }

        private void PopulateUsingAttributes() {
            _properties = DataType.GetProperties()
                .Where(p => p.GetCustomAttributes(typeof(ExcelCellAttribute), false).Length == 1)
                .Select( p => new ExcelCellProperty {
                    Property = p,
                    Attribute = (ExcelCellAttribute) p.GetCustomAttributes(typeof(ExcelCellAttribute), false)[0]
                })
                .OrderBy( p => p.Attribute.Column)
                .ToArray();
        }

        private void PopulateUsingPublicProperties() {
            _properties = DataType.GetProperties(BindingFlags.Public | BindingFlags.Instance)       
                .Select( p => new ExcelCellProperty {
                    Property = p,
                    Attribute = new ExcelCellAttribute { Heading = DeCamelizer.Split(p.Name) }
                })
                .OrderBy( p => p.Attribute.Column)
                .ToArray();
        }

        private class ExcelCellProperty {
            public PropertyInfo Property { get; set; }
            public ExcelCellAttribute Attribute { get; set; }
        }

        private ExcelCellProperty[] _properties;
    }
}
