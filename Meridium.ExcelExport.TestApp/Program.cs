using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace Meridium.ExcelExport.TestApp {
    class Program {
        static void Main(string[] args) {
            var items = new List<MyTestData> {
                new MyTestData {Name = "Ford", Date = DateTime.Now, ProductId = "0003423", SomeNumber = 342.454234m},
                new MyTestData {Name = "Chevy", Date = DateTime.Now, ProductId = "03366423", SomeNumber = 0m},
                new MyTestData {Name = "Volvo", Date = DateTime.Now, ProductId = null, SomeNumber = 1337m},
            };

            var pocos = new List<MyPoco> {
                new MyPoco {DatabaseID = -2, IsFreeFromSin = false, UTCDate = DateTime.UtcNow},
                new MyPoco {DatabaseID = 34, ISO8953Date = "Mon 2/4 2012"},
                new MyPoco {DatabaseID = 0, IsFreeFromSin = true, ISO8953Date = "Tue 9/4 2013"},
                new MyPoco {DatabaseID = 34342, IsFreeFromSin = false, UTCDate = DateTime.UtcNow},
            };

            var temppath = Path.GetTempPath();
            Console.Out.WriteLine("Temppath: {0}", temppath );
            using (var file = File.Create(Path.Combine(temppath, "test.xlsx"))) {
                //items.WriteExcel(file);
                pocos.WriteExcel(isPoco: true, output: file);
            }

            Console.Out.WriteLine("Done!");
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

    class MyPoco {
        public int DatabaseID { get; set; }
        public string ISO8953Date { get; set; }
        public DateTime UTCDate { get; set; }
        public bool IsFreeFromSin { get; set; }
    }
}
