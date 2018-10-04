using System;
using System.IO;
using System.Data;

using ExcelDataReader;

namespace xlsx_playground
{
    class Program
    {
        public class Reader
        {
            private string xlsxFile = "1048576_rows.xlsx";

            public Reader()
            {
                System.Text.Encoding.RegisterProvider(
                    System.Text.CodePagesEncodingProvider.Instance);
            }

            public int RowsUsingReader()
            {
                using (var stream = File.Open(xlsxFile, FileMode.Open,
                    FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        int rows = 0;
                        do
                        {
                            while (reader.Read())
                            {
                                rows++;
                            }
                        }
                        while (reader.NextResult());

                        return rows;
                    }
                }
            }

            public int RowsUsingDataSet()
            {
                using (var stream = File.Open(xlsxFile, FileMode.Open,
                    FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();

                        return result.Tables[0].Rows.Count;
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            var r = new Reader();
            Console.WriteLine($"RowsUsingReader: {r.RowsUsingReader()}");
            Console.WriteLine($"RowsUsingDataSet: {r.RowsUsingDataSet()}");
        }
    }
}
