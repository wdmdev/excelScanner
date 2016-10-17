using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_scanner
{
    public class Program
    {
        static void Main(string[] args)
        {
            var searchWords = args;
            Console.WriteLine("Input file directory path...");
            Console.Write("> ");

            var path = Console.ReadLine();

            var files = Directory.GetFiles(path);
            var fileStreams = RetrieveFileStreams(files).ToList();

            for(var i = 0; i < fileStreams.Count; i++)
            {
                SearchExcelFile(fileStreams[i], searchWords);
                Console.Clear();
                Console.WriteLine($"{(i+1/fileStreams.Count)*100}% completed");
            }
        }

        private static IEnumerable<FileStream> RetrieveFileStreams(string[] filePaths)
        {
            foreach (var path in filePaths)
            {
                if (path.EndsWith(".xlsx"))
                {
                    yield return File.Open(path, FileMode.Open, FileAccess.Read);
                }
            }
        }

        private static void SearchExcelFile(FileStream stream, string[] searchStrings)
        {
            //Choose one of either 1 or 2
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            //Choose one of either 3, 4, or 5
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            //DataSet result = excelReader.AsDataSet();

            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;
            //DataSet result = excelReader.AsDataSet();

            //5. Data Reader methods
            var sBuilder = new StringBuilder();
            var columnAmount = 0;
            var row = 1;

            while (excelReader.Read())
            {
                columnAmount = excelReader.FieldCount;

                foreach (var s in searchStrings)
                {
                    for (var i = 0; i < columnAmount; i++)
                    {
                        var field = excelReader.GetValue(i);
                        if (field == null)
                        {
                            field = "";
                        }

                        if (field.ToString().ToLower().Contains(s.ToLower()))
                        {
                            sBuilder.AppendLine($"Match found on word: {s} in file: {stream.Name}, row: {row}");
                        }
                    }
                }
                row++;
            }

            File.AppendAllText(
                            $"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}/results.txt", sBuilder.ToString());

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
    }
}
