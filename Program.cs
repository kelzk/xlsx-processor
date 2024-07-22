using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Net.Http;
using System.Data;

using ExcelDataReader;
using ExcelNumberFormat;

using HtmlAgilityPack;

namespace ConsoleApp1
{
    public class Program
    {
        public static string BaseUriToScrape {  get; set; } = "https://www.abs.gov.au";
        public static string FirstPageToScrape {  get; set; } = "/statistics/labour/employment-and-unemployment/labour-force-australia";
        public static string FirstPageNode {  get; set; } = "//div[@id='block-views-block-topic-releases-listing-topic-latest-release-block']//a[@href]";
        public static string FirstPageNodeAttribute {  get; set; } = "href";
        public static string SecondPageNode { get; set; } = "//div[@class='file-description-link-formatter']//a[@href]";
        public static string SecondPageNodeAttribute {  get; set; } = "href";
        public static string XlsxFileName { get; set; } = "data.xlsx";
        public static string Folder { get; set; } = "Downloads";
        public static string CsvFileName { get; set; } = "data.csv";
        public static string PrimaryKey { get; set; } = "Series ID";

        public static string CrawlWebpage(HtmlWeb htmlWeb, string pageToScrape, string nodeFilter, string nodeAttribute)
        {
            HtmlDocument pageDocument = htmlWeb.Load(BaseUriToScrape + pageToScrape);
            HtmlNodeCollection nodeCollection = pageDocument.DocumentNode.SelectNodes(nodeFilter);

            string uri = nodeCollection[0].GetAttributeValue(nodeAttribute, "");
            return uri;
        }

        public static async Task DownloadAndSaveFileAsync(string downloadUri, string directory, string fileName)
        {
            Directory.CreateDirectory($"../../{directory}/");
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Other");
            Stream stream = await httpClient.GetStreamAsync(downloadUri);
            FileStream fileStream = new FileStream($"../../{directory}/{fileName}", FileMode.Create);
            await stream.CopyToAsync(fileStream);
            httpClient.Dispose();
            stream.Dispose();
            fileStream.Dispose();

        }

        public static DataTable ReadXlsxFile(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            reader.NextResult();
            DataTable dataTable = new DataTable();

            for (int i = 0; i < reader.FieldCount; i++)
            {
                DataColumn dataColumn = new DataColumn();
                dataColumn.ColumnName = i.ToString();
                dataColumn.DataType = typeof(string);
                dataTable.Columns.Add(dataColumn);
            }
            dataTable.PrimaryKey = new DataColumn[] { dataTable.Columns[0] };
            while (reader.Read())
            {
                DataRow dataRow = dataTable.NewRow();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    // Substitute "Null" at Excel top-left corner for primaryKey rule
                    var value = reader.GetValue(i) ?? "Null";
                    var formatString = reader.GetNumberFormatString(i);
                    dataRow[i] = value;

                    if (formatString != null)
                    {
                        var format = new NumberFormat(formatString);
                        dataRow[i] = format.Format(value, CultureInfo.InvariantCulture);
                    }
                    
                }
                dataTable.Rows.Add(dataRow);
            }
            return dataTable;
        }

        public static DataTable RemoveRowsBefore(DataTable dataTable, string primaryKey)
        {
            int primaryKeyIndex = dataTable.Rows.IndexOf(dataTable.Rows.Find(primaryKey) ?? dataTable.NewRow());
            DataTable rowRemovedDataTable = dataTable;
            if (primaryKeyIndex != -1)
            {
                for (int i = 0; i < primaryKeyIndex; i++)
                {
                    rowRemovedDataTable.Rows.RemoveAt(0);
                }
            }
            return rowRemovedDataTable;
        }

        public static DataTable TransposeDataTable(DataTable dataTable)
        {
            DataTable transposedDataTable = new DataTable();
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                DataColumn dataColumn = new DataColumn();
                dataColumn.DataType = typeof(string);
                dataColumn.ColumnName = i.ToString();
                transposedDataTable.Columns.Add(dataColumn);
            }

            for (int column = 0; column < dataTable.Columns.Count; column++)
            {
                DataRow dataRow = transposedDataTable.NewRow();
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    dataRow[row] = dataTable.Rows[row][column];
                }
                transposedDataTable.Rows.Add(dataRow);

            }

            return transposedDataTable;
        }

        public static void SaveAsCsvFile(DataTable dataTable, string filePath)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (DataRow row in dataTable.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                stringBuilder.AppendLine(string.Join(",", fields));
            }
            File.WriteAllText(filePath, stringBuilder.ToString());
        }

        public static async Task Main(string[] args)
        {
            HtmlWeb htmlWeb = new HtmlWeb();
            string uri = CrawlWebpage(htmlWeb, FirstPageToScrape, FirstPageNode, FirstPageNodeAttribute);
            string downloadUri = CrawlWebpage(htmlWeb, uri, SecondPageNode, SecondPageNodeAttribute);
            string completeUri = BaseUriToScrape + downloadUri;
            await DownloadAndSaveFileAsync(completeUri, Folder, XlsxFileName);
            DataTable dataTable = ReadXlsxFile($"../../{Folder}/{XlsxFileName}");
            DataTable rowRemovedDataTable = RemoveRowsBefore(dataTable, PrimaryKey);
            DataTable transposedDataTable = TransposeDataTable(rowRemovedDataTable);
            SaveAsCsvFile(transposedDataTable, $"../../{Folder}/{CsvFileName}");

            Console.ReadKey();
        }
    }
}
