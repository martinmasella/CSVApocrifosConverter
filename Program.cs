using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Data;
using System.Linq;
using System.Globalization;
using ClosedXML.Excel;
using ExcelDataReader;

namespace CSVToXLSXConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string zipUrl = "https://servicioscf.afip.gob.ar/Facturacion/facturasApocrifas/DownloadFile.aspx";
            string zipFilePath = "downloaded.zip";
            string csvFilePath = "extracted.csv";
            string xlsxFilePath = $"{DateTime.Now.ToString("ddMMyyyyhhmmss")}.xlsx";


            ToLog("Iniciando...");
            DownloadFile(zipUrl, zipFilePath);
            ExtractCsvFromZip(zipFilePath, csvFilePath);
            ConvertCsvToXlsx(csvFilePath, xlsxFilePath);
            ToLog("Fin.");
            Console.WriteLine($"Archivo XLSX generado: {xlsxFilePath}");
        }

        static void ToLog(string texto)
        {
            Console.WriteLine(texto);
        }

        static void DownloadFile(string url, string outputPath)
        {
            ToLog("Iniciando descarga...");
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(url, outputPath);
            }
            ToLog("Fin descarga.");

        }

        static void ExtractCsvFromZip(string zipFilePath, string outputCsvFilePath)
        {
            ToLog("Iniciando descompresión...");

            using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
            {
                //var csvEntry = archive.Entries.FirstOrDefault(e => e.FullName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase));
                var csvEntry = archive.Entries.FirstOrDefault();
                if (csvEntry != null)
                {
                    csvEntry.ExtractToFile(outputCsvFilePath, true);
                }
            }
            ToLog("Fin descompresión.");
        }

        static void ConvertCsvToXlsx(string csvFilePath, string xlsxFilePath)
        {
            ToLog("Iniciando volcado de filas...");
            // Column headers
            string[] headers = { "CUIT", "FechaCondicionApocrifo", "FechaPublicacion" };

            using (var stream = File.Open(csvFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                {
                    AutodetectSeparators = new char[] { ',' },
                    LeaveOpen = false,
                }))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = false // We do not want to use the first row as header.
                        }
                    });

                    var dataTable = result.Tables[0];

                    // Remove the first three rows
                    for (int i = 0; i < 3; i++)
                    {
                        dataTable.Rows[i].Delete();
                    }
                    dataTable.AcceptChanges();

                    // Create a new DataTable with the desired headers
                    var newTable = new DataTable();
                    foreach (var header in headers)
                    {
                        newTable.Columns.Add(header);
                    }

                    // Copy only the first three columns from the original DataTable to the new DataTable
                    foreach (DataRow row in dataTable.Rows)
                    {
                        var newRow = newTable.NewRow();
                        for (int col = 0; col < headers.Length; col++)
                        {
                            newRow[col] = row[col];
                        }
                        newTable.Rows.Add(newRow);
                    }
                    ToLog("Abriendo Excel...");

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add(newTable, "Apócrifos");
                        // Set column widths
                        worksheet.Column(1).Width = 30; // CUIT
                        worksheet.Column(2).Width = 30; // FechaCondicionApocrifo
                        worksheet.Column(3).Width = 30; // FechaPublicacion
                        workbook.SaveAs(xlsxFilePath);
                    }
                    System.Diagnostics.Process.Start(xlsxFilePath);
                }
            }
        }
    }
}
