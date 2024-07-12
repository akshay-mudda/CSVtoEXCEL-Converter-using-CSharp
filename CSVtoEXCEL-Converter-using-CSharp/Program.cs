using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using OfficeOpenXml;

namespace CSVtoEXCEL_Converter_using_CSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFolder = ConfigurationManager.AppSettings["SourceFolder"];
            string destinationFolder = ConfigurationManager.AppSettings["DestinationFolder"];
            string connectionString = ConfigurationManager.ConnectionStrings["FileMoveHistoryDB"].ConnectionString;

            // Set EPPlus license context for non-commercial use
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Check if source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder {sourceFolder} does not exist.");
                return;
            }

            // Ensure destination folder exists
            if (!Directory.Exists(destinationFolder))
            {
                try
                {
                    Directory.CreateDirectory(destinationFolder);
                    Console.WriteLine($"Created folder: {destinationFolder}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error creating folder {destinationFolder}: {ex.Message}");
                    return;
                }
            }

            // Get all CSV files from the source folder
            string[] csvFiles = Directory.GetFiles(sourceFolder, "*.csv");
            List<FileMoveRecord> moveRecords = new List<FileMoveRecord>();

            foreach (string file in csvFiles)
            {
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string destFile = Path.Combine(destinationFolder, $"{fileName}.xlsx");

                    // Convert CSV to Excel
                    if (ConvertCsvToExcel(file, destFile))
                    {
                        File.Delete(file);
                        Console.WriteLine($"Converted and moved {file} to {destFile}");
                        moveRecords.Add(new FileMoveRecord(fileName, "Excel", sourceFolder, destFile));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error moving file {file}: {ex.Message}");
                }
            }

            // Insert move history into SQL table
            DAL.InsertMoveHistory(moveRecords, connectionString);

            // Print summary
            Console.WriteLine("\nMove Process Summary:");
            Console.WriteLine($"Total files converted and moved: {moveRecords.Count}");
        }

        static bool ConvertCsvToExcel(string sourceFile, string destinationFile)
        {
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    var format = new ExcelTextFormat { Delimiter = ',', EOL = "\r\n" };

                    // Load CSV data and ensure it's loaded row by row
                    string[] csvLines = File.ReadAllLines(sourceFile);
                    for (int rowIndex = 0; rowIndex < csvLines.Length; rowIndex++)
                    {
                        string[] rowValues = csvLines[rowIndex].Split(format.Delimiter);
                        for (int colIndex = 0; colIndex < rowValues.Length; colIndex++)
                        {
                            worksheet.Cells[rowIndex + 1, colIndex + 1].Value = rowValues[colIndex];
                        }
                    }

                    package.SaveAs(new FileInfo(destinationFile));
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting file {sourceFile} to Excel: {ex.Message}");
                return false;
            }
        }
    }

    public class FileMoveRecord
    {
        public string FileName { get; }
        public string FileType { get; }
        public string SourcePath { get; }
        public string DestinationPath { get; }

        public FileMoveRecord(string fileName, string fileType, string sourcePath, string destinationPath)
        {
            FileName = fileName;
            FileType = fileType;
            SourcePath = sourcePath;
            DestinationPath = destinationPath;
        }
    }
}