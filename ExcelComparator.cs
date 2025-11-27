using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TestProject1

{
    class ExcelComparator
    {
        static void Main(string[] args)
        {
            // 1. Configuration - Set your file paths here
            string stPath = @"C:\Users\Tyagi\vsp\testProject1\reports\ST.xlsx"; // Source File
            string ptPath = @"C:\Users\Tyagi\vsp\testProject1\reports\PT.xlsx"; // Primary Target
            string rtPath = @"C:\Users\Tyagi\vsp\testProject1\reports\RT.xlsx"; // Reference Target (Backup)

            Console.WriteLine("Starting Excel Comparison (using OpenXML)...");

            try
            {
                // 2. Load DataTables using OpenXML
                DataTable stTable = ReadExcelOpenXml(stPath);
                DataTable ptTable = ReadExcelOpenXml(ptPath);
                DataTable rtTable = ReadExcelOpenXml(rtPath);

                if (stTable == null || ptTable == null || rtTable == null)
                {
                    Console.WriteLine("One or more files could not be read. Exiting.");
                    return;
                }

                Console.WriteLine($"Loaded ST: {stTable.Rows.Count} rows");
                Console.WriteLine($"Loaded PT: {ptTable.Rows.Count} rows");
                Console.WriteLine($"Loaded RT: {rtTable.Rows.Count} rows");

                // 3. Create HashSets for fast lookup
                HashSet<string> ptHashes = GenerateRowHashes(ptTable);
                HashSet<string> rtHashes = GenerateRowHashes(rtTable);

                int foundInPT = 0;
                int foundInRT = 0;
                int missingEverywhere = 0;

                Console.WriteLine("\n--- Comparison Results ---\n");

                // 4. Compare ST against PT and RT
                foreach (DataRow row in stTable.Rows)
                {
                    string rowSignature = GetRowSignature(row);

                    if (ptHashes.Contains(rowSignature))
                    {
                        foundInPT++;
                    }
                    else
                    {
                        if (rtHashes.Contains(rowSignature))
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"[FOUND IN RT] Row Data: {rowSignature}");
                            foundInRT++;
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"[MISSING] Row Data: {rowSignature}");
                            missingEverywhere++;
                        }
                    }
                }

                Console.ResetColor();
                Console.WriteLine("\n--- Summary ---");
                Console.WriteLine($"Total Rows in ST: {stTable.Rows.Count}");
                Console.WriteLine($"Matched in PT (Primary): {foundInPT}");
                Console.WriteLine($"Matched in RT (Reference): {foundInRT}");
                Console.WriteLine($"Missing Completely: {missingEverywhere}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }

            
        }

        /// <summary>
        /// Reads an Excel file into a DataTable using DocumentFormat.OpenXml.
        /// Handles SharedStrings and Sparse Columns.
        /// </summary>
        static DataTable ReadExcelOpenXml(string filePath)
        {
            DataTable dt = new DataTable();

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return null;
            }

            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    SharedStringTable sst = sstPart?.SharedStringTable;

                    // Get the first sheet
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                    if (sheet == null) return null;

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Read rows
                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count == 0) return dt;

                    // 1. Determine columns based on the header row (first row)
                    Row headerRow = rows.First();
                    int maxColIndex = 0;
                    
                    // Dictionary to map Excel Column Name (A, B, C) to Index (0, 1, 2)
                    foreach (Cell cell in headerRow.Elements<Cell>())
                    {
                        int colIndex = GetColumnIndex(cell.CellReference);
                        string colName = GetCellValue(doc, cell);
                        
                        // Ensure DataTable has enough columns (handle sparse headers if necessary)
                        while (dt.Columns.Count <= colIndex)
                        {
                            dt.Columns.Add();
                        }
                        
                        // Set Column Name (optional, helps with debugging)
                        if (!string.IsNullOrEmpty(colName))
                            dt.Columns[colIndex].ColumnName = colName;
                        
                        if (colIndex > maxColIndex) maxColIndex = colIndex;
                    }

                    // 2. Read Data Rows (Skip header)
                    foreach (Row row in rows.Skip(1))
                    {
                        DataRow dataRow = dt.NewRow();
                        bool emptyRow = true;

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            int colIndex = GetColumnIndex(cell.CellReference);

                            // Expand row if this row has more columns than header (unlikely but possible)
                            while (dt.Columns.Count <= colIndex) dt.Columns.Add();

                            string val = GetCellValue(doc, cell);
                            dataRow[colIndex] = val;

                            if (!string.IsNullOrWhiteSpace(val)) emptyRow = false;
                        }

                        if (!emptyRow)
                            dt.Rows.Add(dataRow);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading {filePath}: {ex.Message}");
                return null;
            }

            return dt;
        }

        /// <summary>
        /// Helper to get cell value. Resolves Shared Strings if necessary.
        /// </summary>
        static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.CellValue == null) return string.Empty;

            string value = cell.CellValue.InnerText;

            // If the cell represents a SharedString, look it up
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart sstPart = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sstPart != null && sstPart.SharedStringTable != null)
                {
                    return sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            
            return value;
        }

        /// <summary>
        /// Converts Excel cell reference (e.g., "A1", "C4") to zero-based column index.
        /// </summary>
        static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference)) return 0;

            // Remove digits to get the column letter(s)
            string columnName = Regex.Replace(cellReference, "[0-9]", "");
            int number = 0;
            int pow = 1;
            
            for (int i = columnName.Length - 1; i >= 0; i--)
            {
                number += (columnName[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            
            return number - 1;
        }

        static HashSet<string> GenerateRowHashes(DataTable table)
        {
            HashSet<string> hashes = new HashSet<string>();
            foreach (DataRow row in table.Rows)
            {
                hashes.Add(GetRowSignature(row));
            }
            return hashes;
        }

        static string GetRowSignature(DataRow row)
        {
            StringBuilder sb = new StringBuilder();
            // Loop through all columns in the DataTable
            for (int i = 0; i < row.Table.Columns.Count; i++)
            {
                string val = row[i]?.ToString().Trim() ?? "";
                sb.Append(val);
                if (i < row.Table.Columns.Count - 1) sb.Append("|");
            }
            return sb.ToString();
        }
    }
}