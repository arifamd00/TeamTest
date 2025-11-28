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
    public class ExcelComparator2
    {
        public static void Main(string[] args)
        {
            string stPath = @"C:\Users\Tyagi\vsp\testProject1\reports\ST.xlsx";
            string ptPath = @"C:\Users\Tyagi\vsp\testProject1\reports\PT.xlsx";

            Console.WriteLine("Starting Excel Comparison (ST <-> PT)…");

            try
            {
                DataTable stTable = ReadExcelOpenXml(stPath);
                DataTable ptTable = ReadExcelOpenXml(ptPath);

                if (stTable == null || ptTable == null)
                {
                    Console.WriteLine("One or more files could not be read. Exiting.");
                    return;
                }

                Console.WriteLine($"Loaded ST: {stTable.Rows.Count} rows");
                Console.WriteLine($"Loaded PT: {ptTable.Rows.Count} rows");

                HashSet<string> stHashes = GenerateRowHashes(stTable);
                HashSet<string> ptHashes = GenerateRowHashes(ptTable);

                DataTable missing = CreateMissingRowsTable(stTable);

                Console.WriteLine("\n--- Comparing ST → PT ---\n");

                // ST → PT (PT missing)
                foreach (DataRow row in stTable.Rows)
                {
                    string sig = GetRowSignature(row);
                    if (!ptHashes.Contains(sig))
                        AddMissingRow(missing, row, "PT");
                }

                Console.WriteLine("\n--- Comparing PT → ST ---\n");

                // PT → ST (ST missing)
                foreach (DataRow row in ptTable.Rows)
                {
                    string sig = GetRowSignature(row);
                    if (!stHashes.Contains(sig))
                        AddMissingRow(missing, row, "ST");
                }

                string outputPath = Path.Combine(
                    Path.GetDirectoryName(stPath),
                    "missingRows.xlsx"
                );

                WriteExcel(outputPath, missing);

                Console.WriteLine($"\nMissing rows written to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        // ==========================================
        //  Excel Reading
        // ==========================================
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
                    SharedStringTable sst =
                        workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable;

                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                    if (sheet == null) return null;

                    WorksheetPart worksheetPart =
                        (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                    SheetData sheetData =
                        worksheetPart.Worksheet.Elements<SheetData>().First();

                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count == 0) return dt;

                    Row headerRow = rows.First();
                    int maxColIndex = 0;

                    foreach (Cell cell in headerRow.Elements<Cell>())
                    {
                        int colIndex = GetColumnIndex(cell.CellReference);
                        string colName = GetCellValue(doc, cell);

                        while (dt.Columns.Count <= colIndex)
                            dt.Columns.Add();

                        if (!string.IsNullOrWhiteSpace(colName))
                            dt.Columns[colIndex].ColumnName = colName;

                        if (colIndex > maxColIndex)
                            maxColIndex = colIndex;
                    }

                    foreach (Row row in rows.Skip(1))
                    {
                        DataRow dr = dt.NewRow();
                        bool emptyRow = true;

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            int colIndex = GetColumnIndex(cell.CellReference);
                            while (dt.Columns.Count <= colIndex) dt.Columns.Add();

                            string val = GetCellValue(doc, cell);
                            dr[colIndex] = val;

                            if (!string.IsNullOrWhiteSpace(val)) emptyRow = false;
                        }

                        if (!emptyRow) dt.Rows.Add(dr);
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

        static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.CellValue == null) return string.Empty;
            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTable sst =
                    doc.WorkbookPart.SharedStringTablePart.SharedStringTable;

                return sst.ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }

        static int GetColumnIndex(string cellReference)
        {
            string col = Regex.Replace(cellReference, "[0-9]", "");
            int result = 0, pow = 1;

            for (int i = col.Length - 1; i >= 0; i--)
            {
                result += (col[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return result - 1;
        }

        // ==========================================
        //  Signature + Mismatch Tracking
        // ==========================================
        static HashSet<string> GenerateRowHashes(DataTable table)
        {
            HashSet<string> set = new HashSet<string>();
            foreach (DataRow row in table.Rows)
                set.Add(GetRowSignature(row));
            return set;
        }

        static string GetRowSignature(DataRow row)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < row.Table.Columns.Count; i++)
            {
                sb.Append((row[i]?.ToString() ?? "").Trim());
                if (i < row.Table.Columns.Count - 1) sb.Append("|");
            }
            return sb.ToString();
        }

        static DataTable CreateMissingRowsTable(DataTable structure)
        {
            DataTable dt = structure.Clone();
            dt.Columns.Add("missingIn");  // PT or ST
            return dt;
        }

        static void AddMissingRow(DataTable dt, DataRow sourceRow, string missingIn)
        {
            DataRow newRow = dt.NewRow();
            for (int i = 0; i < sourceRow.Table.Columns.Count; i++)
                newRow[i] = sourceRow[i];

            newRow["missingIn"] = missingIn;
            dt.Rows.Add(newRow);
        }

        // ==========================================
        //  Write Missing Rows to Excel
        // ==========================================
        // ==========================================
        //  Write Missing Rows to Excel (Improved)
        // ==========================================
        static void WriteExcel(string filePath, DataTable dt)
        {
            using (SpreadsheetDocument doc =
                SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wb = doc.AddWorkbookPart();
                wb.Workbook = new Workbook();

                WorksheetPart ws = wb.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                ws.Worksheet = new Worksheet(sheetData);

                Sheets sheets = wb.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = wb.GetIdOfPart(ws),
                    SheetId = 1,
                    Name = "MissingRows"
                });

                // Header row
                Row headerRow = new Row();
                foreach (DataColumn col in dt.Columns)
                {
                    headerRow.Append(new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(col.ColumnName)
                    });
                }
                sheetData.Append(headerRow);

                // Data rows
                // Data rows (Improved version)
                foreach (DataRow dr in dt.Rows)
                {
                    Row r = new Row();

                    foreach (DataColumn col in dt.Columns)
                    {
                        string raw = dr[col]?.ToString() ?? "";

                        Cell cell = new Cell();

                        double num;
                        DateTime dtValue;

                        // 🔵 Case 1: raw string is an OADate (numeric but actually a date)
                        if (double.TryParse(raw, out num) && num > 20000 && num < 60000)
                        {
                            // Convert OADate number → DateTime string
                            DateTime oa = DateTime.FromOADate(num);

                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(oa.ToString("dd-MM-yyyy HH:mm:ss"));
                        }
                        // 🔵 Case 2: looks like a normal date string
                        else if (DateTime.TryParse(raw, out dtValue))
                        {
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(dtValue.ToString("dd-MM-yyyy HH:mm:ss"));
                        }
                        // 🔵 Case 3: normal numeric value
                        else if (double.TryParse(raw, out num))
                        {
                            cell.DataType = null;
                            cell.CellValue = new CellValue(
                                num.ToString(System.Globalization.CultureInfo.InvariantCulture)
                            );
                        }
                        // 🔵 Case 4: text
                        else
                        {
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(raw);
                        }


                        r.Append(cell);
                    }

                    sheetData.Append(r);
                }

            }
        }

    }
}
