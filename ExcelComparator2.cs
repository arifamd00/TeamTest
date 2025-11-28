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
    class ExcelComparator2
    {
        static void Main(string[] args)
        {
            // File paths
            string stPath = @"C:\Users\Tyagi\vsp\testProject1\reports\ST.xlsx"; // ST
            string ptPath = @"C:\Users\Tyagi\vsp\testProject1\reports\PT.xlsx"; // PT

            Console.WriteLine("Starting ST ↔ PT Excel Comparison...");

            try
            {
                // Read Excel files
                DataTable stTable = ReadExcelOpenXml(stPath);
                DataTable ptTable = ReadExcelOpenXml(ptPath);

                if (stTable == null || ptTable == null)
                {
                    Console.WriteLine("Error: Unable to load one or both Excel files.");
                    return;
                }

                Console.WriteLine($"Loaded ST: {stTable.Rows.Count} rows");
                Console.WriteLine($"Loaded PT: {ptTable.Rows.Count} rows");

                // HashSets for fast looku
