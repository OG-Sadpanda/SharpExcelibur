using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
// Reference : https://coderwall.com/p/app3ya/read-excel-file-in-c
// Reference : https://www.c-sharpcorner.com/forums/how-to-get-excel-sheet-name-in-c-sharp

namespace SharpExcelibur
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1 || args.Length > 4)
            {
                Console.WriteLine(@"



 _____ _                      _____             _ _ _                
/  ___| |                    |  ___|           | (_) |               
\ `--.| |__   __ _ _ __ _ __ | |____  _____ ___| |_| |__  _   _ _ __ 
 `--. \ '_ \ / _` | '__| '_ \|  __\ \/ / __/ _ \ | | '_ \| | | | '__|
/\__/ / | | | (_| | |  | |_) | |___>  < (_|  __/ | | |_) | |_| | |   
\____/|_| |_|\__,_|_|  | .__/\____/_/\_\___\___|_|_|_.__/ \__,_|_|   
                       | |                                           
                       |_|                                           
" +
"" +
"Developed By: @sadpanda_sec \n\n" +
"Description: Read Contents of Excel Documents (XLS/XLSX).\n\n" +
"Usage:\n" +
"Check if PW Protected:         SharpExcelibur.exe check C:\\Some\\Path\\To\\ExcelWorkbook\n" +
"List Worksheets:               SharpExcelibur.exe sheets C:\\Some\\Path\\To\\ExcelWorkbook\n" +
"List Protected Worksheets:     SharpExcelibur.exe sheets P@ssw0rd C:\\Some\\Path\\To\\ExcelWorkbook\n" +
"Read Sheet Data:               SharpExcelibur.exe read <sheetname> C:\\Some\\Path\\To\\ExcelWorkbook\n" +
"Read Protected Sheet Data:     SharpExcelibur.exe read <sheetname> P@ssw0rd C:\\Some\\Path\\To\\ExcelWorkbook\n");
            System.Environment.Exit(0);

            }
            else if (args.Length == 2)
            {
                if (args[0].ToLower() == "sheets")
                {
                    if (File.Exists(Path.GetFullPath(args[1])))
                    {
                        if (Path.GetExtension(args[1]).ToLower() == ".xls" || Path.GetExtension(args[1]).ToLower() == ".xlsx")
                        {
                      
                            Excel.Application xlApp = new Excel.Application();
                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(args[1]),Type.Missing,true);
                            String[] sheets = new string[xlWorkbook.Worksheets.Count];
                            int i = 0;
                            foreach (Excel.Worksheet wSheet in xlWorkbook.Worksheets)
                            {

                                Console.WriteLine("Sheet Name Found: " + (sheets[i] = wSheet.Name));
                                i++;                               
                            }

                            xlWorkbook.Close(false, null, null);
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlWorkbook);
                            Marshal.ReleaseComObject(xlApp);
                            System.Environment.Exit(0);
                        }
                        else
                        {
                            Console.WriteLine("Error...Did not provide an Excel Worksheet");
                            System.Environment.Exit(0);
                        }

                    }
                  
                }
                else if (args[0].ToLower() == "check")
                {
                    if (File.Exists(Path.GetFullPath(args[1])))
                    {
                        if (Path.GetExtension(args[1]).ToLower() == ".xls" || Path.GetExtension(args[1]).ToLower() == ".xlsx")
                        {
                            try
                            {
                                Excel.Application xlApp = new Excel.Application();
                                xlApp.Visible = false;
                                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(args[1]),Type.Missing,true,Type.Missing," ");
                                xlWorkbook.Close(false, null, null);
                                xlApp.Quit();
                                Marshal.ReleaseComObject(xlWorkbook);
                                Marshal.ReleaseComObject(xlApp);
                                Console.WriteLine("Excel Document NOT Password Protected");
                                System.Environment.Exit(0);
                            }
                            catch (COMException ex)
                            {
                                Console.WriteLine("Excel Document is Password Protected");
                                System.Environment.Exit(0);
                            }
                            
                        }
                        else
                        {
                            Console.WriteLine("Error...Did not provide an Excel Worksheet");
                            System.Environment.Exit(0);
                        }

                    }

                }
                else
                {
                        Console.WriteLine("Error...Wrong arguments passed...See help menu");
                        Console.WriteLine($"Provided arguments: SharpExcelibur.exe" + " " + args[0] + " " + args[1]);
                        System.Environment.Exit(0);
                }

            }

            else if (args.Length == 3) 
            {
                if (args[0].ToLower() == "sheets")
                {
                    if (File.Exists(Path.GetFullPath(args[2])))
                    {
                        if (Path.GetExtension(args[2]).ToLower() == ".xls" || Path.GetExtension(args[2]).ToLower() == ".xlsx")
                        {
                            try
                            {
                                string pass = args[1].ToString();
                                Excel.Application xlApp = new Excel.Application();
                                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(args[2]), Type.Missing, true, Type.Missing, pass, pass);
                                xlWorkbook.Password = pass;
                                xlWorkbook.Unprotect(pass);
                                String[] sheets = new string[xlWorkbook.Worksheets.Count];

                                int i = 0;
                                foreach (Excel.Worksheet wSheet in xlWorkbook.Worksheets)
                                {
                                    wSheet.Unprotect(pass);
                                    Console.WriteLine("Sheet Name Found: " + (sheets[i] = wSheet.Name));
                                    i++;
                                }

                                xlWorkbook.Close(false, null, null);
                                xlApp.Quit();
                                Marshal.ReleaseComObject(xlWorkbook);
                                Marshal.ReleaseComObject(xlApp);
                                System.Environment.Exit(0);
                            }
                            catch (COMException)
                            {
                                Console.WriteLine("Error: Supplied Incorrect Password");
                                System.Environment.Exit(0);
                            }

                        }
                        else
                        {
                            Console.WriteLine("Error...Did not provide an Excel Worksheet");
                            System.Environment.Exit(0);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Error...Did not provide an Excel Worksheet");
                        System.Environment.Exit(0);
                    }
                }
                else if (args[0].ToLower() == "read")
                {
                    if (File.Exists(Path.GetFullPath(args[2])))
                    {
                        if (Path.GetExtension(args[2]).ToLower() == ".xls" || Path.GetExtension(args[2]).ToLower() == ".xlsx")
                        {
                            Excel.Application xlApp = new Excel.Application();
                            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(args[2]),Type.Missing,true);
                            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[args[1]];
                            Excel.Range xlRange = xlWorksheet.UsedRange;
                            int rowCount = xlRange.Rows.Count;
                            int colCount = xlRange.Columns.Count;

                            PrintExcelGrid(xlRange, rowCount, colCount);

                            xlWorkbook.Close(false, null, null);
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlWorksheet);
                            Marshal.ReleaseComObject(xlWorkbook);
                            Marshal.ReleaseComObject(xlApp);
                            System.Environment.Exit(0);
                        }

                        else
                        {
                            Console.WriteLine("Error...Did not provide path to Excel Worksheet");
                            System.Environment.Exit(0);
                        }

                    }
                    else
                    {
                        Console.WriteLine("Error...Excel file doesnt exists");
                        System.Environment.Exit(0);
                    }
                }
                else
                {
                    Console.WriteLine("Error...Wrong arguments passed...See help menu");
                    Console.WriteLine($"Provided arguments: SharpExcelibur.exe" + " " + args[0] + " " + args[1] + " " + args[2]);
                    System.Environment.Exit(0);

                }
            }
            else if (args.Length == 4)
            {
                if (args[0].ToLower() == "read")
                {
                    if (File.Exists(Path.GetFullPath(args[3])))
                    {
                        if (Path.GetExtension(args[3]).ToLower() == ".xls" || Path.GetExtension(args[3]).ToLower() == ".xlsx")
                        {
                            try
                            {
                                string pass = args[2].ToString();
                                Excel.Application xlApp = new Excel.Application();
                                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(args[3]), Type.Missing, true, Type.Missing, pass, pass);
                                xlWorkbook.Password = pass;
                                xlWorkbook.Unprotect(pass);
                                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[args[1]];
                                xlWorksheet.Unprotect(pass);
                                Excel.Range xlRange = xlWorksheet.UsedRange;
                                int rowCount = xlRange.Rows.Count;
                                int colCount = xlRange.Columns.Count;

                                PrintExcelGrid(xlRange, rowCount, colCount);

                                xlWorkbook.Close(false, null, null);
                                xlApp.Quit();
                                Marshal.ReleaseComObject(xlWorksheet);
                                Marshal.ReleaseComObject(xlWorkbook);
                                Marshal.ReleaseComObject(xlApp);
                                System.Environment.Exit(0);
                            }
                            catch (COMException)
                            {
                                Console.WriteLine("Error: Supplied Incorrect Password");
                                System.Environment.Exit(0);
                            }
                            
                        }
                        else
                        {
                            Console.WriteLine("Error...Did not provide path to Excel Worksheet");
                            System.Environment.Exit(0);
                        }

                    }
                    else
                    {
                        Console.WriteLine("Error...Excel file doesnt exists");
                        System.Environment.Exit(0);
                    }
                }
                else
                {
                    Console.WriteLine("Error...Wrong arguments passed...See help menu");
                    Console.WriteLine($"Provided arguments: SharpExcelibur.exe" + " " + args[0] + " " + args[1] + " " + args[2]);
                    System.Environment.Exit(0);

                }
            }
            else
            {
                Console.WriteLine("Error...Wrong arguments passed...See help menu");
                Console.WriteLine($"Provided arguments: SharpExcelibur.exe" + " " + args[0] + " " + args[1] + " " + args[2] + " " + args[3]);
                System.Environment.Exit(0);

            }

        }

        static void PrintExcelGrid(Excel.Range xlRange, int rowCount, int colCount)
        {
            int[] maxLengths = new int[colCount];
            string[,] cellContents = new string[rowCount, colCount];

            // First pass: store contents and find max lengths
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    string content = xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Text != null
                        ? xlRange.Cells[i, j].Text.ToString()
                        : "";
                    cellContents[i - 1, j - 1] = content;
                    maxLengths[j - 1] = Math.Max(maxLengths[j - 1], content.Length);
                }
            }

            // Print top border
            PrintHorizontalBorder(maxLengths);

            // Print rows
            for (int i = 0; i < rowCount; i++)
            {
                Console.Write("|");
                for (int j = 0; j < colCount; j++)
                {
                    Console.Write(" " + cellContents[i, j].PadRight(maxLengths[j]) + " |");
                }
                Console.WriteLine();

                // Print horizontal border between rows
                PrintHorizontalBorder(maxLengths);
            }
        }

        static void PrintHorizontalBorder(int[] maxLengths)
        {
            Console.Write("+");
            for (int j = 0; j < maxLengths.Length; j++)
            {
                Console.Write(new string('-', maxLengths[j] + 2));
                Console.Write(j < maxLengths.Length - 1 ? "+" : "+");
            }
            Console.WriteLine();
        }
    }
}
