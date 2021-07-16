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
            if (args.Length < 1 || args.Length > 3)
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
"   List Worksheets: SharpExcelibur.exe sheets C:\\Some\\Path\\To\\ExcelWorkbook\n" +
"   Read Sheet Data: SharpExcelibur.exe read <sheetname> C:\\Some\\Path\\To\\ExcelWorkbook\n\n" +
"");
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

                            xlWorkbook.Close(true, null, null);
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
                else
                {
                        Console.WriteLine("Error...Wrong arguments passed...See help menu");
                        Console.WriteLine($"Provided arguments: SharpExcelibur.exe" + " " + args[0] + " " + args[1]);
                        System.Environment.Exit(0);

                }

            }
            else if (args.Length == 3) 
            {
                if (args[0].ToLower() == "read")
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

                            for (int i = 1; i <= rowCount; i++)
                            {
                                for (int j = 1; j <= colCount; j++)
                                {
                                    //new line
                                    if (j == 1)
                                        Console.Write("\r\n");

                                    //write the value to the console
                                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Text != null)
                                    Console.Write(xlRange.Cells[i, j].Text.ToString() + "\t");
                                }
                                
                            }
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
            else
            {
                Console.WriteLine("Error...Wrong arguments passed...See help menu");
                Console.WriteLine($"Provided arguments: SharpExcelibur.exe" + " " + args[0] + " " + args[1] + " " + args[2]);
                System.Environment.Exit(0);

            }

        }
    }
}


