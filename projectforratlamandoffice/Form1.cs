using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Globalization;                                                           
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Diagnostics;
//using Application = Microsoft.Office.Interop.Excel.Application;

namespace projectforratlamandoffice
{
    public partial class Form1 : Form
    {
        string selectedFile;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void ValidateSelectedFile(object sender, CancelEventArgs e)
        {
            var openFile = (OpenFileDialog)sender;
            var fileName = FileUtils.GetFileNameWithoutExtension(openFile.FileName);

            if (!FileUtils.IsValidFile(fileName))
            {
                ShowError("Please select a valid Excel file");
                e.Cancel = true;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            var dialog = new OpenFileDialog()
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Excel Files|*.xlsx",
                ValidateNames = true,
                DereferenceLinks = false
            };

            dialog.FileOk += ValidateSelectedFile;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                selectedFile = dialog.FileName;
                GenerateReport();
            }
            else
            {
                ShowError("File selection was cancelled");
            }
        }

        private void GenerateReport()
        {
            Cursor.Current = Cursors.WaitCursor;
            var workbook = ExcelUtils.OpenWorkbook(selectedFile);
            var sheet = workbook.ActiveSheet;
            ReportUtils.PopulateData(sheet);
            ExcelUtils.SaveAndCloseWorkbook(workbook);
            Cursor.Current = Cursors.Default;
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message);
        }
    
    
       
        private object lockObject = new object();

        public object[,] GetObjArr(string filename)
        {
            lock (lockObject)
            {
                Excel.Application tempexcel = null;
                tempexcel = new Excel.Application();
                tempexcel.Visible = true;
                Excel.Workbook tempwkb = null;
                tempwkb = Open(tempexcel, filename);
                Excel._Worksheet tempsheet = tempwkb.Sheets[1];
                Excel.Range temprange = tempsheet.UsedRange;
                // Read all data from data range in the worksheet
                object[,] valueArray = (object[,])temprange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                tempwkb.Close(true);
                tempexcel.Quit();
                // CLEAN UP.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tempexcel);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tempwkb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tempsheet);
                return valueArray;
            }
        }
     

        // copydata function will drop the data in the abc.txt file into the list variable 

        List<string> list = new List<string>();
        public void copydata()
        {
           string filename = "ankitjinames.txt";
            string filePath = @"C:\ratlamfile\" + filename;
            using (var file = new StreamReader(filePath))
            {
                var line = string.Empty;
                while ((line = file.ReadLine()) != null)
                {
                    list.Add(Convert.ToString(line, CultureInfo.InvariantCulture));
                }
            }
        }

        // filter function will match the data in the listnew variable and the data in the list variable 
        // after that it removes the data from the listnew which are not in list 
        // and at the same time add the results with rs 0 if they are not in ws 
        public void Filter(List<object[]> ws, List<string> values)
        {
            var result1 = ws.Select(m => m[0]).ToList();
            var result = result1.Except(list);
            ws.RemoveAll(m => result.Contains(m[0]));
            for (int i = 0; i < list.Count; i++)
            {
                if (!result1.Contains(list[i]))
                {
                    ws.Add(new object[] { list[i], "0 ₹" });
                }
            }
        }

        public void Method1()
        {
            // sort the source range 
            int lastRow = range1.Row + range1.Rows.Count - 1;
            int lastColumn = range1.Column + range1.Columns.Count - 1;
            // Sort the first column in ascending order, starting from row 3
            Excel.Range sortRange = sheet.Range[sheet.Cells[3, 1], sheet.Cells[lastRow - 1, lastColumn]];

           // prepare all customers file 
            string outputpathforoffice = @"C:\ratlamfile\officeNew-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
           // Excel.Application excelforofice = new Excel.Application();
            Excel.Workbook workbookforoffice = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheetforoffice = (Excel.Worksheet)workbookforoffice.ActiveSheet;

            sheetforoffice.Cells[1, 1].Value = "Name";
            sheetforoffice.Cells[1, 2].Value = "C. Balance";
            sheetforoffice.Cells[1, 3].Value = "L. Payment";
            sheetforoffice.Cells[1, 4].Value = "L. Payment Date";
            sheetforoffice.Cells[1, 5].Value = "Name";
            sheetforoffice.Cells[1, 6].Value = "C. Balance";
            sheetforoffice.Cells[1, 7].Value = "L. Payment";
            sheetforoffice.Cells[1, 8].Value = "L. Payment Date";

            int row =2 ;
            int column = 1;
            for (int i = 3; i <= rowindex - 1; i++)
            {
                string valueA = (string)range1.Cells[i, 1].Value;
                if (range1.Cells[i, 11].Value != null && range1.Cells[i, 11].Value > 1000 && !valueA.ToLower().Contains("udaan")/*&&!valueA.ToLower().Contains("shop")*/)
                {
                    decimal valueB; // = (decimal)range1.Cells[i, 11].Value; // take the value in second column and convert to decimal 
                    string cellValue = range1.Cells[i, 11].NumberFormat.ToString();
                    cellValue = cellValue.Replace("\"", "");
                    // Check if the cell value contains the string "cr"
                    if (cellValue.Contains("Cr"))
                    {
                        // Remove the "cr" from the cell value and convert it to a decimal
                         valueB = (decimal)range1.Cells[i, 11].Value * -1;

                        // Set the cell value to the new value
                       // range1.Cells[i, 11].Value = valueB;
                    }
                    else
                    {
                        // Convert the cell value to a decimal
                        valueB = (decimal)range1.Cells[i, 11].Value;

                        // Set the cell value to the new value
                      //  range1.Cells[i, 11].Value = valueB;
                    }

                    sheetforoffice.Cells[row, column].Value = valueA;
                    sheetforoffice.Cells[row, ++column].Value = valueB;

                    decimal valueC;
                    if (range1.Cells[i, 10].Value == null)
                    {
                        valueC = 0; // set a default value if needed
                        range1.Cells[i, 10].Value = "Data UnFound";
                        sheetforoffice.Cells[row, ++column].Value = range1.Cells[i, 10].Value;
                    }
                    else
                    {
                        sheetforoffice.Cells[row, ++column].Value = valueC = (decimal)range1.Cells[i, 10].Value;
                    }
                    DateTime? valueD = range1.Cells[i, 8].Value as DateTime?;
                    if (valueD == null)
                    {
                        range1.Cells[i, 8].Value = "Data UnFound";
                        sheetforoffice.Cells[row, ++column].Value = range1.Cells[i, 8].Value;
                    }
                    else
                    {
                        sheetforoffice.Cells[row, ++column].Value = valueD;
                    }

                    //if (valueB > 1000) // if the value in second column ie current balance is greater than 1000
                    //{
                    column++;
                    if (column > 8) //if the column number exceeds 6 then 
                        {
                            column = 1;  // revert back to column 1 
                            row += 1;   // and change row 
                        }                     
                    
                      
                        
                   // }
                }
            }

            string fileContents = File.ReadAllText(@"C:\ratlamfile\newnames.txt");

            // Split the file contents by colon and newline characters
            string[] lines = fileContents.Split(new char[] { ':', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Create a dictionary to store the name mappings
            Dictionary<string, string> nameMappings = new Dictionary<string, string>();

            // Add the name mappings to the dictionary
            for (int i = 0; i < lines.Length; i += 2)
            {
                string name = lines[i].Trim();
                string replacement = lines[i + 1].Trim();
                nameMappings[name] = replacement;
            }

            // Get the range of cells in column A starting from row 4
          //  Excel.Range range = sheetforoffice.Range["A1", sheetforoffice.Cells[sheetforoffice.UsedRange.Rows.Count, "A"]];
            Excel.Range range = sheetforoffice.UsedRange;
            // Get the values in the range as a 2D array
            object[,] values = range.Value;

            // Replace the names in the array using the nameMappings dictionary
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (values[i, 1] != null)
                {
                    string name = values[i, 1].ToString().Trim();
                    if (nameMappings.ContainsKey(name))
                    {
                        values[i, 1] = nameMappings[name];
                    }
                }
                if (values[i, 5] != null)
                {
                    string name = values[i, 5].ToString().Trim();
                    if (nameMappings.ContainsKey(name))
                    {
                        values[i, 5] = nameMappings[name];
                    }
                }
            }

            // Set the values in the range to the modified array
            range.Value = values;
           

            sheetforoffice.Range["A:H"].EntireColumn.Font.Bold = true;
            sheetforoffice.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00 ₹;[>=100000] ##\,##\,##0.00 ₹;##,##0.00 ₹";
            sheetforoffice.Range["C1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00 ₹;[>=100000] ##\,##\,##0.00 ₹;##,##0.00 ₹";
            sheetforoffice.Range["F1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00 ₹;[>=100000] ##\,##\,##0.00 ₹;##,##0.00 ₹";
            sheetforoffice.Range["G1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00 ₹;[>=100000] ##\,##\,##0.00 ₹;##,##0.00 ₹";
            Range rangeforoffice = sheetforoffice.UsedRange;
            Borders borderforoffice = rangeforoffice.Borders;
            borderforoffice[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borderforoffice[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borderforoffice[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            borderforoffice[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borderforoffice.Color = Color.Black;
            borderforoffice[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borderforoffice[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borderforoffice[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borderforoffice[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            rangeforoffice.Borders.Color = Color.Black;
            sheetforoffice.Columns["A:H"].AutoFit(); // after bold , now autofit , not before making bold  
            rangeforoffice.Select();
            sheetforoffice.UsedRange.Select();
            workbookforoffice.SaveAs(outputpathforoffice);
            workbookforoffice.Close();
           // excelforofice.Quit();
            // CLEAN UP.
          //  System.Runtime.InteropServices.Marshal.ReleaseComObject(excelforofice);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbookforoffice);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheetforoffice);

        }
        
    }

    // open and close excel files 
    public static class ExcelUtils
    {

        public static Workbook OpenWorkbook(string filePath)
        {
            var excelApp = new Excel.Application();
            return excelApp.Workbooks.Open(filePath);
        }

        public static void SaveAndCloseWorkbook(Workbook workbook)
        {
            workbook.Save();
            workbook.Close();
        }

    }

    // create reports 
    public static class ReportUtils
    {
        static Excel.Range range1 = null;
        static int rowindex;
        static int columnindex;

        public static void PopulateData(Worksheet sheet)
        {
            excel = new Excel.Application();
            excel.Visible = true;         
     
            range1 = sheet.UsedRange; // returns range till grandtotal 
            rowindex = range1.Rows.Count;
            columnindex = range1.Columns.Count;
            Method1();
                    
            wkb.Save();
            wkb.Close(true);
            excel.Quit();
            // CLEAN UP.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            System.Windows.Forms.Application.Exit();
        }

    }

    // return file 
    public static class FileUtils
    {

        public static string GetFileNameWithoutExtension(string path)
        {
            return Path.GetFileNameWithoutExtension(path);
        }

        public static bool IsValidFile(string fileName)
        {
            return !string.IsNullOrEmpty(fileName) &&
                   !fileName.Contains(".lnk");
        }

    }




}
