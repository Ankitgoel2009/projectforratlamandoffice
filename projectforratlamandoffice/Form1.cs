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
        void OpenKeywordsFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            OpenFileDialog fileDialog = sender as OpenFileDialog;
            selectedFile = System.IO.Path.GetFileNameWithoutExtension(fileDialog.FileName);
            if (string.IsNullOrEmpty(selectedFile) || selectedFile.Contains(".lnk"))
            {
                MessageBox.Show("Please select a valid Excel File");
                e.Cancel = true;
            }
            return;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            choofdlog.Multiselect = false;
            choofdlog.ValidateNames = true;
            choofdlog.DereferenceLinks = false; // Will return .lnk in shortcuts.
            choofdlog.Filter = "Excel |*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.FileOk += new System.ComponentModel.CancelEventHandler(OpenKeywordsFileDialog_FileOk);

            DialogResult result = choofdlog.ShowDialog();
            if (result == DialogResult.OK)
            {
                selectedFile = choofdlog.FileName;

                // Set cursor as hourglass
                Cursor.Current = Cursors.WaitCursor;
                dowork();
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
            }
            else if (result == DialogResult.Cancel)
            {
                // Handle the case when the user clicks the Cancel button
                MessageBox.Show("You have cancelled the file selection.");
            }

        }

        public static Excel.Workbook Open(Excel.Application excelInstance, string fileName, bool readOnly = false, bool editable = true, bool updateLinks = true)
        {
            Excel.Workbook book = excelInstance.Workbooks.Open(
                fileName, updateLinks, readOnly,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, editable, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
            return book;
        }


        Excel.Application excel = null;
        Excel.Workbook wkb = null;
        Excel._Worksheet sheet = null;
        Excel.Range range1 = null;
        int rowindex, columnindex;     
        void dowork()
        {
            excel = new Excel.Application();
            excel.Visible = true;
            // return instance of excel file by sending its path 
            wkb = Open(excel, selectedFile);
            sheet = wkb.Sheets[1];
            sheet.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            sheet.Range["C1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            range1 = sheet.UsedRange; // returns range till grandtotal 
            rowindex = range1.Rows.Count;
            columnindex = range1.Columns.Count;

            Thread t1 = new Thread(Method1);
            Thread t2 = new Thread(Method2);
            Thread t3 = new Thread(method3);
            t1.Start();
            t2.Start();
            t3.Start();


            // wait for both threads to finish
            t1.Join();
             t2.Join();
             t3.Join();

            //// Set up Chrome driver
            //// find correct version of driver at https://sites.google.com/chromium.org/driver/downloads?authuser=0
            ///// add chromedriver to environment variables 
            ///chromedriver -version  // check before 
            /////cd "C:\Program Files\Google\Chrome\Application"
            /// chrome.exe --remote-debugging-port=9222 --user-data-dir=D:\chromedata
            ///
            // Process proc = new Process();
            // before copying below please note c:\ instaedof c 
            //proc.StartInfo.FileName = @"C\Program Files\Google\Chrome\Application\chrome.exe";
            // to avoid typing above query add path of chrome to environmane varibale 
            //proc.StartInfo.Arguments = "--remote-debugging-port=9222 --user-data-dir=D:\\chromedata";
            // proc.Start();

            //   var options = new ChromeOptions();
            //   options.DebuggerAddress= "127.0.0.1:9222";
            //options.AddArguments("--profile-directory=Profile 1");
            //options.AddArguments("--disable-extensions");
            //options.AddArguments("--headless");
            //options.AddArguments("--no-sandbox", "--disable-dev-shm-usage");
            //   IWebDriver driver = new ChromeDriver(@"C:\Users\Umesh Aggarwal\Desktop\chromedriver_win32", options);

            // driver.Manage().Window.Maximize();
            ////////// Navigate to Whatsapp web
            //   driver.Navigate().GoToUrl("https://web.whatsapp.com/");
            //IReadOnlyCollection<string> windowHandles = driver.WindowHandles;

            //// Find already opened window with Chrome
            //string chromeWindow = "";
            //foreach (string window in windowHandles)
            //{
            //    driver.SwitchTo().Window(window);
            //    if (driver.Title.Contains("Google Chrome"))
            //    {
            //        chromeWindow = window;
            //        break;
            //    }
            //}

            //// Switch to Chrome window
            //driver.SwitchTo().Window(chromeWindow);

            //// Get all open tabs in Chrome window
            //IReadOnlyCollection<string> tabHandles = driver.WindowHandles;

            //// Find already opened tab of Whatsapp web
            //string whatsappTab = "";
            //foreach (string tab in tabHandles)
            //{
            //    driver.SwitchTo().Window(tab);
            //    if (driver.Title.Contains("Whatsapp"))
            //    {
            //        whatsappTab = tab;
            //        break;
            //    }
            //}

            ////// Switch to Whatsapp tab
            //driver.SwitchTo().Window(whatsappTab);
            //// IWebElement whatsappTab = driver.FindElement(By.XPath("//title[contains(text(), 'whatsapp')]"));
            //// Thread.Sleep(5000);
            ////  whatsappTab.Click();
            ////   Get list of contacts to send message to
            //string[] contacts = { "Umesh Ji" };

            ////// Loop through each contact
            //foreach (string contact in contacts)
            //{
            //    //     Find contact in chat list
            //    IWebElement contactElement = driver.FindElement(By.XPath($"//span[contains(text(), '{contact}')]"));
            //    contactElement.Click();

            //    // Click on attachment icon
            //    IWebElement attachmentIcon = driver.FindElement(By.XPath("//div[@title='Attach']"));
            //    attachmentIcon.Click();

            //    // Select file to attach
            //    IWebElement fileInput = driver.FindElement(By.XPath("//input[@accept='*']"));
            //    fileInput.SendKeys(@"C:\Users\Public\ratlam.xlsx");

            //    // Wait for file to upload
            //    Thread.Sleep(5000);

            //    // Click on send button
            //    IWebElement sendButton = driver.FindElement(By.XPath("//span[@data-icon='send']"));
            //    sendButton.Click();
            //    // }

            //    //  Close browser
            //    driver.Quit();
            wkb.Close(true);
            excel.Quit();
            // CLEAN UP.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            System.Windows.Forms.Application.Exit();
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
            // prepare all customers file 
            string outputpathforoffice = @"C:\ratlamfile\office-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
           // Excel.Application excelforofice = new Excel.Application();
            Excel.Workbook workbookforoffice = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheetforoffice = (Excel.Worksheet)workbookforoffice.ActiveSheet;
            int row = 1;
            int column = 1;
            for (int i = 10; i <= rowindex - 1; i++)
            {
                string valueA = (string)range1.Cells[i, 1].Value;
                if (range1.Cells[i, 2].Value != null)
                {
                    decimal valueB = (decimal)range1.Cells[i, 2].Value;
                    if (valueB > 1000)
                    {
                        if (column > 4)
                        {
                            column = 1;
                            row += 1;
                        }
                        sheetforoffice.Cells[row, column].Value = valueA;
                        sheetforoffice.Cells[row, ++column].Value = valueB;
                        column++;
                    }
                }
            }
            sheetforoffice.Columns["A:D"].AutoFit();
            sheetforoffice.Columns[1].ColumnWidth = 36;
            sheetforoffice.Columns[2].ColumnWidth = 14.44;
            sheetforoffice.Columns[3].ColumnWidth = 37.44;
            sheetforoffice.Columns[4].ColumnWidth = 14.44;
            // Set the row height for a range of rows
            Excel.Range rowRange = sheetforoffice.Range["59:" + sheetforoffice.Cells[sheetforoffice.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row];
            rowRange.RowHeight = 9.80;
            rowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            sheetforoffice.Range["A:D"].EntireColumn.Font.Bold = true;
            sheetforoffice.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00 ₹;[>=100000] ##\,##\,##0.00 ₹;##,##0.00 ₹";
            sheetforoffice.Range["D1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00 ₹;[>=100000] ##\,##\,##0.00 ₹;##,##0.00 ₹";
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
        public void Method2()
        {
            // listnew will contain all the data from the row 10 upto the last row excluding the final total row 
            List<object[]> listnew = new List<object[]>();
            for (int i = 10; i <= rowindex - 1; i++)
            {
                if (range1.Cells[i, 1].Text != "")
                {
                    if (range1.Cells[i, 2].Text != "" && range1.Cells[i, 3].Text == "")
                    {
                        listnew.Add(new object[] { range1.Cells[i, 1].Text, "+ " + range1.Cells[i, 2].Text + " ₹" });
                    }
                    else if (range1.Cells[i, 2].Text == "" && range1.Cells[i, 3].Text != "")
                    {
                        listnew.Add(new object[] { range1.Cells[i, 1].Text, "- " + range1.Cells[i, 3].Text + " ₹" });
                    }

                }
            }
            // commented out because i started adding from element 10 and also loop stops before length -1 
            // list.RemoveAt(list.Count - 1);
            //  list.RemoveRange(0, 10);         


            // 6. copy text file names to a variable called list 
            copydata();

            // filter from only index which conatins  names and add the balance 
            Filter(listnew, list);

            // names are  ready , now modify it 

            Dictionary<object, object> dic1 = new Dictionary<object, object>();
            for (int i = 0; i <= listnew.Count - 1; i++)
            {
                dic1.Add(listnew[i][0], listnew[i][1]);

            }

            // data which tells which names are in which state
            var arr2 = GetObjArr(@"C:\ratlamfile\statewisenames.xlsx");
            List<object> list1 = new List<object>();
            for (int i = 1; i <= arr2.GetLength(0); i++)
            {
                if (arr2[i, 1] != null && arr2[i, 2] == null)
                {
                    list1.Add(arr2[i, 1]);
                }
                if (arr2[i, 2] != null && arr2[i, 1] == null)
                {
                    list1.Add(arr2[i, 2]);
                }

            }

            string outputpath = @"C:\ratlamfile\Ankit_ji_Ratlam-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
           // Excel.Application excel1 = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet3 = (Excel.Worksheet)workbook.ActiveSheet;

            for (int i = 0; i < list1.Count; i++)
            {
                if (list1[i].ToString().Trim() == "U.P" || list1[i].ToString().Trim() == "Rajasthan" || list1[i].ToString().Trim() == "Bihar" || list1[i].ToString().Trim() == "Punjab" || list1[i].ToString().Trim() == "Odisha" || list1[i].ToString().Trim() == "Chhatisgarh" || list1[i].ToString().Trim() == "West bengal" || list1[i].ToString().Trim() == "Madhya Pradesh" || list1[i].ToString().Trim() == "Jharkhand" || list1[i].ToString().Trim() == "Maharashtra" || list1[i].ToString().Trim() == "Market" || list1[i].ToString().Trim() == "Uttarakhand" || list1[i].ToString().Trim() == "Assam" || list1[i].ToString().Trim() == "Tripura")
                {
                    //sheet.Cells[i + 1, 2].Value = dic1[list[i]];

                    sheet3.Range[sheet3.Cells[i + 1, 1], sheet3.Cells[i + 1, 2]].EntireColumn.Font.Bold = true;
                    sheet3.Range[sheet3.Cells[i + 1, 1], sheet3.Cells[i + 1, 2]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    sheet3.Range[sheet3.Cells[i + 1, 1], sheet3.Cells[i + 1, 2]].Merge();
                    sheet3.Range[sheet3.Cells[i + 1, 1], sheet3.Cells[i + 1, 2]].Cells.Font.Size = 20;
                    sheet3.Range[sheet3.Cells[i + 1, 1], sheet3.Cells[i + 1, 2]].Font.Italic = true;

                }
                else
                {
                    sheet3.Cells[i + 1, 2].Value = dic1[list1[i]];
                }

                sheet3.Cells[i + 1, 1].Value = list1[i];
            }
            //  sheet1.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0;##,##0.00";
            sheet3.Columns["A:B"].AutoFit();
            sheet3.Range["A1"].EntireColumn.Font.Bold = true;
            sheet3.Range["B1"].EntireColumn.Font.Bold = true;

            range1 = sheet3.UsedRange;
            Borders border = range1.Borders;
            border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Color = Color.Black;
            border[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            border[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            border[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            border[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            range1.Borders.Color = Color.Black;
            range1.Select();
            sheet3.UsedRange.Select();
            workbook.SaveAs(outputpath);
            workbook.Close();
          //  excel1.Quit();
            // CLEAN UP.
           // System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet3);
          
        }


       // Excel.Range range3 = null;
        public void method3()
        {
            List<object[]> initialoutputlistforharshit = new List<object[]>();
            for (int i = 10; i <= rowindex - 1; i++)
            {
                if (range1.Cells[i, 1].Text != "")
                {
                    if (range1.Cells[i, 2].Text != "" && range1.Cells[i, 3].Text == "")
                    {
                        initialoutputlistforharshit.Add(new object[] { range1.Cells[i, 1].Text, "+ " + range1.Cells[i, 2].Text + " ₹" });
                    }
                    else if (range1.Cells[i, 2].Text == "" && range1.Cells[i, 3].Text != "")
                    {
                        initialoutputlistforharshit.Add(new object[] { range1.Cells[i, 1].Text, "- " + range1.Cells[i, 3].Text + " ₹" });
                    }

                }
            }

            copydata1();
            filter1(initialoutputlistforharshit);


            Dictionary<object, object> dictionarynamesforharhit = new Dictionary<object, object>();
            for (int i = 0; i <= initialoutputlistforharshit.Count - 1; i++)
            {
                dictionarynamesforharhit.Add(initialoutputlistforharshit[i][0], initialoutputlistforharshit[i][1]);

            }

            // data which tells which names are in which state
            var nameexcelarrayforharshit = GetObjArr(@"C:\ratlamfile\harshitstatewisenames.xlsx");
            List<object> listforharshit = new List<object>();
            for (int i = 1; i <= nameexcelarrayforharshit.GetLength(0); i++)
            {
                if (nameexcelarrayforharshit[i, 1] != null && nameexcelarrayforharshit[i, 2] == null)
                {
                    listforharshit.Add(nameexcelarrayforharshit[i, 1]);
                }
                if (nameexcelarrayforharshit[i, 2] != null && nameexcelarrayforharshit[i, 1] == null)
                {
                    listforharshit.Add(nameexcelarrayforharshit[i, 2]);
                }

            }


            string outputpath = @"C:\ratlamfile\Harshit_ji_Ratlam-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
          //  Excel.Application excel1 = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.ActiveSheet;


            for (int i = 0; i < listforharshit.Count; i++)
            {
                if (listforharshit[i].ToString().Trim() == "U.P" || listforharshit[i].ToString().Trim() == "Rajasthan" || listforharshit[i].ToString().Trim() == "Bihar" || listforharshit[i].ToString().Trim() == "Punjab" || listforharshit[i].ToString().Trim() == "Odisha" || listforharshit[i].ToString().Trim() == "Chhatisgarh" || listforharshit[i].ToString().Trim() == "West bengal" || listforharshit[i].ToString().Trim() == "Madhya Pradesh" || listforharshit[i].ToString().Trim() == "Jharkhand" || listforharshit[i].ToString().Trim() == "Maharashtra" || listforharshit[i].ToString().Trim() == "Market" || listforharshit[i].ToString().Trim() == "Uttarakhand" || listforharshit[i].ToString().Trim() == "Assam" || listforharshit[i].ToString().Trim() == "Tripura")
                {
                    //sheet.Cells[i + 1, 2].Value = dic1[list[i]];

                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].EntireColumn.Font.Bold = true;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Merge();
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Cells.Font.Size = 20;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Font.Italic = true;

                }
                else
                {
                    sheet1.Cells[i + 1, 2].Value = dictionarynamesforharhit[listforharshit[i]];
                }

                sheet1.Cells[i + 1, 1].Value = listforharshit[i];
            }
            //  sheet1.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0;##,##0.00";
            sheet1.Columns["A:B"].AutoFit();
            sheet1.Range["A1"].EntireColumn.Font.Bold = true;
            sheet1.Range["B1"].EntireColumn.Font.Bold = true;

            range1 = sheet1.UsedRange;
            Borders border = range1.Borders;
            border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Color = Color.Black;
            border[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            border[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            border[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            border[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            range1.Borders.Color = Color.Black;
            range1.Select();
            sheet1.UsedRange.Select();
            workbook.SaveAs(outputpath);
            workbook.Close();
           //  excel1.Quit();
            // CLEAN UP.
           // System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
          
        }

        List<string> textnamesforharshit = new List<string>();
        public void copydata1()
        {
            string filename = "harshitjinames.txt";
            string filePath = @"C:\ratlamfile\" + filename;
            using (var file = new StreamReader(filePath))
            {
                var line = string.Empty;
                while ((line = file.ReadLine()) != null)
                {
                    textnamesforharshit.Add(Convert.ToString(line, CultureInfo.InvariantCulture));
                }
            }
        }

        public void filter1(List<object[]> ws1)
        {
            var result1 = ws1.Select(m => m[0]).ToList();
            var result = result1.Except(textnamesforharshit);
            ws1.RemoveAll(m => result.Contains(m[0]));
            for (int i = 0; i < textnamesforharshit.Count; i++)
            {
                if (!result1.Contains(textnamesforharshit[i]))
                {
                    ws1.Add(new object[] { textnamesforharshit[i], "0 ₹" });
                }
            }
        }
    }
}
