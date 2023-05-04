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
        // code when the user presses Ok  
        // step 2 
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

        // step 1 
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

        // step 4 
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
        // step 3 
        void dowork()
        {
            excel = new Excel.Application();
            excel.Visible = true;
            // return instance of excel file by sending its path 
            wkb = Open(excel, selectedFile);
            sheet = wkb.Sheets[1];
           // sheet.Range["H1"].EntireColumn.NumberFormat = "DD/MM/YYYY";
            range1 = sheet.UsedRange; // returns range till grandtotal 
            rowindex = range1.Rows.Count;
            columnindex = range1.Columns.Count;
            Method1();
            Method2();
            //  Thread t1 = new Thread(Method1);
            // Thread t2 = new Thread(Method2);
            //  Thread t3 = new Thread(method3);
            //   t1.Start();
            //   t2.Start();
            //   t3.Start();


            // wait for both threads to finish
            //   t1.Join();
            //  t2.Join();
            //  t3.Join();

         
            wkb.Save();
            wkb.Close(true);
            excel.Quit();
            // CLEAN UP.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            System.Windows.Forms.Application.Exit();
        }

        private object lockObject = new object();

        // opens an excel and returns an objevt array 
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


        // copydata function will drop the data in the ankitjinames.txt file into the list variable 

        List<string> list = new List<string>();
        public void copydata(string nameoffile)
        {
           string filename = nameoffile;
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
        public void Filter(Dictionary<string, List<object>> copydata, List<string> listofnames)
        {
            var keysToRemove = copydata.Keys.Except(listofnames).ToList();
            foreach (var key in keysToRemove)
            {
                copydata.Remove(key);
            }

            var keysToAdd = listofnames.Except(copydata.Keys);
            foreach (var key in keysToAdd)
            {
                copydata.Add(key, new List<object> { "Try Again", "Try Again", "Try Again", "Try Again", "Try Again", "Try Again", "0 ₹" });
            }
        }

        public void Method1()
        {

            // Create a source dictionary to store data
            Dictionary<string, List<object>> data = new Dictionary<string, List<object>>();

            // Loop through each cell in the used range starting from row 3
            for (int startrow = 3; startrow <= rowindex - 1; startrow++)
            {
                decimal valueofcolumn11;
                string cellValue = range1.Cells[startrow, 11].NumberFormat.ToString();
                cellValue = cellValue.Replace("\"", "");
                //// Check if the cell value contains the string "cr"
                if (cellValue.Contains("Cr"))
                {
                    // Remove the "cr" from the cell value and convert it to a decimal
                    valueofcolumn11 = (decimal)range1.Cells[startrow, 11].Value * -1;

                }
                else
                {
                    // Convert the cell value to a decimal
                    valueofcolumn11 = (decimal)range1.Cells[startrow, 11].Value;
                }

                // Get key value from first column
                string key = range1.Cells[startrow, 1].Value.ToString();

                // Loop through remaining columns and add values to dictionary
                List<object> valuesinlist = new List<object>();
                for (int startcol = 3; startcol <= columnindex; startcol++)
                {
                    if (startcol == 3 || startcol == 6 || startcol == 7 || startcol == 11)
                    {
                        string value = range1.Cells[startrow, startcol].Value?.ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            if (startcol == 11)
                            {
                                value = string.Format("{0:N} ₹", valueofcolumn11);
                            }
                            else
                            {
                                value = string.Format("{0:N} ₹", decimal.Parse(value));
                            }
                        }
                        else
                        {
                            value = "not found in current month";
                        }
                        valuesinlist.Add(value);
                    }
                    else if (startcol != 5 && startcol != 9) // read every column except column 5 and column 9
                    {
                        object value = range1.Cells[startrow, startcol].Value;
                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                        {
                            valuesinlist.Add("not found in current month");
                        }
                        else
                        {
                            valuesinlist.Add(value.ToString());
                        }
                    }
                }
                // Add the value of column 11 to the valuesinlist
                valuesinlist.Add(string.Format("{0:N} ₹", valueofcolumn11));
                data.Add(key, valuesinlist);
            }

        
            // Sort the dictionary by key in alphabetical order
            data = data.OrderBy(d => d.Key).ToDictionary(d => d.Key, d => d.Value);

           


            // copy desired names into a list 
            copydata("ankitjinames.txt");

            // create a new dictionary so that the modifications will not reflect in orginal dictionary 
            Dictionary<string, List<object>> dataCopy = new Dictionary<string, List<object>>(data);

          

            Filter(dataCopy, list);

            //upto here good 

            // data which tells which names are in which state
            var arr2 = GetObjArr(@"C:\ratlamfile\statewisenames.xlsx");
            List<object> listofstatewisenamesforankit = new List<object>();
            for (int i = 1; i <= arr2.GetLength(0); i++)
            {
                if (arr2[i, 1] != null && arr2[i, 2] == null)
                {
                    listofstatewisenamesforankit.Add(arr2[i, 1]);
                }
                if (arr2[i, 2] != null && arr2[i, 1] == null)
                {
                    listofstatewisenamesforankit.Add(arr2[i, 2]);
                }
            }

            Excel.Workbook workbook2 = excel.Workbooks.Add();
            Excel._Worksheet worksheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
            worksheet2.Name = "Ankit";
            // Add headers to worksheet2
            Excel.Range headerRange2 = worksheet2.Range["A1:E1"];
            headerRange2.Merge();
            headerRange2.Value = "Sales + payment report Current Month -Ankit ji";
            headerRange2.Font.Bold = true;
            headerRange2.Font.Size = 18;
           // headerRange2.Interior.Color = System.Drawing.Color.Yellow;

            //string outputpath = @"C:\ratlamfile\Ankit_ji_Ratlam-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
            List<string> states = new List<string>() { "U.P", "Rajasthan", "Bihar", "Punjab", "Odisha", "Chhatisgarh", "West bengal", "Madhya Pradesh", "Jharkhand", "Maharashtra", "Market", "Uttarakhand", "Assam", "Tripura" };
            int currowIndex = 2; // starting row index
            for (int i = 0; i < listofstatewisenamesforankit.Count; i++)
            {
                if (states.Contains(listofstatewisenamesforankit[i].ToString().Trim()))
                {
                   
                    worksheet2.Cells[currowIndex, 1].Value = listofstatewisenamesforankit[i].ToString();
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].EntireColumn.Font.Bold = true;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Merge();
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].HorizontalAlignment = XlHAlign.xlHAlignCenter;                    
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Cells.Font.Size = 20;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Font.Italic = true;
                    

                    worksheet2.Cells[currowIndex+1, 1] = "Party Name";
                    worksheet2.Cells[currowIndex+1, 2] = "Current Balance";
                    worksheet2.Cells[currowIndex+1, 3] = "Total Sales";
                    worksheet2.Cells[currowIndex+1, 4] = "Total Receipt";
                    worksheet2.Cells[currowIndex + 1, 5] = "Last payment Date";
                    worksheet2.Range["A" + (currowIndex + 1).ToString() + ":E" + (currowIndex + 1).ToString()].Font.Bold = true;
                    worksheet2.Range["A" + (currowIndex + 1).ToString() + ":E" + (currowIndex + 1).ToString()].Interior.Color = System.Drawing.Color.LightSeaGreen;
                    currowIndex += 2; // move to next row

                }
                else
                {
                    List<object> finalvalues = dataCopy[listofstatewisenamesforankit[i].ToString()];
                    string stringValue = finalvalues[4].ToString();
                    DateTime dateValue;
                    if (DateTime.TryParse(stringValue, out dateValue)) // try to convert cell value to DateTime
                    {
                        worksheet2.Cells[currowIndex, 5].Value = dateValue;
                        worksheet2.Cells[currowIndex, 5].NumberFormat = "dd-mm-yyyy"; // explicitly set the cell format
                    }
                    else
                    {
                        worksheet2.Cells[currowIndex, 5].Value = stringValue; // use string value as it is
                    }
                    worksheet2.Cells[currowIndex, 1].Value = listofstatewisenamesforankit[i].ToString();
                    worksheet2.Cells[currowIndex, 2].Value = finalvalues[6];
                    worksheet2.Cells[currowIndex, 3].Value = finalvalues[0];
                    worksheet2.Cells[currowIndex, 4].Value = finalvalues[3];
                    worksheet2.Cells[currowIndex, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    worksheet2.Cells[currowIndex, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    worksheet2.Cells[currowIndex, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    // worksheet2.Cells[rowIndex, colIndex + 4].Value = finalvalues[3];
                    currowIndex++; // move to next row
                }
            }
            
            worksheet2.Range["C1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            worksheet2.Range["D1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            
            Range rangeforoffice = worksheet2.UsedRange;
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
            worksheet2.Columns["A:H"].AutoFit(); // after bold , now autofit , not before making bold  
            worksheet2.UsedRange.Select();
            workbook2.SaveAs(@"C:\ratlamfile\Ankit ji excel Report - " + DateTime.Now.ToString("dd-MM-yyyy"));
            workbook2.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook2);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet2);

            // attach this created file and send to whatsapp
            // Set up Chrome driver
            // find correct version of driver at https://sites.google.com/chromium.org/driver/downloads?authuser=0
            // add chromedriver to environment variables 
            ///chromedriver -version  // check before 
            /////cd "C:\Program Files\Google\Chrome\Application"
            /// chrome.exe --remote-debugging-port=9222 --user-data-dir=D:\chromedata
            ///
             Process proc = new Process();
            // before copying below please note c:\ instaed of c 
             proc.StartInfo.FileName = @"C\Program Files\Google\Chrome\Application\chrome.exe";
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






        }


        // for ankit 
        public void Method2()
        {
            // Create a source dictionary to store data
            Dictionary<string, List<object>> data = new Dictionary<string, List<object>>();

            // Loop through each cell in the used range starting from row 3
            for (int startrow = 3; startrow <= rowindex - 1; startrow++)
            {
                decimal valueofcolumn11;
                string cellValue = range1.Cells[startrow, 11].NumberFormat.ToString();
                cellValue = cellValue.Replace("\"", "");
                //// Check if the cell value contains the string "cr"
                if (cellValue.Contains("Cr"))
                {
                    // Remove the "cr" from the cell value and convert it to a decimal
                    valueofcolumn11 = (decimal)range1.Cells[startrow, 11].Value * -1;

                }
                else
                {
                    // Convert the cell value to a decimal
                    valueofcolumn11 = (decimal)range1.Cells[startrow, 11].Value;
                }

                // Get key value from first column
                string key = range1.Cells[startrow, 1].Value.ToString();

                // Loop through remaining columns and add values to dictionary
                List<object> valuesinlist = new List<object>();
                for (int startcol = 3; startcol <= columnindex; startcol++)
                {
                    if (startcol == 3 || startcol == 6 || startcol == 7 || startcol == 11)
                    {
                        string value = range1.Cells[startrow, startcol].Value?.ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            if (startcol == 11)
                            {
                                value = string.Format("{0:N} ₹", valueofcolumn11);
                            }
                            else
                            {
                                value = string.Format("{0:N} ₹", decimal.Parse(value));
                            }
                        }
                        else
                        {
                            value = "not found in current month";
                        }
                        valuesinlist.Add(value);
                    }
                    else if (startcol != 5 && startcol != 9) // read every column except column 5 and column 9
                    {
                        object value = range1.Cells[startrow, startcol].Value;
                        if (value == null || string.IsNullOrEmpty(value.ToString()))
                        {
                            valuesinlist.Add("not found in current month");
                        }
                        else
                        {
                            valuesinlist.Add(value.ToString());
                        }
                    }
                }
                // Add the value of column 11 to the valuesinlist
                valuesinlist.Add(string.Format("{0:N} ₹", valueofcolumn11));
                data.Add(key, valuesinlist);
            }


            // Sort the dictionary by key in alphabetical order
            data = data.OrderBy(d => d.Key).ToDictionary(d => d.Key, d => d.Value);




            // copy desired names into a list 
            copydata("harshitjinames.txt");

            // create a new dictionary so that the modifications will not reflect in orginal dictionary 
            Dictionary<string, List<object>> dataCopy = new Dictionary<string, List<object>>(data);



            Filter(dataCopy, list);

            //upto here good 

            // data which tells which names are in which state
            var arr2 = GetObjArr(@"C:\ratlamfile\harshitstatewisenames.xlsx");
            List<object> listofstatewisenamesforankit = new List<object>();
            for (int i = 1; i <= arr2.GetLength(0); i++)
            {
                if (arr2[i, 1] != null && arr2[i, 2] == null)
                {
                    listofstatewisenamesforankit.Add(arr2[i, 1]);
                }
                if (arr2[i, 2] != null && arr2[i, 1] == null)
                {
                    listofstatewisenamesforankit.Add(arr2[i, 2]);
                }
            }

            Excel.Workbook workbook2 = excel.Workbooks.Add();
            Excel._Worksheet worksheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
            worksheet2.Name = "Ankit";
            // Add headers to worksheet2
            Excel.Range headerRange2 = worksheet2.Range["A1:E1"];
            headerRange2.Merge();
            headerRange2.Value = "Sales + payment report current month -Harsheet Ji";
            headerRange2.Font.Bold = true;
            headerRange2.Font.Size = 18;
            // headerRange2.Interior.Color = System.Drawing.Color.Yellow;

            //string outputpath = @"C:\ratlamfile\Ankit_ji_Ratlam-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
            List<string> states = new List<string>() { "U.P", "Rajasthan", "Bihar", "Punjab", "Odisha", "Chhatisgarh", "West bengal", "Madhya Pradesh", "Jharkhand", "Maharashtra", "Market", "Uttarakhand", "Assam", "Tripura" };
            int currowIndex = 2; // starting row index
            for (int i = 0; i < listofstatewisenamesforankit.Count; i++)
            {
                if (states.Contains(listofstatewisenamesforankit[i].ToString().Trim()))
                {

                    worksheet2.Cells[currowIndex, 1].Value = listofstatewisenamesforankit[i].ToString();
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].EntireColumn.Font.Bold = true;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Merge(); // merge should come here only 
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Cells.Font.Size = 20;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Font.Italic = true;


                    worksheet2.Cells[currowIndex + 1, 1] = "Party Name";
                    worksheet2.Cells[currowIndex + 1, 2] = "Current Balance";
                    worksheet2.Cells[currowIndex + 1, 3] = "Total Sales";
                    worksheet2.Cells[currowIndex + 1, 4] = "Total Receipt";
                    worksheet2.Cells[currowIndex + 1, 5] = "Last payment Date";
                    worksheet2.Range["A" + (currowIndex + 1).ToString() + ":E" + (currowIndex + 1).ToString()].Font.Bold = true;
                    worksheet2.Range["A" + (currowIndex + 1).ToString() + ":E" + (currowIndex + 1).ToString()].Interior.Color = System.Drawing.Color.LightSeaGreen;
                    currowIndex += 2; // move to next row

                }
                else
                {
                    List<object> finalvalues = dataCopy[listofstatewisenamesforankit[i].ToString()];
                    string stringValue = finalvalues[4].ToString();
                    DateTime dateValue;
                    if (DateTime.TryParse(stringValue, out dateValue)) // try to convert cell value to DateTime
                    {
                        worksheet2.Cells[currowIndex, 5].Value = dateValue;
                        worksheet2.Cells[currowIndex, 5].NumberFormat = "dd-mm-yyyy"; // explicitly set the cell format
                    }
                    else
                    {
                        worksheet2.Cells[currowIndex, 5].Value = stringValue; // use string value as it is
                    }
                    worksheet2.Cells[currowIndex, 1].Value = listofstatewisenamesforankit[i].ToString();
                    worksheet2.Cells[currowIndex, 2].Value = finalvalues[6];
                    worksheet2.Cells[currowIndex, 3].Value = finalvalues[0];
                    worksheet2.Cells[currowIndex, 4].Value = finalvalues[3];
                    worksheet2.Cells[currowIndex, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    worksheet2.Cells[currowIndex, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    worksheet2.Cells[currowIndex, 5].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    // worksheet2.Cells[rowIndex, colIndex + 4].Value = finalvalues[3];
                    currowIndex++; // move to next row
                }
            }

            worksheet2.Range["C1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            worksheet2.Range["D1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            Range rangeforoffice = worksheet2.UsedRange;
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
            worksheet2.Columns["A:H"].AutoFit(); // after bold , now autofit , not before making bold  
            worksheet2.UsedRange.Select();
            workbook2.SaveAs(@"C:\ratlamfile\Harsheet ji excel Report - " + DateTime.Now.ToString("dd-MM-yyyy"));
            workbook2.Close();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook2);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet2);
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
