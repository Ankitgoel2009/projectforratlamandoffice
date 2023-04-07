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
           // sheet.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
          //  sheet.Range["G1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
           // sheet.Range["H1"].EntireColumn.NumberFormat = "DD/MM/YYYY";
            range1 = sheet.UsedRange; // returns range till grandtotal 
            rowindex = range1.Rows.Count;
            columnindex = range1.Columns.Count;
            Method1();
           // Method2();
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
                copydata.Add(key, new List<object> { 0,0,0,0,0,0,0 });
            }
        }

        public void Method1()
        {

            // Create a source dictionary to store data
            Dictionary<string, List<object>> data = new Dictionary<string, List<object>>();
            
            // Loop through each cell in the used range starting from row 3
            for (int startrow = 3; startrow <= rowindex-1; startrow++)
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
                    if (startcol == 11)
                    {
                        // Add the Rupee symbol to the value and add it to the list
                        string value = valueofcolumn11 != null ? "₹" + valueofcolumn11.ToString() : "No records found";
                        valuesinlist.Add(value);
                    }
                    else if (startcol != 5 && startcol != 9)
                    {
                        // Add the Rupee symbol to the value and add it to the list
                        object value = range1.Cells[startrow, startcol].Value;
                        string stringValue = value != null ? "₹" + value.ToString() : "No records found";
                        valuesinlist.Add(stringValue);
                    }
                    else
                    {
                        // Add null value to the list
                        valuesinlist.Add(null);
                    }
                }
                        data.Add(key, valuesinlist);
                
            }
            // Sort the dictionary by key in alphabetical order
            data = data.OrderBy(d => d.Key).ToDictionary(d => d.Key, d => d.Value);

            // copy desired names into a list 
            copydata("ankitjinames.txt");

            // create a new dictionary so that the modifications will not reflect in orginal dictionary 
            Dictionary<string, List<object>> dataCopy = new Dictionary<string, List<object>>(data);

            Filter(dataCopy, list);

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
            headerRange2.Value = "Ankit ji This week";
            headerRange2.Font.Bold = true;
            headerRange2.Font.Size = 18;
            headerRange2.Interior.Color = System.Drawing.Color.Yellow;

            //string outputpath = @"C:\ratlamfile\Ankit_ji_Ratlam-" + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".xlsx";
            List<string> states = new List<string>() { "U.P", "Rajasthan", "Bihar", "Punjab", "Odisha", "Chhatisgarh", "West bengal", "Madhya Pradesh", "Jharkhand", "Maharashtra", "Market", "Uttarakhand", "Assam", "Tripura" };
            int currowIndex = 2; // starting row index
            for (int i = 0; i < listofstatewisenamesforankit.Count; i++)
            {
                if (states.Contains(listofstatewisenamesforankit[i].ToString().Trim()))
                {
                    
                    worksheet2.Cells[currowIndex, 1].Value = listofstatewisenamesforankit[i].ToString();
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].EntireColumn.Font.Bold = true;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].HorizontalAlignment = XlHAlign.xlHAlignCenter;                    
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Cells.Font.Size = 20;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Font.Italic = true;
                    worksheet2.Range[worksheet2.Cells[currowIndex, 1], worksheet2.Cells[currowIndex, 5]].Merge();

                    worksheet2.Cells[currowIndex+1, 1] = "Party Name";
                    worksheet2.Cells[currowIndex+1, 2] = "Current Balance";
                    worksheet2.Cells[currowIndex+1, 3] = "Total Sales";
                    worksheet2.Cells[currowIndex+1, 4] = "Total Receipt";
                    //  worksheet2.Range["A2:E2"].Font.Bold = true;
                    currowIndex += 2; // move to next row

                }
                else
                {
                    List<object> finalvalues = dataCopy[listofstatewisenamesforankit[i].ToString()];
                    worksheet2.Cells[currowIndex, 1 ].Value = listofstatewisenamesforankit[i].ToString();
                    worksheet2.Cells[currowIndex, 2].Value = finalvalues[6];
                    worksheet2.Cells[currowIndex, 3].Value = finalvalues[0];
                    worksheet2.Cells[currowIndex, 4].Value = finalvalues[3];
                    // worksheet2.Cells[rowIndex, colIndex + 4].Value = finalvalues[3];
                    currowIndex++; // move to next row
                }
            }

            // prepare file 
            Excel.Workbook workbook1 = excel.Workbooks.Add();
            
            Excel.Workbook workbook3 = excel.Workbooks.Add();

            // Set the worksheet names
            Excel._Worksheet worksheet1 = (Excel.Worksheet)workbook1.ActiveSheet;
            worksheet1.Name = "Office";
           
            Excel._Worksheet worksheet3 = (Excel.Worksheet)workbook3.ActiveSheet;
            worksheet3.Name = "Harsheet";

         











            // Add headers to worksheet3
            Excel.Range headerRange3 = worksheet3.Range["A1:D1"];
            headerRange3.Merge();
            headerRange3.Value = "Harsheet ji this week";
            headerRange3.Font.Bold = true;
            headerRange3.Font.Size = 18;
            headerRange3.Interior.Color = System.Drawing.Color.Yellow;

            worksheet3.Cells[2, 1] = "Name";
            worksheet3.Cells[2, 2] = "Current Balance";
            worksheet3.Cells[2, 3] = "Total Sales";
            worksheet3.Cells[2, 4] = "Total Receipt";
            worksheet3.Range["A2:D2"].Font.Bold = true;

            // Import data into worksheet2
            int rowofdestinationfiles = 3;
            foreach (var item in data)
            {
                string key = item.Key;
                List<object> values1 = item.Value;

                // Set name value
                worksheet2.Cells[rowofdestinationfiles, 1] = key;

                // Set current balance value
                worksheet2.Cells[rowofdestinationfiles, 2] = values1[6];

                // Set total sales value
                worksheet2.Cells[rowofdestinationfiles, 3] = values1[0];

                // Set total receipt value
                worksheet2.Cells[rowofdestinationfiles, 4] = values1[3];

                rowofdestinationfiles++;
            }

            // Import data into worksheet3
            rowofdestinationfiles = 3;
            foreach (var item in data)
            {
                string key = item.Key;
                List<object> values1 = item.Value;

                // Set name value
                worksheet3.Cells[rowofdestinationfiles, 1] = key;

                // Set current balance value
                worksheet3.Cells[rowofdestinationfiles, 2] = values1[6];

                // Set total sales value
                worksheet3.Cells[rowofdestinationfiles, 3] = values1[0];

                // Set total receipt value
                worksheet3.Cells[rowofdestinationfiles, 4] = values1[3];

                rowofdestinationfiles++;
            }

            





            // sort the source range 
            int lastRow = range1.Row + range1.Rows.Count - 1;
            int lastColumn = range1.Column + range1.Columns.Count - 1;

            // Sort the first column in ascending order, starting from row 3
            Excel.Range sortRange = sheet.Range[sheet.Cells[3, 1], sheet.Cells[lastRow - 1, lastColumn]];

            // Define the sort keys (sort by column 1 ascending)
            Excel.SortFields sortFields = sheet.Sort.SortFields;
            Excel.SortField sortField1 = sortFields.Add(sortRange.Columns[1], Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending);

            // Apply the sort
            sheet.Sort.SetRange(sortRange);
            sheet.Sort.Header = Excel.XlYesNoGuess.xlNo;
            sheet.Sort.MatchCase = false;
            sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
            sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            sheet.Sort.Apply();


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
                if (range1.Cells[i, 11].Value != null && range1.Cells[i, 11].Value > 1000 && !valueA.ToLower().Contains("udaan")&&!valueA.ToLower().Contains("shop"))
                {
                    decimal valueB; // = (decima                                                                                                                                       l)range1.Cells[i, 11].Value; // take the value in second column and convert to decimal 
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
                        range1.Cells[i, 10].Value = "B/F 15/11/2022";
                        sheetforoffice.Cells[row, ++column].Value = range1.Cells[i, 10].Value;
                    }
                    else
                    {
                        sheetforoffice.Cells[row, ++column].Value = valueC = (decimal)range1.Cells[i, 10].Value;
                    }
                    DateTime? valueD = range1.Cells[i, 8].Value as DateTime?;
                    if (valueD == null)
                    {
                        range1.Cells[i, 8].Value = "B/F 15/11/2022";
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
            //sheetforoffice.Columns[1].ColumnWidth = 25;
            //sheetforoffice.Columns[2].ColumnWidth = 8.25;
            //sheetforoffice.Columns[3].ColumnWidth = 9.72;
            //sheetforoffice.Columns[4].ColumnWidth = 9.56;
            //sheetforoffice.Columns[5].ColumnWidth = 25;
            //sheetforoffice.Columns[6].ColumnWidth = 8.25;
            //sheetforoffice.Columns[7].ColumnWidth = 9.72;
            //sheetforoffice.Columns[8].ColumnWidth = 9.56; ;


            // Set the row height for a range of rows
          //  Excel.Range rowRange = sheetforoffice.Range["59:" + sheetforoffice.Cells[sheetforoffice.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row];
          //  rowRange.RowHeight = 9.80;
          //  rowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

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


        // for ankit 
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
            copydata("ankitjinames.txt");

            // filter from only index which conatins  names and add the balance 
          //  Filter(listnew);

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
