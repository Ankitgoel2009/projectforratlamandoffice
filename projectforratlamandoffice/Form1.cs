using System;using System.Collections.Generic;using System.ComponentModel;using System.Data;using System.Drawing;using System.Linq;using System.Text;using System.Threading.Tasks;using System.Windows.Forms;using Excel = Microsoft.Office.Interop.Excel;using System.IO;using Microsoft.Office.Interop.Excel;using System.Globalization;using System.Text.RegularExpressions;
using System.Collections.Specialized;
using System.Collections;

namespace projectforratlamandoffice{    public partial class Form1 : Form    {        string selectedFile;        List<string> list = new List<string>();        public Form1()        {            InitializeComponent();        }        private void Form1_Load(object sender, EventArgs e)        {
            button1.Enabled = false;            comboBox1.SelectedIndex = 0;            label1.Visible = false;            comboBox2.Visible = false;            button2.Visible = false;

        }        void OpenKeywordsFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)        {            OpenFileDialog fileDialog = sender as OpenFileDialog;            selectedFile = System.IO.Path.GetFileNameWithoutExtension(fileDialog.FileName);            if (string.IsNullOrEmpty(selectedFile) || selectedFile.Contains(".lnk"))            {                MessageBox.Show("Please select a valid Excel File");                e.Cancel = true;            }            return;        }        private void button1_Click(object sender, EventArgs e)        {            OpenFileDialog choofdlog = new OpenFileDialog();            choofdlog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);            choofdlog.Multiselect = false;            choofdlog.ValidateNames = true;            choofdlog.DereferenceLinks = false; // Will return .lnk in shortcuts.
            choofdlog.Filter = "Excel |*.xlsx";            choofdlog.FilterIndex = 1;            choofdlog.FileOk += new System.ComponentModel.CancelEventHandler(OpenKeywordsFileDialog_FileOk);            DialogResult result = choofdlog.ShowDialog();            if (result == DialogResult.OK)            {                selectedFile = choofdlog.FileName;
                button1.Enabled = false;                label1.Visible = true;                comboBox2.Visible = true;
            }
            else if (result == DialogResult.Cancel)
            {
                //User pressed cancel
                comboBox1.SelectedIndex = 0;
                return;
            }
                }

        public static Excel.Workbook Open(Excel.Application excelInstance, string fileName, bool readOnly = false, bool editable = true, bool updateLinks = true)        {            Excel.Workbook book = excelInstance.Workbooks.Open(                fileName, updateLinks, readOnly,                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,                Type.Missing, editable, Type.Missing, Type.Missing, Type.Missing,                Type.Missing, Type.Missing);            return book;        }
        // METHOD TO LOOK and REPLACE SIMILAR LOOKING strings  
        static string CaseInsenstiveReplace(string originalString, string oldValue, string newValue)
        {
            Regex regEx = new Regex(oldValue,
            RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return regEx.Replace(originalString, newValue);
        }

        void dowork()        {            Excel.Application excel = null;            excel = new Excel.Application();            excel.Visible = true;            Excel.Workbook wkb = null;            wkb = Open(excel, selectedFile);            Excel._Worksheet sheet = wkb.Sheets[1];
            Excel.Range range1 = sheet.UsedRange;

            // start
            var arr1 = (object[,])range1.get_Value(XlRangeValueDataType.xlRangeValueDefault);            var listname = comboBox1.SelectedItem.ToString();

            // step 1 .  listnew contains all the names of all the models above 60 
            // ok block 
            List<object[]> listnew = new List<object[]>();            listnew.Add(new object[] { listname });            for (int i = 10; i <= arr1.GetLength(0) - 1; i++)            {
                // replace pcs from second row and convert to int 
                if (arr1[i, 1] != null)                {                    int id = Convert.ToInt32(arr1[i, 2].ToString().Replace("Pcs", "").Trim());                    if (id > Convert.ToInt32(comboBox2.SelectedItem)) // remove below 60 pcs 
                    {
                        listnew.Add(new object[] { CaseInsenstiveReplace(arr1[i, 1].ToString(), listname, "").Trim() });
                    }                }            }

            // list of all names are ready

            // step 2 . extract sub heading from this list and store it in list1
            // ok block
            int firstSpaceIndex;            String firstString;            List<string> list1 = new List<string>(); // just for sub-headings 
            for (int i = 1; i < listnew.Count; i++)            {
                firstSpaceIndex = listnew[i][0].ToString().IndexOf(" "); // get index upto first space found
                firstString = listnew[i][0].ToString().Substring(0, firstSpaceIndex); // get string upto first space found
                if (list1.Contains(firstString, StringComparer.CurrentCultureIgnoreCase)) // if not previously present in list 
                {                    continue;                   // list1.Add(firstString); 
                }
                else
                {
                    list1.Add(firstString); // Add to sub-heading list 
                }
            }
            // List of sub-heading are  ready , now prepare final list

            // create a datatype in which i can store headings as key and 
            // their respective data  as their values 
            // for this dictionary is used . but 
            //    https://www.ict.ru.ac.za/resources/thinksharply/thinksharply/dictionaries.html 
            // the order in which a dictionary stores its pairs is unpredictable. (So the order in which we’ll get them delivered by a foreach becomes unpredictable.
            //ok block
            OrderedDictionary dic = new OrderedDictionary(StringComparer.CurrentCultureIgnoreCase);  // string comparer is used so that all evaluations on the key act according to the rules of the comparer: case-insensitive.                  
            for (int i = 0; i < list1.Count; i++)            {                string key = list1[i];                List<string> l = new List<string>();                for (var item = 1; item < listnew.Count; item++) // deliberately starts from 1 
                {                    if (listnew[item][0].ToString().StartsWith(key, StringComparison.OrdinalIgnoreCase))                    {                        l.Add(listnew[item][0].ToString());
                    }                }
                if (String.Equals(key, "sam", StringComparison.OrdinalIgnoreCase))
                {
                    dic.Add("Samsung", l);
                }
                else if (String.Equals(key, "moto", StringComparison.OrdinalIgnoreCase))
                {
                    dic.Add("Motorola", l);
                }
                else if (String.Equals(key, "Z", StringComparison.OrdinalIgnoreCase)) // this else-if block will move to upper block
                {
                    continue;
                }
                else
                {
                    dic.Add(key, l);
                }

            }
            // sort in descending order         

            // Create a temporary list to store the keys and values of the dictionary.
            List<KeyValuePair<string, List<string>>> templist = new List<KeyValuePair<string, List<string>>>();
            foreach (DictionaryEntry entry in dic)
            {
                templist.Add(new KeyValuePair<string, List<string>>(entry.Key.ToString(), (List<string>)entry.Value));
            }

            // Sort the list in descending order 
            templist.Sort((x, y) => y.Value.Count.CompareTo(x.Value.Count));
            // Clear the dictionary.
            dic.Clear();
            foreach (KeyValuePair<string, List<string>> item in templist)
            {
                dic.Add(item.Key.ToString(), item.Value);
            }


            // sort by listofsequence 
            List<string> listofsequence = new List<string>();            listofsequence.Add("Samsung");            listofsequence.Add("Redmi");            listofsequence.Add("Oppo");            listofsequence.Add("Vivo");
            foreach (var item in listofsequence)
            {
                foreach (var kvp in dic.Cast<DictionaryEntry>().Reverse().Reverse())
                {
                   // if (item == kvp.Key.ToString())
                   if(String.Equals(item, kvp.Key.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        dic.Remove(kvp.Key);
                        dic.Insert(listofsequence.IndexOf(item), kvp.Key, kvp.Value);
                    }
                }
            }

            // ok 
            // count each key and total values in each key 
            // to be used in creating excel file later in project 
            OrderedDictionary count = new OrderedDictionary();
            int finaltotal = 0;            foreach (var kvp in dic.Cast<DictionaryEntry>())
            {
                List<string> list = kvp.Value as List<string>;
                count.Add(kvp.Key.ToString(), list.Count);
            }
            finaltotal = count.Cast<DictionaryEntry>().Sum(i => Convert.ToInt32(i.Value));
            string outputpath = @"C:\Users\Public\"+comboBox1.Text+".xlsx";            Excel.Application excel1 = new Excel.Application();            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.ActiveSheet;

            // create main heading 
            sheet1.Cells[1, 1].Value = listnew[0][0].ToString().ToUpper();
            // stretch it to three columns and other designing 
            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Font.Bold = true;            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].HorizontalAlignment = XlHAlign.xlHAlignCenter;            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Merge();            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Cells.Font.Size = 30;            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Font.Italic = true;            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Interior.Color = Color.Yellow;           // max columns that i want             var columnMax = 3;
            var remainder = 0;
            var rowMaxineachcolumn = 0;
            rowMaxineachcolumn = Math.DivRem((finaltotal + dic.Keys.Count + ((dic.Keys.Count - 1) * 2)), columnMax, out remainder);
            if (remainder == 1 || remainder == 2)
            {
                rowMaxineachcolumn = rowMaxineachcolumn + 1;
            }
            else
            {
                rowMaxineachcolumn = rowMaxineachcolumn + remainder;
            }

            int columnnumber = 1;            int rownumber = 1;
            bool nextcolumn = true;
            int rowsinpreviouscolumn=0;
            int rowsincurrentcolumn=0;

            // code block for looping in dictionary 
            for (int i = 0; i < dic.Count; i++)
            {
                if (columnnumber > 3) // if you reach the last column 
                {
                   columnnumber = 1; // get back to column 1                   
                }             
                    if (columnnumber == 1)
                    {
                        rownumber = getlastrowincolumn(sheet1, "A");
                        rownumber++;  // start printing from row2                  
                    }
                    else if (columnnumber == 2)
                    {
                        rownumber = getlastrowincolumn(sheet1, "B");                       
                        rownumber++;  // start printing from row2                   
                }
                    else if (columnnumber == 3)
                    {
                        rownumber = getlastrowincolumn(sheet1, "C");
                        rownumber++; // start printing from row2
                }


                // code for printing headings ie samsung , redmi , oppo , vivo , techno etc 
                string key = ((DictionaryEntry)dic.Cast<DictionaryEntry>().ElementAt(i)).Key.ToString();               
                sheet1.Cells[rownumber, columnnumber].Value = key.ToUpper();
                sheet1.Cells[rownumber, columnnumber].Font.Bold = true;
                sheet1.Cells[rownumber, columnnumber].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                sheet1.Cells[rownumber, columnnumber].Cells.Font.Size = 20;
                sheet1.Cells[rownumber, columnnumber].Interior.Color =  System.Drawing.Color.FromArgb(172, 185, 202);

                List<string> values = (List<string>)dic[key];
                foreach (string value in values)
                {
                   
                    rownumber++; // change row 
                    sheet1.Cells[rownumber, columnnumber].Value = value;
                    sheet1.Cells[rownumber, columnnumber].Font.Bold = true;
                    sheet1.Cells[rownumber, columnnumber].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    sheet1.Cells[rownumber, columnnumber].Cells.Font.Size = 10;
                    // sheet1.Cells[rows, column].Interior.Color = Color.Green;
                }
                if (i + 1 < dic.Count)
                {
                   
                    string nextKey = (string)dic.Cast<DictionaryEntry>().ElementAt(i + 1).Key;
                    if (!listofsequence.Contains(nextKey, StringComparer.CurrentCultureIgnoreCase))
                    {
                        // Print the number of values for this key.
                        int countofnextkey_valuepair = ((((DictionaryEntry)dic.Cast<DictionaryEntry>().ElementAt(i + 1)).Value as List<string>).Count) + 1;
                        if(nextcolumn == true)
                        {
                            columnnumber++;
                            nextcolumn = false;
                        }
                        else if(nextcolumn == false)
                        {
                            
                            if(columnnumber ==2)
                            {
                                rowsinpreviouscolumn = getlastrowincolumn(sheet1, "A")-2; // use getlastrowincolumn(sheet1, "A")-2; here if row exceeds 
                                rowsincurrentcolumn = getlastrowincolumn(sheet1, "B");

                            }
                            else if(columnnumber==3)
                            {
                                rowsinpreviouscolumn = getlastrowincolumn(sheet1, "B")-2; // use getlastrowincolumn(sheet1, "B")-2; here if row exceeds 
                                rowsincurrentcolumn = getlastrowincolumn(sheet1, "C");
                            }
                            if (rowsinpreviouscolumn < rowsincurrentcolumn + countofnextkey_valuepair )
                            {
                                columnnumber++;
                                
                            }
                            else
                            {
                                rownumber = rownumber + 2;
                            }
                        }
                    }
                    else
                    {
                        columnnumber++;
                    }
                }


            }
            sheet1.Columns["A:C"].AutoFit();
            sheet1.Columns["A:A"].ColumnWidth = 40.57;
            sheet1.Columns["B:B"].ColumnWidth = 40.57;
            sheet1.Columns["C:C"].ColumnWidth = 40.57;
           // sheet1.Range["B1"].ColumnWidth = 30.00;
               // sheet1.Range["C1"].ColumnWidth = 30.00;
                sheet1.Range["A1"].EntireColumn.Font.Bold = true;
                sheet1.Range["B1"].EntireColumn.Font.Bold = true;
                sheet1.Range["C1"].EntireColumn.Font.Bold = true;
               
                // sheet1.Columns["A"].AutoFit();
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
                excel1.Quit();
                // CLEAN UP.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);
                wkb.Close(true);
                excel.Quit();
                // CLEAN UP.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                System.Windows.Forms.Application.Exit();
            }


        public static int getlastrow(Worksheet ws)
        {
            return ws.Cells.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious).Row;
          }
        public static int getlastrowincolumn(Worksheet ws, string column)
        {
            for (int x = getlastrow(ws); x > 0; x--)
            {
                if (("" + ws.Range[column + x].Value) != "")
                {
                    if (x > 1)
                    {
                        return x=x+2;
                    }
                    else
                    {
                        return x;
                    }
                }
            }
            return 1;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 1 || comboBox1.SelectedIndex == 2 || comboBox1.SelectedIndex == 3
               || comboBox1.SelectedIndex == 4 || comboBox1.SelectedIndex == 5 || comboBox1.SelectedIndex == 6)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
                label1.Visible = false;
                comboBox2.Visible = false;
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;            dowork();
            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex >= 0)
            {
                button2.Visible = true;


            }
            else
            {
                button1.Enabled = false;
                label1.Visible = false;
                comboBox2.Visible = false;
                button2.Visible = false;
            }
        }
    }
     
    }
