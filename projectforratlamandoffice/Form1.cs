using System;using System.Collections.Generic;using System.ComponentModel;using System.Data;using System.Drawing;using System.Linq;using System.Text;using System.Threading.Tasks;using System.Windows.Forms;using Excel = Microsoft.Office.Interop.Excel;using System.IO;using Microsoft.Office.Interop.Excel;using System.Globalization;using System.Text.RegularExpressions;
using System.Collections.Specialized;
using System.Collections;

namespace projectforratlamandoffice{    public partial class Form1 : Form    {        string selectedFile;        List<string> list = new List<string>();        public Form1()        {            InitializeComponent();        }        private void Form1_Load(object sender, EventArgs e)        {                           }        void OpenKeywordsFileDialog_FileOk(object sender, System.ComponentModel.CancelEventArgs e)        {            OpenFileDialog fileDialog = sender as OpenFileDialog;            selectedFile = System.IO.Path.GetFileNameWithoutExtension(fileDialog.FileName);            if (string.IsNullOrEmpty(selectedFile) || selectedFile.Contains(".lnk"))            {                MessageBox.Show("Please select a valid Excel File");                e.Cancel = true;            }            return;        }        private void button1_Click(object sender, EventArgs e)        {            OpenFileDialog choofdlog = new OpenFileDialog();            choofdlog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);            choofdlog.Multiselect = false;            choofdlog.ValidateNames = true;            choofdlog.DereferenceLinks = false; // Will return .lnk in shortcuts.            choofdlog.Filter = "Excel |*.xlsx";            choofdlog.FilterIndex = 1;            choofdlog.FileOk += new System.ComponentModel.CancelEventHandler(OpenKeywordsFileDialog_FileOk);            DialogResult result = choofdlog.ShowDialog();            if (result == DialogResult.OK)            {                selectedFile = choofdlog.FileName;                   }
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            dowork();
            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;        }            public static Excel.Workbook Open(Excel.Application excelInstance, string fileName, bool readOnly = false, bool editable = true, bool updateLinks = true)        {            Excel.Workbook book = excelInstance.Workbooks.Open(                fileName, updateLinks, readOnly,                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,                Type.Missing, editable, Type.Missing, Type.Missing, Type.Missing,                Type.Missing, Type.Missing);            return book;        }
       // METHOD TO LOOK and REPLACE SIMILAR LOOKING strings  
        static string CaseInsenstiveReplace(string originalString, string oldValue, string newValue)
        {
            Regex regEx = new Regex(oldValue,
            RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return regEx.Replace(originalString, newValue);
        }

        void dowork()        {            Excel.Application excel = null;            excel = new Excel.Application();            excel.Visible = true;            Excel.Workbook wkb = null;            wkb = Open(excel, selectedFile);            Excel._Worksheet sheet = wkb.Sheets[1];                      Excel.Range range1 = sheet.UsedRange;            // start            var arr1 = (object[,])range1.get_Value(XlRangeValueDataType.xlRangeValueDefault);            var listname = "Vintage Flip";                       // step 1 .  listnew contains all the names of all the models above 60             // ok block             List<object[]> listnew = new List<object[]>();            listnew.Add(new object[] { listname });            for (int i = 10; i <= arr1.GetLength(0) - 1; i++)            {                // replace pcs from second row and convert to int                 if (arr1[i, 1] != null)                {                    int id = Convert.ToInt32(arr1[i, 2].ToString().Replace("Pcs", "").Trim());                    if (id > 60) // remove below 60 pcs                     {
                        listnew.Add(new object[] { CaseInsenstiveReplace(arr1[i, 1].ToString(), listname, "").Trim()});                                                               }                }            }            // list of all names are ready            // step 2 . extract sub heading from this list and store it in list1            // ok block            int firstSpaceIndex;            String firstString;            List<string> list1 = new List<string>(); // just for sub-headings             for (int i = 1; i < listnew.Count; i++)            {                               firstSpaceIndex = listnew[i][0].ToString().IndexOf(" "); // get index upto first space found                firstString = listnew[i][0].ToString().Substring(0, firstSpaceIndex); // get string upto first space found                if (!list1.Contains(firstString)) // if not previously present in list                 {                    list1.Add(firstString); // Add to sub-heading list                 }             }
            // List of sub-heading are  ready , now prepare final list


           
                        // create a datatype in which i can store headings as key and             // their respective data  as their values             // for this dictionary is used .             //ok block            OrderedDictionary dic = new OrderedDictionary(StringComparer.CurrentCultureIgnoreCase);  // string comparer is used so that all evaluations on the key act according to the rules of the comparer: case-insensitive.                  
            for (int i = 0; i < list1.Count; i++)            {                string key = list1[i];                List<string> l = new List<string>();                for (var item=1;item<listnew.Count;item++) // deliberately starts from 1                 {                    if (listnew[item][0].ToString().StartsWith(key))                    {                        l.Add(listnew[item][0].ToString());                    }                }                                                                         if (String.Equals(key, "sam", StringComparison.OrdinalIgnoreCase))                    {                        dic.Add("Samsung", l);                    }
                    else if (String.Equals(key, "moto", StringComparison.OrdinalIgnoreCase))                    {                        dic.Add("Motorola", l);                    }                    else if (String.Equals(key, "Z", StringComparison.OrdinalIgnoreCase)) // this else-if block will move to upper block                    {                        continue;                    }                    else                    {                        dic.Add(key, l);                    }                                                }
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
                    if (item == kvp.Key.ToString())
                    {
                        dic.Remove(kvp.Key);
                        dic.Insert(listofsequence.IndexOf(item), kvp.Key, kvp.Value);
                    }
                }
            }
        
            // ok 
            // count each key and total values in each key 
            // to be used in creating excel file later in project 
            OrderedDictionary count = new OrderedDictionary();                       int finaltotal=0;            foreach (var kvp in dic.Cast<DictionaryEntry>())
            {
                List<string> list = kvp.Value as List<string>;
                count.Add(kvp.Key.ToString(), list.Count);
            }
            finaltotal = count.Cast<DictionaryEntry>().Sum(i => Convert.ToInt32(i.Value));
             //    https://www.ict.ru.ac.za/resources/thinksharply/thinksharply/dictionaries.html             // the order in which a dictionary stores its pairs is unpredictable. (So the order in which we’ll get them delivered by a foreach becomes unpredictable.                

            // sort count according to list named sequences
            // count = count.OrderBy(d => listofsequence.IndexOf(d.Key)).ToDictionary(x => x.Key, x => x.Value);

            // sort dic according to list named sequence
            //  dic = dic.OrderBy(d => listofsequence.IndexOf(d.Key)).ToDictionary(x => x.Key, x => x.Value);
                     //  dic = dic.OrderByDescending(d => d.Value.Count).ToDictionary(x => x.Key, x => x.Value);
          //   dic = dic.OrderByDescending(d => listofsequence.IndexOf(d.Key)).ToDictionary(x => x.Key, x => x.Value);
            // ready for output 
            string outputpath = @"C:\Users\Public\vintagelist.xlsx";            Excel.Application excel1 = new Excel.Application();            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.ActiveSheet;            // create main heading             sheet1.Cells[1, 1].Value = listnew[0][0].ToString().ToUpper();            // stretch it to three columns and other designing             sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Font.Bold = true;            sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 3]].HorizontalAlignment = XlHAlign.xlHAlignCenter;            sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 3]].Merge();            sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 3]].Cells.Font.Size = 30;            sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 3]].Font.Italic = true;            sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 3]].Interior.Color = Color.Yellow;            int columnnumber = 1;            int rownumber = 1;            var columnMax = 3;
            var remainder = 0;
            var rowMaxineachcolumn = 0;
            rowMaxineachcolumn = Math.DivRem((finaltotal + dic.Keys.Count + ((dic.Keys.Count - 1) * 2)), 3, out remainder);
            if (remainder == 1 || remainder == 2)
            {
                rowMaxineachcolumn = rowMaxineachcolumn + 1;
            }
            else
            {
                rowMaxineachcolumn = rowMaxineachcolumn + remainder;
            }                        //  var totalrowsmax = finaltotal;            bool extrarow = true;
            
            foreach (DictionaryEntry entry in dic)
            {
                var LastRow = sheet1.UsedRange.Rows.Count;
                var LastCol = sheet1.UsedRange.Columns.Count;
                //if ((rownumber + dic[keyvalue].Count) > rowMaxineachcolumn)
                //{
                //    rownumber = 1;
                //    columnnumber++;
                //}
                // if (extrarow == false)
                //{
                //  rownumber += 2;
                //}
                if (columnnumber > 3) // if you reach the last column 
                {
                    rownumber++; // print heading in next line 
                    columnnumber = 1; // get back to column 1 
                }
                else
                {
                    if (rownumber == 1)
                    {
                        rownumber = LastRow +1;
                        //rownumber++;
                    }
                    else
                    {
                        rownumber = LastRow++;
                    }
                }
                // code for printing headings 
                sheet1.Cells[rownumber, columnnumber].Value = entry.Key.ToString().ToUpper();
                sheet1.Cells[rownumber, columnnumber].Font.Bold = true;
                sheet1.Cells[rownumber, columnnumber].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                // sheet1.Cells[rownumber, columnnumber].Cells.Font.Size = 20;
                sheet1.Cells[rownumber, columnnumber].Interior.Color = Color.Blue;
                foreach (string value in (List<string>)entry.Value)
                {                    //if (rownumber > rowMaxineachcolumn)                    //{                    //    rownumber = 1;                    //    columnnumber++;                    //}
                    rownumber++;
                    sheet1.Cells[rownumber, columnnumber].Value = value;
                    sheet1.Cells[rownumber, columnnumber].Font.Bold = true;
                    sheet1.Cells[rownumber, columnnumber].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    sheet1.Cells[rownumber, columnnumber].Cells.Font.Size = 10;
                    // sheet1.Cells[rows, column].Interior.Color = Color.Green;
                }                extrarow = false;
                columnnumber++;

            }

            //
            sheet1.Range["B1"].ColumnWidth = 30.00;            sheet1.Range["C1"].ColumnWidth = 30.00;            sheet1.Range["A1"].EntireColumn.Font.Bold = true;            sheet1.Range["B1"].EntireColumn.Font.Bold = true;            sheet1.Range["C1"].EntireColumn.Font.Bold = true;            sheet1.Columns["A:C"].AutoFit();
            // sheet1.Columns["A"].AutoFit();
            range1 = sheet1.UsedRange;            Borders border = range1.Borders;            border[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;            border[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;            border[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;            border[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;            border.Color = Color.Black;            border[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;            border[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;            border[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;            border[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;            range1.Borders.Color = Color.Black;            range1.Select();            sheet1.UsedRange.Select();            workbook.SaveAs(outputpath);            workbook.Close();            excel1.Quit();            // CLEAN UP.            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet1);            wkb.Close(true);            excel.Quit();            // CLEAN UP.            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);            System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);            System.Windows.Forms.Application.Exit();        }

      
    }}