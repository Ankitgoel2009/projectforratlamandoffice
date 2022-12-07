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

namespace projectforratlamandoffice
{
    public partial class Form1 : Form
    {
        string selectedFile;
        int iRow, iCol = 1;
        object[,] valueArray;
        List<string> list = new List<string>();
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


            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                selectedFile = choofdlog.FileName;

            }
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            dowork();
            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;

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
        void dowork()
        {
            Excel.Application excel = null;
            excel = new Excel.Application();
            excel.Visible = true;
            Excel.Workbook wkb = null;
            wkb = Open(excel, selectedFile);
            Excel._Worksheet sheet = wkb.Sheets[1];
            // working starts from here 
            //1. setting the range for deleting the heading rows
            Excel.Range range1 = sheet.get_Range("A1", "A9".ToString());
            range1.EntireRow.Delete(Excel.XlDirection.xlUp);

            //2. count all rows and delete the last one 
            var LastRow = sheet.UsedRange.Rows.Count;
            var LastCol = sheet.UsedRange.Columns.Count;
            LastRow = LastRow + sheet.UsedRange.Row - 1;
            ((Excel.Range)sheet.Rows[LastRow]).Delete(Excel.XlDirection.xlUp);

            // 3. convert the second column to number format 
            sheet.Columns[2].TextToColumns();
            sheet.Columns[3].TextToColumns();
            //sheet.Columns[2].NumberFormat = "0.00 # [$$-en-US]";
            sheet.Columns[2].NumberFormat = "+ 0.00 ₹";


            //4. when found null, copy value of cell 3 to cell 2 
            for (iRow = 1; iRow <= sheet.UsedRange.Rows.Count; iRow++)
            {
                if (sheet.Cells[iRow, 2].value == null)
                {
                    sheet.Cells[iRow, 2].value = sheet.Cells[iRow, 3].value;
                    sheet.Cells[iRow, 2].NumberFormat = "- 0.00 ₹";

                }

            }
            //5.  delete column c 
            range1 = sheet.get_Range("C1:C500");
            range1.EntireColumn.Delete(Excel.XlDirection.xlUp);

            // 6. copy text file data to list 
            copydata();

            // filter only required names 
            Filter(sheet,1, list);

            //3. Make bold the first column 
            sheet.Range["A1"].EntireColumn.Font.Bold = true;
            sheet.Range["B1"].EntireColumn.Font.Bold = true;

            // Get range of data in the worksheet
            range1 = sheet.UsedRange;            
            Borders borders = range1.Borders;
            borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Color = Color.Black;
            borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            range1.Borders.Color = Color.Black;
            range1.Select();
            sheet.UsedRange.Select();

            // sort according to column of name 
            dynamic allDataRange = sheet.UsedRange;
            allDataRange.Sort(allDataRange.Columns[1], Excel.XlSortOrder.xlAscending);

            range1 = sheet.UsedRange;
            var arr1 = (object[,])range1.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            Dictionary<object, object> dic1 = new Dictionary<object, object>();
            for (int i = 1; i <= arr1.GetLength(0); i++)
            {
                dic1.Add(arr1[i, 1], arr1[i, 2]);

            }
            var arr2 = GetObjArr(@"C:\Users\Public\excel2.xlsx");
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

            string outputpath = @"C:\Users\Public\ratlam.xlsx";
            Excel.Application excel1 = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.ActiveSheet;

            for (int i = 0; i < list1.Count; i++)
            {
                if (list1[i].ToString().Trim() == "U.P" || list1[i].ToString().Trim() == "Rajasthan" || list1[i].ToString().Trim() == "Bihar" || list1[i].ToString().Trim() == "Punjab" || list1[i].ToString().Trim() == "Odisha" || list1[i].ToString().Trim() == "Chhatisgarh" || list1[i].ToString().Trim() == "West bengal" || list1[i].ToString().Trim() == "Madhya Pradesh" || list1[i].ToString().Trim() == "Jharkhand" || list1[i].ToString().Trim() == "Maharashtra" || list1[i].ToString().Trim() == "Market" || list1[i].ToString().Trim() == "Uttarakhand"|| list1[i].ToString().Trim() =="Assam" || list1[i].ToString().Trim() == "Tripura")
                {
                    //sheet.Cells[i + 1, 2].Value = dic1[list[i]];
                  
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].EntireColumn.Font.Bold  = true;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Merge();
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Cells.Font.Size = 20;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Font.Italic = true;

                }
                else
                {
                    sheet1.Cells[i + 1, 2].Value = dic1[list1[i]];
                }

                sheet1.Cells[i + 1, 1].Value = list1[i];
            }
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


        public object[,] GetObjArr(string filename)
        {

            Excel.Application excel = null;
            excel = new Excel.Application();
            excel.Visible = true;
            Excel.Workbook wkb = null;
            wkb = Open(excel, filename);
            Excel._Worksheet sheet = wkb.Sheets[1];
            Excel.Range range1 = sheet.UsedRange;
            // Read all data from data range in the worksheet
            valueArray = (object[,])range1.get_Value(XlRangeValueDataType.xlRangeValueDefault);


            wkb.Close(true);
            excel.Quit();
            // CLEAN UP.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);



            return valueArray;
        }
            public void copydata()
        {
            string filename = "abc.txt.txt";
            string filePath = @"C:\hello\" + filename;
            using (var file = new StreamReader(filePath))
            {
                var line = string.Empty;
                while ((line = file.ReadLine()) != null)
                {
                    list.Add(Convert.ToString(line, CultureInfo.InvariantCulture));
                }
            }

        }
        public void Filter(Excel._Worksheet ws, int columnIndex, List<string> values)
        {
            int i = 1;
            // set initial position to 0 
            int currentposition = 0;
           
            for (i=1; i <= ws.UsedRange.Rows.Count; i++)
            {
                Excel.Range range = ws.Cells[i, columnIndex] as Excel.Range;
                string value = range.Value.ToString();
                if (!values.Contains(value))
                {
                    range.EntireRow.Delete();
                    // range.EntireRow.Hidden = true;
                    //range.Delete(XlDeleteShiftDirection.xlShiftUp);
                    // remember last position 
                    i = currentposition;
                }
                else
                {  
                    // remove the word which have been used from list too 
                    values.Remove(value);
                    currentposition = i;
                }
                
            }
            
            // now the list contains those word whose balances are 0 
            var item = ws.UsedRange.Rows.Count;
            for (int il=0; il < values.Count; il++)
            {
                ws.Cells[item+1, 1].Value = values[il];
                ws.Cells[item + 1, 2].NumberFormat = " 0.00 ₹";
                ws.Cells[item+1, 2].Value = "0";                
                item++;
            }
            


        }
    }
}
