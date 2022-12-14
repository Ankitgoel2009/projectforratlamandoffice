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
            sheet.Range["B1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            sheet.Range["C1"].EntireColumn.NumberFormat = @"[>=10000000]##\,##\,##\,##0.00;[>=100000] ##\,##\,##0.00;##,##0.00";
            Excel.Range range1 = sheet.UsedRange;
            int rowindex = range1.Rows.Count;
            int columnindex = range1.Columns.Count;
            List<object[]> listnew = new List<object[]>();
            for (int i = 10; i <= rowindex - 1; i++)
            {
                if (range1.Cells[i,1].Text != "")
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
         

            // 6. copy text file data to list 
            copydata();

            // filter from only index which conatins  names and add the balance 
            Filter(listnew, list);

            // names are  ready , now modify it 
                       
            Dictionary<object, object> dic1 = new Dictionary<object, object>();
            for (int i = 0; i <= listnew.Count-1; i++)
            {
                dic1.Add(listnew[i][0], listnew[i][1]);

            }

            // data which tells which names are in which state
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
    public void Filter(List<object[]> ws, List<string> values)
        {
            var result1 = ws.Select(m => m[0]).ToList();
            var result = result1.Except(list);
            ws.RemoveAll(m => result.Contains(m[0]));
            for (int i = 0; i < list.Count; i++)
            {
                if (!result1.Contains(list[i]))
                {
                    ws.Add(new object[] { list[i], "0 ₹"  });
                }
            }
          
            
        }
    }
}
