﻿using System;
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
            Excel.Range range1 = sheet.UsedRange;
            var arr1 = (object[,])range1.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            var listname = "Vintage Flip";
            List<object[]> listnew = new List<object[]>();
            List<string> list1 = new List<string>(); // just for sub-headings 
            
            int firstSpaceIndex;
            String firstString;
            listnew.Add(new object[] { "VINTAGE FLIP" });
            for (int i = 10; i <= arr1.GetLength(0) - 1; i++)
            {
                if (arr1[i, 1] != null)
                {
                    int id = Convert.ToInt32(arr1[i, 2].ToString().Replace("Pcs", "").Trim());
                    if (id > 30)
                    {                       
                        listnew.Add(new object[] { arr1[i, 1].ToString().Replace(listname,"").Trim() });                       
                    }
                }
            }

            // list of names are ready
            
            // extract sub heading 
            for (int i = 1; i < listnew.Count; i++)
            {               
                    firstSpaceIndex = listnew[i][0].ToString().IndexOf(" "); // get index upto first space found
                    firstString = listnew[i][0].ToString().Substring(0, firstSpaceIndex); // get string upto first space found
                    if (!list1.Contains(firstString))
                    {
                        list1.Add(firstString); // Add to sub-heading list 
                    }
               
            }

            // List of sub-heading are  ready , now prepare final list
             Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();

            for (int i = 0; i < list1.Count; i++)
            {
                string key = list1[i];
                List<string> l = new List<string>();
                foreach (var item in listnew)
                {
                    if (item[0].ToString().Contains(key))
                    {
                        l.Add(item[0].ToString().Replace(key, ""));
                    }

                }
                dic.Add(key, l);
            }

            string outputpath = @"C:\Users\Public\VintageList.xlsx";
            Excel.Application excel1 = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.ActiveSheet;
            // create main heading 
            sheet1.Cells[1, 1].Value = listnew[0];
            // stretch it to three columns 
            //sheet1.Range[sheet1.Cells[1, 1], sheet1.Cells[1, 2], sheet1.Cells[1, 3]].EntireColumn.Font.Bold = true;
            //sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 2], sheet1.Cells[1, 3]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 2], sheet1.Cells[1, 3]].Merge();
            //sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 2], sheet1.Cells[1, 3]].Cells.Font.Size = 20;
            //sheet1.Range[sheet1.Cells[ 1, 1], sheet1.Cells[1, 2], sheet1.Cells[1, 3]].Font.Italic = true;

            // create inner cells
            for (int i = 1; i < listnew.Count; i++)
            {
                if (listnew[i].ToString().Trim() == "U.P" || listnew[i].ToString().Trim() == "Rajasthan" || listnew[i].ToString().Trim() == "Bihar" )
                {
                    //sheet.Cells[i + 1, 2].Value = dic1[list[i]];
                  
                    sheet1.Range[sheet1.Cells[i + 1, 1]].EntireColumn.Font.Bold  = true;
                    sheet1.Range[sheet1.Cells[i + 1]] = XlHAlign.xlHAlignCenter;
                    sheet1.Range[sheet1.Cells[i + 1, 1], sheet1.Cells[i + 1, 2]].Cells.Font.Size = 10;
                }
                else
                {
                    // remove below comment 
                  //  sheet1.Cells[i + 1, 2].Value = dic1[list1[i]];
                }

                sheet1.Cells[i + 1, 1].Value = list1[i];
            }
            
            sheet1.Columns["A:B:C"].AutoFit();
            sheet1.Range["A1"].EntireColumn.Font.Bold = true;
            sheet1.Range["B1"].EntireColumn.Font.Bold = true;
            sheet1.Range["C1"].EntireColumn.Font.Bold = true;
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


  
    }
}
