using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VFPToolkit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Accounting_PL
{
    public partial class Form1 : Form
    {

        string appPath = AppDomain.CurrentDomain.BaseDirectory;
        string curDir = Files.AddBS(Files.CurDir());
        // MessageBox.Show("here " + curDir);
        string baseCurDir = Files.AddBS(Path.GetFullPath(Path.Combine(Files.CurDir(), @"..\..\..\")));
        // MessageBox.Show("here " + baseCurDir);

        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;
            Excel.Range range;
            Excel.Range chartRange;
            Excel.Range formatRange;

            string lexfolder = Files.AddBS(baseCurDir + "ExcelHold");
            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(lexfolder))
                {
                    DirectoryInfo di = Directory.CreateDirectory(lexfolder);
                    // MessageBox.Show("The directory was created successfully at " + Directory.GetCreationTime(lexfolder));
                }

            }
            catch { }

            string lexfile = lexfolder + "ExcelHolder.xlsx";

            xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            // xlWorkBook = xlApp.Workbooks.Open(@"d:\csharp-Excel.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0)
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //add data 
            xlWorkSheet.Cells[4, 2] = "";
            xlWorkSheet.Cells[4, 3] = "Student1";
            xlWorkSheet.Cells[4, 4] = "Student2";
            xlWorkSheet.Cells[4, 5] = "Student3";

            xlWorkSheet.Cells[5, 2] = "Term1";
            xlWorkSheet.Cells[5, 3] = "80";
            xlWorkSheet.Cells[5, 4] = "65";
            xlWorkSheet.Cells[5, 5] = "45";

            xlWorkSheet.Cells[6, 2] = "Term2";
            xlWorkSheet.Cells[6, 3] = "78";
            xlWorkSheet.Cells[6, 4] = "72";
            xlWorkSheet.Cells[6, 5] = "60";

            xlWorkSheet.Cells[7, 2] = "Term3";
            xlWorkSheet.Cells[7, 3] = "82";
            xlWorkSheet.Cells[7, 4] = "80";
            xlWorkSheet.Cells[7, 5] = "65";

            xlWorkSheet.Cells[8, 2] = "Term4";
            xlWorkSheet.Cells[8, 3] = "75";
            xlWorkSheet.Cells[8, 4] = "82";
            xlWorkSheet.Cells[8, 5] = "68";

            xlWorkSheet.Cells[9, 2] = "Total";
            xlWorkSheet.Cells[9, 3] = "315";
            xlWorkSheet.Cells[9, 4] = "299";
            xlWorkSheet.Cells[9, 5] = "238";

            formatRange = xlWorkSheet.get_Range("a1", "b1");
            formatRange.NumberFormat = "mm/dd/yyyy";
            //formatRange.NumberFormat = "mm/dd/yyyy hh:mm:ss";
            xlWorkSheet.Cells[1, 1] = "31/5/2014";

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";

            xlApp.Visible = true;

            xlWorkBook.SaveAs(lexfile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            //xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }

        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);  //  System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            panel2.Visible = true;
            panel2.BringToFront();
            panel3.Visible = false;
            panel3.SendToBack();
            panel4.Visible = false;
            panel4.SendToBack();
            panel5.Visible = false;
            panel5.SendToBack();
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel2.SendToBack();
            panel3.Visible = false;
            panel3.SendToBack();
            panel4.Visible = false;
            panel4.SendToBack();
            panel5.Visible = true;
            panel5.BringToFront();
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel2.SendToBack();
            panel3.Visible = true;
            panel3.BringToFront();
            panel4.Visible = false;
            panel4.SendToBack();
            panel5.Visible = false;
            panel5.SendToBack();
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel2.SendToBack();
            panel3.Visible = false;
            panel3.SendToBack();
            panel4.Visible = true;
            panel4.BringToFront();
            panel5.Visible = false;
            panel5.SendToBack();
        }
    }
}
