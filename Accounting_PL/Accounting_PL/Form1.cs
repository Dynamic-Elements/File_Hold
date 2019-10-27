using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
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

        /// <summary>
        /// Excel Code
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

            string lexfile = lexfolder + "TestExcelHolder.xlsx";

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

            xlWorkBook.SaveAs(lexfile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            //xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);

        }

        private void ReleaseObject(object obj)
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


        /// <summary>
        /// Scanner Button
        /// This will handle the scanner feature.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {

        }


        /// <summary>
        /// Food button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Labor Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Expenses Button.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Overhead Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Setup fields right now. Will add more later.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            var date = DateTime.Now;
            var lastSunday = Dates.DTOC(date.AddDays(-(int)date.DayOfWeek));  // Grabs the past Sunday for Week End

            textBox1.Text = lastSunday;

            textBox2.Text = lastSunday.Substring(lastSunday.Length - 4, 4);   // Yr.Substring(0,4);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button7_Click(object sender, EventArgs e)
        {
            string lcServer = "salt.db.elephantsql.com";
            string lcODBC = "PostgreSQL ANSI";
            string lcDB = "pffejyte";
            // string lcPort = "5432";  //  Port=" + lcPort + ";
            string lcUser = "pffejyte";
            string lcPass = "Or2m-sdyDidrOWGaXBD--8b1-itKL92b";
            string lcSQL = "";
            var dataSet1 = new DataSet();
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=SQLOLEDB;Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            lcSQL = "SELECT * FROM ~public~.~tb_inbuillding~ LIMIT 100".Replace('~', '"');

            OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            cnn.Open();
            OdbcCommand com = new OdbcCommand(lcSQL, cnn);
            OdbcDataReader reader = com.ExecuteReader();
            MessageBox.Show(vfpData.RecCount().ToString());
            


            //string lcYear = textBox2.Text;
            //string lcEOW = textBox1.Text;

            //string lcfPrimSupp = textBox84.Text;
            //string lcfOthSupp = textBox77.Text;
            //string lcfBread = textBox76.Text;
            //string lcfBev = textBox75.Text;
            //string lcfProd = textBox69.Text;
            //string lcfCarbon = textBox68.Text;

            //string lcoMort = textBox83.Text;
            //string lcoLoan = textBox82.Text;
            //string lcoAssoc = textBox81.Text;
            //string lcoPropTax = textBox80.Text;
            //string lcoAdvCoop = textBox79.Text;
            //string lcoNatAdver = textBox78.Text;
            //string lcoLicenseFee = textBox73.Text;

            //string lceAccount = textBox27.Text;
            //string lceBank = textBox26.Text;
            //string lceCC = textBox25.Text;
            //string lceFuel = textBox24.Text;
            //string lceLegal = textBox23.Text;
            //string lceLicensePerm = textBox22.Text;
            //string lcePayroll = textBox28.Text;
            //string lceInsur = textBox30.Text;
            //string lceWorkComp = textBox29.Text;
            //string lceAdvertise = textBox32.Text;
            //string lceCharitable = textBox31.Text;
            //string lceAuto = textBox21.Text;
            //string lceCash = textBox20.Text;
            //string lceElect = textBox34.Text;
            //string lceGeneral = textBox33.Text;
            //string lceHVAC = textBox19.Text;
            //string lceLawn = textBox35.Text;
            //string lcePaint = textBox36.Text;
            //string lcePlumb = textBox37.Text;
            //string lceRemodel = textBox38.Text;
            //string lceStruct = textBox39.Text;
            //string lceDishMach = textBox43.Text;
            //string lceJanitorial = textBox42.Text;
            //string lceOfficeComp = textBox44.Text;
            //string lceRestaurant = textBox41.Text;
            //string lceUniform = textBox40.Text;
            //string lceData = textBox18.Text;
            //string lceElectric = textBox45.Text;
            //string lceMusic = textBox46.Text;
            //string lceNatGas = textBox47.Text;

            //string lclHost = textBox90.Text;
            //string lclCook = textBox89.Text;
            //string lclServer = textBox88.Text;
            //string lclDMO = textBox87.Text;
            //string lclSuperv = textBox86.Text;
            //string lclOvertime = textBox85.Text;
            //string lclGenManager = textBox74.Text;
            //string lclManager = textBox72.Text;
            //string lclBonus = textBox71.Text;
            //string lclPayTax = textBox70.Text;



            // DataView goView = VfpData.SqlExecute(oConn, lcSQL, "A_List");
            // VfpData.Browse(goView);
            // VfpData.Browse(command.ExecuteScalar());

            /// oConn = SqlConnect(lcConnectionString);
            /// goView = SqlExecute(oConn, lcSQL, "CustomerList");
            /// SqlDisConnect(oConn);
            /// 
            /// //Select the default cursor and browse it
            /// VFPToolkit.vfpData.Select(goView);
            /// Browse();

            MessageBox.Show("Done!");

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string lcServer = "salt.db.elephantsql.com";
            string lcODBC = "PostgreSQL ANSI";
            string lcDB = "pffejyte";
            // string lcPort = "5432";  //  Port=" + lcPort + ";
            string lcUser = "pffejyte";
            string lcPass = "Or2m-sdyDidrOWGaXBD--8b1-itKL92b";
            string lcSQL = "";
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=SQLOLEDB;Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            OleDbConnection oConn = VfpData.SqlStringConnect(lcConnectionString);

        }
    }
}
