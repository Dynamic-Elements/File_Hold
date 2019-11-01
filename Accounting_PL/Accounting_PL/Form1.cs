using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
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
            xlWorkSheet.Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(1); 

            //add data 
            
            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";

            xlApp.Visible = true;

            //  xlWorkBook.Worksheets.Add();

            var coll = new Excel.Worksheet[13];

            for (int i = 2; i < 13; i++)
            {
                coll[i] = xlWorkBook.Worksheets.Add();
                coll[i].Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i);

                coll[i].Cells[1, 1] = "ID";
                coll[i].Cells[1, 2] = "Name";
                coll[i].Cells[2, 1] = "1";
                coll[i].Cells[2, 2] = "One";
                coll[i].Cells[4, 3] = "Student1";

            }

            xlWorkBook.Worksheets.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Name = "YTD";

            // xlWorkBook.SaveAs(lexfile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            // xlWorkBook.Close(true, misValue, misValue);
            // xlApp.Quit();
            // xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            // xlWorkBook.Close(true, misValue, misValue);
            // xlApp.Quit();

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
            /// Notes
            /// https://www.scanitto.com/
            /// https://www.vintasoft.com/download.html
            /// http://www.viscomsoft.com/


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
        /// Save button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button7_Click(object sender, EventArgs e)
        {
            //string lcServer = "salt.db.elephantsql.com";
            //string lcODBC = "PostgreSQL ANSI";
            //string lcDB = "pffejyte";
            //// string lcPort = "5432";  //  Port=" + lcPort + ";
            //string lcUser = "pffejyte";
            //string lcPass = "Or2m-sdyDidrOWGaXBD--8b1-itKL92b";
            //string lcSQL = "";
            //string lcConnectionString = "Driver={" + lcODBC + "};Provider=SQLOLEDB;Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";

            string lcYear = textBox2.Text.Trim();
            string lcEOW = textBox1.Text.Trim();
            string lcNetSales = textBox3.Text.Trim();

            string lcfPrimSupp = textBox84.Text.Trim();
            string lcfOthSupp = textBox77.Text.Trim();
            string lcfBread = textBox76.Text.Trim();
            string lcfBev = textBox75.Text.Trim();
            string lcfProd = textBox69.Text.Trim();
            string lcfCarbon = textBox68.Text.Trim();
            string lcfTotFood = textBox4.Text.Trim();

            string lcoMort = textBox83.Text.Trim();
            string lcoLoan = textBox82.Text.Trim();
            string lcoAssoc = textBox81.Text.Trim();
            string lcoPropTax = textBox80.Text.Trim();
            string lcoAdvCoop = textBox79.Text.Trim();
            string lcoNatAdver = textBox78.Text.Trim();
            string lcoLicenseFee = textBox73.Text.Trim();
            string lcoTotOverhead = textBox6.Text.Trim();

            string lceAccount = textBox27.Text.Trim();
            string lceBank = textBox26.Text.Trim();
            string lceCC = textBox25.Text.Trim();
            string lceFuel = textBox24.Text.Trim();
            string lceLegal = textBox23.Text.Trim();
            string lceLicensePerm = textBox22.Text.Trim();
            string lcePayroll = textBox28.Text.Trim();
            string lceInsur = textBox30.Text.Trim();
            string lceWorkComp = textBox29.Text.Trim();
            string lceAdvertise = textBox32.Text.Trim();
            string lceCharitable = textBox31.Text.Trim();
            string lceAuto = textBox21.Text.Trim();
            string lceCash = textBox20.Text.Trim();
            string lceElect = textBox34.Text.Trim();
            string lceGeneral = textBox33.Text.Trim();
            string lceHVAC = textBox19.Text.Trim();
            string lceLawn = textBox35.Text.Trim();
            string lcePaint = textBox36.Text.Trim();
            string lcePlumb = textBox37.Text.Trim();
            string lceRemodel = textBox38.Text.Trim();
            string lceStruct = textBox39.Text.Trim();
            string lceDishMach = textBox43.Text.Trim();
            string lceJanitorial = textBox42.Text.Trim();
            string lceOfficeComp = textBox44.Text.Trim();
            string lceRestaurant = textBox41.Text.Trim();
            string lceUniform = textBox40.Text.Trim();
            string lceData = textBox18.Text.Trim();
            string lceElectric = textBox45.Text.Trim();
            string lceMusic = textBox46.Text.Trim();
            string lceNatGas = textBox47.Text.Trim();
            string lceSecurity = textBox48.Text.Trim();
            string lceTrash = textBox49.Text.Trim();
            string lceWaterSewer = textBox50.Text.Trim();
            string lceTotExpense = textBox7.Text.Trim();

            string lclHost = textBox90.Text.Trim();
            string lclCook = textBox89.Text.Trim();
            string lclServer = textBox88.Text.Trim();
            string lclDMO = textBox87.Text.Trim();
            string lclSuperv = textBox86.Text.Trim();
            string lclOvertime = textBox85.Text.Trim();
            string lclGenManager = textBox74.Text.Trim();
            string lclManager = textBox72.Text.Trim();
            string lclBonus = textBox71.Text.Trim();
            string lclPayTax = textBox70.Text.Trim();
            string lclTotLabor = textBox5.Text.Trim();

            string lcServer = "67.222.39.62";
            string lcODBC = "PostgreSQL ANSI";
            string lcDB = "Tb_Test";
            string lcPort = "3306";  //  Port=" + lcPort + ";
            string lcUser = "dynamkr0_pgtest";
            string lcProv = "SQLOLEDB";
            string lcPass = "fzk4pktb";
            string lcSQL = "";
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";Port=" + lcPort + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            cnn.Open();
            lcSQL = "SELECT * from table where where Week=" + lcEOW;      // lcSQL = "SELECT * from ~public~.~tb_Residents~ LIMIT 100".Replace('~', '"');

            OdbcCommand com = new OdbcCommand(lcSQL, cnn);
            int result = com.ExecuteNonQuery();
            if (result > 0)
            {
                /// Update records
                // MessageBox.Show(result.ToString());
                lcSQL = " Update table set NetSales=" + lcNetSales + ", PrimSupp=" + lcfPrimSupp + ", OthSupp=" + lcfOthSupp + ", Bread=" + lcfBread + ", Bever=" + lcfBev + ", Produce=" + lcfProd + "," +
                    " CarbDio=" + lcfCarbon + ", FoodC=" + lcfTotFood + ", HostCash=" + lclHost + ", Cooks=" + lclCook + ", Servers=" + lclServer + ", DMO=" + lclDMO + ", Superv=" + lclSuperv + ", Overt=" + lclOvertime + "," +
                    " GenMan=" + lclGenManager + ", Manager=" + lclManager + ", Bonus=" + lclBonus + ", PayTax=" + lclPayTax + ", HealthCare=, Retire=, LaborC=" + lclTotLabor + ", Accounting=" + lceAccount + "," +
                    " Bank=" + lceBank + ", CreditC=" + lceCC + ", Fuel=" + lceFuel + ", Legal=" + lceLegal + ", License=" + lceLicensePerm + ", PayRollP=" + lcePayroll + ", Insurance=" + lceInsur + "," +
                    " WorkComp=" + lceWorkComp + ", Ads=" + lceAdvertise + ", Charitable=" + lceCharitable + ", Auto=" + lceAuto + ", Cash=" + lceCash + ", Electrical=" + lceElect + "," +
                    " General=" + lceGeneral + ", HVAC=" + lceHVAC + ", Lawn=" + lceLawn + ", Paint=" + lcePaint + ", Plumb=" + lcePlumb + ", Remodel=" + lceRemodel + ", DishM=" + lceDishMach + "," +
                    " Janitorial=" + lceJanitorial + ", Office=" + lceOfficeComp + ", Restaurant=" + lceRestaurant + ", Uniforms=" + lceUniform + ", Data=" + lceData + ", Electricity=" + lceElectric + "," +
                    " Music=" + lceMusic + ", NaturalG=" + lceNatGas + ", Security=" + lceSecurity + ", Trash=" + lceTrash + ", Water=" + lceWaterSewer + ", Expenses=" + lceTotExpense + ", Mortgage=" + lcoMort + "," +
                    " Loan=" + lcoLoan + ", Association=" + lcoAssoc + ", PropertyT=" + lcoPropTax + ", Advertising=" + lcoAdvCoop + ", NationalAds=" + lcoNatAdver + ", LicensingF=" + lcoLicenseFee + "," +
                    " OverheadC=" + lcoTotOverhead + ", Structural=" + lceStruct + " where Week=" + lcEOW;
            }
            else
            {
                /// Insert records
                // MessageBox.Show("Hello There, no records");
                //lcSQL = " Insert into table (Week,NetSales,PrimSupp,OthSupp,Bread,Bever,Produce,CarbDio,FoodC,HostCash,Cooks,Servers,DMO,Superv,Overt,GenMan,Manager,Bonus,PayTax,HealthCare,Retire,LaborC,Accounting," +
                //    "Bank,CreditC,Fuel,Legal,License,PayRollP,Insurance,WorkComp,Ads,Charitable,Auto,Cash,Electrical,General,HVAC,Lawn,Paint,Plumb,Remodel,DishM,Janitorial,Office,Restaurant,Uniforms,Data," +
                //    "Electricity,Music,NaturalG,Security,Trash,Water,Expenses,Mortgage,Loan,Association,PropertyT,Advertising,NationalAds,LicensingF,OverheadC,IDs,Structural) " +
                //    " values " +
                //    " ('" + lcEOW + "','" + lcNetSales + "','" + lcfPrimSupp + "','"+ lcfOthSupp + "','"+ lcfBread + "','"+ lcfBev + "','"+ lcfProd + "','"+ lcfCarbon + "','"+ lcfTotFood + "','"+ lclHost + "'," +
                //    "'"+ lclCook + "','"+ lclServer + "','"+ lclDMO + "','"+ lclSuperv + "','"+ lclOvertime + "','"+ lclGenManager + "','"+ lclManager + "','"+ lclBonus + "','"+ lclPayTax + "','"++"'," +
                //    "'"++"','" + lclTotLabor + "','" + lceAccount + "','" + lceBank + "','" + lceCC + "','" + lceFuel + "','" + lceLegal + "','" + lceLicensePerm + "','" + lcePayroll + "','" + lceInsur + "','" + lceWorkComp + "'," +
                //    "'" + lceAdvertise + "','" + lceCharitable + "','" + lceAuto + "','" + lceCash + "','" + lceElect + "','" + lceGeneral + "','" + lceHVAC + "','" + lceLawn + "','" + lcePaint + "','" + lcePlumb + "'," +
                //    "'" + lceRemodel + "','" + lceDishMach + "','" + lceJanitorial + "','" + lceOfficeComp + "','" + lceRestaurant + "','" + lceUniform + "','" + lceData + "','" + lceElectric + "','" + lceMusic + "','" + lceNatGas + "'," +
                //    "'" + lceSecurity + "','" + lceTrash + "','" + lceWaterSewer + "','" + lceTotExpense + "','" + lcoMort + "','" + lcoLoan + "','" + lcoAssoc + "','" + lcoPropTax + "','" + lcoAdvCoop + "','" + lcoNatAdver + "'," +
                //    "'" + lcoLicenseFee + "','" + lcoTotOverhead + "','" + lceStruct + "') ";
            }

            OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            int rowsAdded = cmd.ExecuteNonQuery();
            if (rowsAdded > 0)
                MessageBox.Show("Row inserted!!");
            else
                // Well this should never really happen
                MessageBox.Show("No row inserted");

            cnn.Close();
            MessageBox.Show("Done!");

        }
    }
}
