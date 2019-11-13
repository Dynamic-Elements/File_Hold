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
            //Excel.Range range;
            //Excel.Range chartRange;
            //Excel.Range formatRange;

            string lexfolder = Files.AddBS(baseCurDir + "FinancialFolder");
            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(lexfolder))
                {
                    /// If it does not exist then create it. 
                    DirectoryInfo di = Directory.CreateDirectory(lexfolder);
                }

            }
            catch { }

            string lexfile = lexfolder + "FinancialSheets.xlsx";

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

            xlWorkBook.SaveAs(lexfile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
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
            /// https://duckduckgo.com/?q=c%23+ocr+scanning+documents+and+texts&t=ffab&atb=v1-1&ia=web
            /// https://www.codingame.com/playgrounds/10058/scanned-pdf-to-ocr-textsearchable-pdf-using-c
            /// https://asprise.com/royalty-free-library/c%23-sharp.net-ocr-source-code-examples-demos.html
            /// https://github.com/tesseract-ocr/tesseract
            /// https://itextpdf.com/en/products/itext-7/pdfxfa
            /// https://www.nuget.org/packages/itext7/
            /// https://ghostscript.com/download/gsdnld.html


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

            //string lcServer = "67.222.39.62";
            //string lcODBC = "PostgreSQL ANSI";
            //string lcDB = "Tb_Test";
            //string lcPort = "3306";  //  Port=" + lcPort + ";
            //string lcUser = "dynamkr0_pgtest";
            //string lcProv = "SQLOLEDB";
            //string lcPass = "fzk4pktb";

            /// (New) tb_Play
            /// tb_HelpingHand
            /// playgroup
            /// tbmaster
            /// Smartman55
            /// (new) playgroup ((US) East US)

            string lcServer = "playgroup.database.windows.net";
            string lcODBC = "OODBC Driver 17 for SQL Server";
            string lcDB = "tb_HelpingHand";
            // string lcPort = "3306";  //  Port=" + lcPort + ";
            string lcUser = "tbmaster";
            string lcProv = "SQLOLEDB";
            string lcPass = "Smartman55";

            string lcSQL = "";
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            cnn.Open();
            lcSQL = "SELECT * from tb_datahold where where Week=" + lcEOW;      // lcSQL = "SELECT * from ~public~.~tb_Residents~ LIMIT 100".Replace('~', '"');

            OdbcCommand com = new OdbcCommand(lcSQL, cnn);
            int result = com.ExecuteNonQuery();
            if (result > 0)
            {
                /// Update records
                // MessageBox.Show(result.ToString());
                lcSQL = " Update tb_datahold set NetSales=@lcNetSales, PrimSupp=@lcfPrimSupp, OthSupp=@lcfOthSupp, Bread=@lcfBread, Beverage=@lcfBev," +
                    " Produce=@lcfProd,CarbonDioxide=@lcfCarbon, FoodCost=@lcfTotFood, HostCashier=@lclHost, Cooks=@lclCook, Servers=@lclServer," +
                    " DMO=@lclDMO, Supervisor=@lclSuperv, Overtime=@lclOvertime,GeneralManager=@lclGenManager, Manager=@lclManager, Bonus=@lclBonus," +
                    " PayrollTax=@lclPayTax, Healthcare=, Retirement=, LaborCost=@lclTotLabor, Accounting=@lceAccount,Bank=@lceBank, CreditCard=@lceCC," +
                    " Fuel=@lceFuel, Legal=@lceLegal, License=@lceLicensePerm, PayrollProc=@lcePayroll, Insurance=@lceInsur,WorkersComp=@lceWorkComp," +
                    " Advertising=@lceAdvertise, Charitable=@lceCharitable, Auto=@lceAuto, CashShortage=@lceCash, Electrical=@lceElect,General=@lceGeneral," +
                    " HVAC=@lceHVAC, Lawn=@lceLawn, Painting=@lcePaint, Plumbing=@lcePlumb, Remodeling=@lceRemodel, Structural=@lceStruct," +
                    " DishMachine=@lceDishMach,Janitorial=@lceJanitorial, Office=@lceOfficeComp, Restaurant=@lceRestaurant, Uniforms=@lceUniform," +
                    " Data=@lceData, Electricity=@lceElectric,Music=@lceMusic, NaturalGas=@lceNatGas, Security=@lceSecurity, Trash=@lceTrash," +
                    " WaterSewer=@lceWaterSewer, ExpenseCost=@lceTotExpense, Mortgage=@lcoMort,LoanPayment=@lcoLoan, Association=@lcoAssoc," +
                    " PropertyTax=@lcoPropTax, AdvertisingCoop=@lcoAdvCoop, NationalAdvertise=@lcoNatAdver, LicensingFee=@lcoLicenseFee," +
                    "OverheadCost=@lcoTotOverhead where Week=@lcEOW";
            }
            else
            {
                /// Insert records
                // MessageBox.Show("Hello There, no records");
                lcSQL = " Insert into tb_datahold (Week,NetSales,PrimSupp,OthSupp,Bread,Beverage,Produce,CarbonDioxide,FoodCost,HostCashier,Cooks,Servers,DMO,Supervisor," +
                    "Overtime,GeneralManager,Manager,Bonus,PayrollTax,Healthcare,Retirement,LaborCost,Accounting,Bank,CreditCard,Fuel,Legal,License,PayrollProc," +
                    "Insurance,WorkersComp,Advertising,Charitable,Auto,CashShortage,Electrical,General,HVAC,Lawn,Painting,Plumbing,Remodeling,Structural,DishMachine," +
                    "Janitorial,Office,Restaurant,Uniforms,Data,Electricity,Music,NaturalGas,Security,Trash,WaterSewer,ExpenseCost,Mortgage,LoanPayment,Association," +
                    "PropertyTax,AdvertisingCoop,NationalAdvertise,LicensingFee,OverheadCost,IDs) " +
                    " values " +
                    " ('@lcEOW','@lcNetSales','@lcfPrimSupp','@lcfOthSupp','@lcfBread','@lcfBev','@lcfProd','@lcfCarbon','@lcfTotFood','@lclHost','@lclCook','@lclServer','@lclDMO'," +
                    "'@lclSuperv','@lclOvertime','@lclGenManager','@lclManager','@lclBonus','@lclPayTax','@lcHealth','@lcRetire','@lclTotLabor','@lceAccount','@lceBank','@lceCC'," +
                    "'@lceFuel','@lceLegal','@lceLicensePerm','@lcePayroll','@lceInsur','@lceWorkComp','@lceAdvertise','@lceCharitable','@lceAuto','@lceCash','@lceElect','@lceGeneral'," +
                    "'@lceHVAC','@lceLawn','@lcePaint','@lcePlumb','@lceRemodel','@lceStruct','@lceDishMach','@lceJanitorial','@lceOfficeComp','@lceRestaurant','@lceUniform','@lceData'," +
                    "'@lceElectric','@lceMusic','@lceNatGas','@lceSecurity','@lceTrash','@lceWaterSewer','@lceTotExpense','@lcoMort','@lcoLoan','@lcoAssoc','@lcoPropTax'," +
                    "'@lcoAdvCoop','@lcoNatAdver','@lcoLicenseFee','@lcoTotOverhead')";
            }

            OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            //// Pass values to Parameters
            cmd.Parameters.AddWithValue("@lcEOW", lcEOW);
            cmd.Parameters.AddWithValue("@lcNetSales", lcNetSales);
            cmd.Parameters.AddWithValue("@lcfPrimSupp", lcfPrimSupp);
            cmd.Parameters.AddWithValue("@lcfOthSupp", lcfOthSupp);
            cmd.Parameters.AddWithValue("@lcfBread", lcfBread);
            cmd.Parameters.AddWithValue("@lcfBev", lcfBev);
            cmd.Parameters.AddWithValue("@lcfProd", lcfProd);
            cmd.Parameters.AddWithValue("@lcfCarbon", lcfCarbon);
            cmd.Parameters.AddWithValue("@lcfTotFood", lcfTotFood);
            cmd.Parameters.AddWithValue("@lclHost", lclHost);
            cmd.Parameters.AddWithValue("@lclCook", lclCook);
            cmd.Parameters.AddWithValue("@lclServer", lclServer);
            cmd.Parameters.AddWithValue("@lclDMO", lclDMO);
            cmd.Parameters.AddWithValue("@lclSuperv", lclSuperv);
            cmd.Parameters.AddWithValue("@lclOvertime", lclOvertime);
            cmd.Parameters.AddWithValue("@lclGenManager", lclGenManager);
            cmd.Parameters.AddWithValue("@lclManager", lclManager);
            cmd.Parameters.AddWithValue("@lclBonus", lclBonus);
            cmd.Parameters.AddWithValue("@lclPayTax", lclPayTax);
            // cmd.Parameters.AddWithValue("@lcHealth",lcHealth);  //  HealthCare= 
            // cmd.Parameters.AddWithValue("@lcRetire",lcRetire);  //  Retire=
            cmd.Parameters.AddWithValue("@lclTotLabor", lclTotLabor);
            cmd.Parameters.AddWithValue("@lceAccount", lceAccount);
            cmd.Parameters.AddWithValue("@lceBank", lceBank);
            cmd.Parameters.AddWithValue("@lceCC", lceCC);
            cmd.Parameters.AddWithValue("@lceFuel", lceFuel);
            cmd.Parameters.AddWithValue("@lceLegal", lceLegal);
            cmd.Parameters.AddWithValue("@lceLicensePerm", lceLicensePerm);
            cmd.Parameters.AddWithValue("@lcePayroll", lcePayroll);
            cmd.Parameters.AddWithValue("@lceInsur", lceInsur);
            cmd.Parameters.AddWithValue("@lceWorkComp", lceWorkComp);
            cmd.Parameters.AddWithValue("@lceAdvertise", lceAdvertise);
            cmd.Parameters.AddWithValue("@lceCharitable", lceCharitable);
            cmd.Parameters.AddWithValue("@lceAuto", lceAuto);
            cmd.Parameters.AddWithValue("@lceCash", lceCash);
            cmd.Parameters.AddWithValue("@lceElect", lceElect);
            cmd.Parameters.AddWithValue("@lceGeneral", lceGeneral);
            cmd.Parameters.AddWithValue("@lceHVAC", lceHVAC);
            cmd.Parameters.AddWithValue("@lceLawn", lceLawn);
            cmd.Parameters.AddWithValue("@lcePaint", lcePaint);
            cmd.Parameters.AddWithValue("@lcePlumb", lcePlumb);
            cmd.Parameters.AddWithValue("@lceRemodel", lceRemodel);
            cmd.Parameters.AddWithValue("@lceStruct", lceStruct);
            cmd.Parameters.AddWithValue("@lceDishMach", lceDishMach);
            cmd.Parameters.AddWithValue("@lceJanitorial", lceJanitorial);
            cmd.Parameters.AddWithValue("@lceOfficeComp", lceOfficeComp);
            cmd.Parameters.AddWithValue("@lceRestaurant", lceRestaurant);
            cmd.Parameters.AddWithValue("@lceUniform", lceUniform);
            cmd.Parameters.AddWithValue("@lceData", lceData);
            cmd.Parameters.AddWithValue("@lceElectric", lceElectric);
            cmd.Parameters.AddWithValue("@lceMusic", lceMusic);
            cmd.Parameters.AddWithValue("@lceNatGas", lceNatGas);
            cmd.Parameters.AddWithValue("@lceSecurity", lceSecurity);
            cmd.Parameters.AddWithValue("@lceTrash", lceTrash);
            cmd.Parameters.AddWithValue("@lceWaterSewer", lceWaterSewer);
            cmd.Parameters.AddWithValue("@lceTotExpense", lceTotExpense);
            cmd.Parameters.AddWithValue("@lcoMort", lcoMort);
            cmd.Parameters.AddWithValue("@lcoLoan", lcoLoan);
            cmd.Parameters.AddWithValue("@lcoAssoc", lcoAssoc);
            cmd.Parameters.AddWithValue("@lcoPropTax", lcoPropTax);
            cmd.Parameters.AddWithValue("@lcoAdvCoop", lcoAdvCoop);
            cmd.Parameters.AddWithValue("@lcoNatAdver", lcoNatAdver);
            cmd.Parameters.AddWithValue("@lcoLicenseFee", lcoLicenseFee);
            cmd.Parameters.AddWithValue("@lcoTotOverhead", lcoTotOverhead);
            //  cmd.Parameters.AddWithValue("@",);

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
