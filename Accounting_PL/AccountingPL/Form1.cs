using ScanIt;
using IronOcr;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using VFPToolkit;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Azure.Storage.Files.Shares;
using Azure;

namespace AccountingPL
{
    public partial class Form1 : Form
    {

        string appPath = AppDomain.CurrentDomain.BaseDirectory;
        string curDir = Files.AddBS(Files.CurDir());
        string baseCurDir = Files.AddBS(Path.GetFullPath(Path.Combine(Files.CurDir(), @"..\..\..\")));
        string fileCurDir = Files.AddBS(Path.GetFullPath(Path.Combine(Files.CurDir(), @"..\..\")));
        // string lcStoreName = "IHOP158-AZTEC # Manager".Trim().Substring(4, 3);
        string lcStoreName = System.Environment.MachineName.Trim();    // " DESKTOP-6HBQUIA"
        string scanFileDir = Files.AddBS(Path.GetFullPath(Path.Combine(Files.CurDir(), @"..\..\..\")) + "ScannedDocuments");
        string excelFileDir = Files.AddBS(Path.GetFullPath(Path.Combine(Files.CurDir(), @"..\..\..\")) + "FinancialFolder");


        public Form1()
        {
            InitializeComponent();
        }


        /// <summary>
        /// Setup fields right now. Will add more later.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {

            /// Will make sure to grab the store number on any computer and be used for testing.
            if (lcStoreName != "DESKTOP-6HBQUIA")
            {
                lcStoreName = System.Environment.MachineName.Trim().Substring(4, 3);
            }
            else
            {
                lcStoreName = "IHOP158-AZTEC # Manager".Trim().Substring(4, 3);
            }

            var date = DateTime.Now;
            var lastSunday = Dates.DTOC(date.AddDays(-(int)date.DayOfWeek));  // Grabs the past Sunday for Week End
            var lYear = DateTime.Now.Year.ToString();
            txtWeek.Text = lastSunday;
            txtYear.Text = lYear;
            txtInvHold.Text = "FOOD";

            txtStoreNumb.Text = lcStoreName;
            txtInvDate.Text = DateTime.Now.ToString("MM/dd/yyyy");

            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = "Data Source=dynamicelements.database.windows.net;Initial Catalog=dynamicelements;Persist Security Info=True;User ID=tbmaster;Password=Crazy$Times44";
            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = "select category from dynamicelements..tb_category order by category";
            DataTable dt = new DataTable();

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(dt);
            cbCategory.DataSource = dt; // setting the datasource property of combobox
            cbCategory.DisplayMember = "category"; // Display Member which will display on screen
            cbCategory.ValueMember = "category"; // ID Member using which you will get the selected Item ID
            conn.Close();

            //// This will create records for the new week so the system just needs to update data
            string lcSQL = " Exec dynamicelements..CheckRecord @IDs=" + lcStoreName;  // 138  158  168  180  192  197  209  218  222
            SQLCommand(lcSQL);

            refreshFormFields();

        }



        /// <summary>
        /// This should make using SQL Connections easier. 
        /// Use this for Update, Delete, Insert  -->  but no Selecting
        /// </summary>
        /// <param name="queryString"></param>
        private static void SQLCommand(string queryString)  //  , string connectionString)
        {
            string lcServer = "dynamicelements.database.windows.net";
            string lcDB = "dynamicelements";
            string lcUser = "tbmaster";
            string lcPass = "Crazy$Times44";
            string connectionString = "Data Source=" + lcServer + ";Initial Catalog=" + lcDB + ";Persist Security Info=True;User ID=" + lcUser + ";Password=" + lcPass;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
                command.Connection.Close();
            }

            //string lcServer = "dynamicelements.database.windows.net";  // playgroup.database.windows.net
            //string lcODBC = "ODBC Driver 17 for SQL Server";
            //string lcDB = "dynamicelements";
            //string lcUser = "tbmaster";
            //string lcProv = "SQLOLEDB";
            //string lcPass = "Crazy$Times44";
            //string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            //OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            //cnn.Open();
            //lcSQL = "select * from dynamicelements..vw_OrderLogs where week='" + lcEOW + "' and AddressID=" + lcStoreName;
            //OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            //OdbcDataReader reader = cmd.ExecuteReader();
        }



        public static OdbcDataReader GetData(string queryString)
        {
            // https://www.codeproject.com/Questions/679137/fill-gridview-from-datareader
            string lcServer = "dynamicelements.database.windows.net";  // playgroup.database.windows.net
            string lcODBC = "ODBC Driver 17 for SQL Server";
            string lcDB = "dynamicelements";
            string lcUser = "tbmaster";
            string lcProv = "SQLOLEDB";
            string lcPass = "Crazy$Times44";
            string connectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";

            OdbcConnection connection = new OdbcConnection(connectionString);
            connection.Open();
            OdbcCommand cmd = new OdbcCommand(queryString, connection);
            OdbcDataReader reader = cmd.ExecuteReader();
            // connection.Close();
            return reader;
        }




        /// <summary>
        /// Best to use this after updating the database with data
        /// </summary>
        private void refreshFormFields()
        {

            string lcEOW = txtWeek.Text.Trim();
            string lcSQL = "select * from dynamicelements..vw_OrderLogs where week='" + lcEOW + "' and AddressID=" + lcStoreName;
            OdbcDataReader reader = GetData(lcSQL);

            if (reader.HasRows)
            {

                txtNetSales.Text = reader["NetSales"].ToString();
                txtRetire.Text = reader["Healthcare"].ToString();
                txtHealth.Text = reader["Retirement"].ToString();

                txtPrimSup.Text = reader["PrimSupp"].ToString();
                txtOtherSupp.Text = reader["OthSupp"].ToString();
                txtBread.Text = reader["Bread"].ToString();
                txtBev.Text = reader["Beverage"].ToString();
                txtProd.Text = reader["Produce"].ToString();
                txtCarbDio.Text = reader["CarbonDioxide"].ToString();
                txtFoodTot.Text = reader["FoodCost"].ToString();

                txtMortgage.Text = reader["Mortgage"].ToString();
                txtLoan.Text = reader["LoanPayment"].ToString();
                txtAssociation.Text = reader["Association"].ToString();
                txtPropTax.Text = reader["PropertyTax"].ToString();
                txtAdvCoop.Text = reader["AdvertisingCoop"].ToString();
                txtNationalAdv.Text = reader["NationalAdvertise"].ToString();
                txtLicenseFee.Text = reader["LicensingFee"].ToString();
                txtTotOverhead.Text = reader["OverheadCost"].ToString();

                txtAccount.Text = reader["Accounting"].ToString();
                txtBank.Text = reader["Bank"].ToString();
                txtCC.Text = reader["CreditCard"].ToString();
                txtFuel.Text = reader["Fuel"].ToString();
                txtLegal.Text = reader["Legal"].ToString();
                txtLicense.Text = reader["License"].ToString();
                txtPayroll.Text = reader["PayrollProc"].ToString();
                txtInsur.Text = reader["Insurance"].ToString();
                txtWorkComp.Text = reader["WorkersComp"].ToString();
                txtAdvertising.Text = reader["Advertising"].ToString();
                txtCharitableComp.Text = reader["Charitable"].ToString();
                txtAuto.Text = reader["Auto"].ToString();
                txtCashShort.Text = reader["CashShortage"].ToString();
                txtElectrical.Text = reader["Electrical"].ToString();
                txtGeneral.Text = reader["General"].ToString();
                txtHVAC.Text = reader["HVAC"].ToString();
                txtLawn.Text = reader["Lawn"].ToString();
                txtPaint.Text = reader["Painting"].ToString();
                txtPlumb.Text = reader["Plumbing"].ToString();
                txtRemodel.Text = reader["Remodeling"].ToString();
                txtStructural.Text = reader["Structural"].ToString();
                txtDishMach.Text = reader["DishMachine"].ToString();
                txtJanitorial.Text = reader["Janitorial"].ToString();
                txtOffice.Text = reader["Office"].ToString();
                txtRestaurant.Text = reader["Restaurant"].ToString();
                txtUniform.Text = reader["Uniforms"].ToString();
                txtDataTele.Text = reader["Data"].ToString();
                txtElectricity.Text = reader["Electricity"].ToString();
                txtMusic.Text = reader["Music"].ToString();
                txtNatGas.Text = reader["NaturalGas"].ToString();
                txtSecurity.Text = reader["Security"].ToString();
                txtTrash.Text = reader["Trash"].ToString();
                txtWater.Text = reader["WaterSewer"].ToString();
                txtTotExpense.Text = reader["ExpenseCost"].ToString();

                txtHost.Text = reader["HostCashier"].ToString();
                txtCooks.Text = reader["Cooks"].ToString();
                txtServers.Text = reader["Servers"].ToString();
                txtDMO.Text = reader["DMO"].ToString();
                txtSupervisor.Text = reader["Supervisor"].ToString();
                txtOvertime.Text = reader["Overtime"].ToString();
                txtGenManager.Text = reader["GeneralManager"].ToString();
                txtManager.Text = reader["Manager"].ToString();
                txtBonus.Text = reader["Bonus"].ToString();
                txtPayrollTax.Text = reader["PayrollTax"].ToString();
                txtTotLabor.Text = reader["LaborCost"].ToString();

            }
            else
            {
                txtNetSales.Text = "0.00";
                txtRetire.Text = "0.00";
                txtHealth.Text = "0.00";

                txtPrimSup.Text = "0.00";
                txtOtherSupp.Text = "0.00";
                txtBread.Text = "0.00";
                txtBev.Text = "0.00";
                txtProd.Text = "0.00";
                txtCarbDio.Text = "0.00";
                txtFoodTot.Text = "0.00";

                txtMortgage.Text = "0.00";
                txtLoan.Text = "0.00";
                txtAssociation.Text = "0.00";
                txtPropTax.Text = "0.00";
                txtAdvCoop.Text = "0.00";
                txtNationalAdv.Text = "0.00";
                txtLicenseFee.Text = "0.00";
                txtTotOverhead.Text = "0.00";

                txtAccount.Text = "0.00";
                txtBank.Text = "0.00";
                txtCC.Text = "0.00";
                txtFuel.Text = "0.00";
                txtLegal.Text = "0.00";
                txtLicense.Text = "0.00";
                txtPayroll.Text = "0.00";
                txtInsur.Text = "0.00";
                txtWorkComp.Text = "0.00";
                txtAdvertising.Text = "0.00";
                txtCharitableComp.Text = "0.00";
                txtAuto.Text = "0.00";
                txtCashShort.Text = "0.00";
                txtElectrical.Text = "0.00";
                txtGeneral.Text = "0.00";
                txtHVAC.Text = "0.00";
                txtLawn.Text = "0.00";
                txtPaint.Text = "0.00";
                txtPlumb.Text = "0.00";
                txtRemodel.Text = "0.00";
                txtStructural.Text = "0.00";
                txtDishMach.Text = "0.00";
                txtJanitorial.Text = "0.00";
                txtOffice.Text = "0.00";
                txtRestaurant.Text = "0.00";
                txtUniform.Text = "0.00";
                txtDataTele.Text = "0.00";
                txtElectricity.Text = "0.00";
                txtMusic.Text = "0.00";
                txtNatGas.Text = "0.00";
                txtSecurity.Text = "0.00";
                txtTrash.Text = "0.00";
                txtWater.Text = "0.00";
                txtTotExpense.Text = "0.00";

                txtHost.Text = "0.00";
                txtCooks.Text = "0.00";
                txtServers.Text = "0.00";
                txtDMO.Text = "0.00";
                txtSupervisor.Text = "0.00";
                txtOvertime.Text = "0.00";
                txtGenManager.Text = "0.00";
                txtManager.Text = "0.00";
                txtBonus.Text = "0.00";
                txtPayrollTax.Text = "0.00";
                txtTotLabor.Text = "0.00";
            }
            reader.Close();
            updateFormat();
        }



        /// <summary>
        /// Excel Code
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {

            updateCalculations();

            //Excel.Application xlApp;
            //Excel.Workbook xlWorkBook;
            //Excel.Worksheet xlWorkSheet;
            //object misValue = Missing.Value;

            string lexfolder = excelFileDir;  //  Files.AddBS(baseCurDir + "FinancialFolder");
            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(lexfolder))
                {
                    /// If it does not exist then create it. 
                    Directory.CreateDirectory(lexfolder);
                }
            }
            catch { }

            string lexfile = lexfolder + "FinancialSheets.xlsx";
            string lcYear = txtYear.ToString();
            string lexeApp = fileCurDir + "createexcel.EXE";
            string lvar = lcYear + "_" + lcStoreName + "_" + lexfolder;

            File.WriteAllText(curDir + "VarforVfp.txt", lvar);
            Process.Start(lexeApp);


            //string lcServer = "dynamicelements.database.windows.net";  // playgroup.database.windows.net
            //string lcODBC = "ODBC Driver 17 for SQL Server";
            //string lcDB = "dynamicelements";
            //string lcUser = "tbmaster";
            //string lcProv = "SQLOLEDB";
            //string lcPass = "Crazy$Times44";     // Smartman55  Crazy$Times44
            //string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            //OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            //cnn.Open();

            //// string lcSQL = "SELECT * from dynamicelements..vw_OrderLogs where year='" + lcYear + "' order by week";
            //string lcSQL = "EXEC dynamicelements..MakePNL @year = '" + lcYear + "', @AddressID='" + lcStoreName + "'";
            //OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            //OdbcDataAdapter adapter = new OdbcDataAdapter(lcSQL, cnn);
            //OdbcDataReader reader = cmd.ExecuteReader();

            //DataSet ds = new DataSet();
            //adapter.Fill(ds);

            //System.Data.DataTable dtable = new System.Data.DataTable();
            //dtable = (ds.Tables[0]);

            //cnn.Close();

            //xlApp = new Excel.Application();
            //if (xlApp == null)
            //{
            //    MessageBox.Show("Excel is not properly installed!!");
            //    return;
            //}

            //xlApp.DisplayAlerts = false;
            //xlWorkBook = xlApp.Workbooks.Add(misValue);
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //// xlWorkSheet.Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(1);
            ////  xlWorkBook.Worksheets.Add();

            ////iWeeksPerMonth = 4  sMonth = "January"
            ////iWeeksPerMonth = 4  sMonth = "February"
            ////iWeeksPerMonth = 5  sMonth = "March"
            ////iWeeksPerMonth = 4  sMonth = "April"
            ////iWeeksPerMonth = 4  sMonth = "May"
            ////iWeeksPerMonth = 5  sMonth = "June"
            ////iWeeksPerMonth = 4  sMonth = "July"
            ////iWeeksPerMonth = 4  sMonth = "August"
            ////iWeeksPerMonth = 5  sMonth = "September"
            ////iWeeksPerMonth = 4  sMonth = "October"
            ////iWeeksPerMonth = 4  sMonth = "November"
            ////iWeeksPerMonth = 5 or iWeeksPerMonth = 6  sMonth = "December"


            //var coll = new Excel.Worksheet[14];

            //for (int i = 1; i < 14; i++)
            //{
            //    coll[i] = xlWorkBook.Worksheets.Add();
            //    coll[i].Name = (i == 13) ? "YTD" : CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i);

            //    coll[i].Cells[1, 1] = "Miami Springs - ###";
            //    coll[i].Cells[1, 1].Font.Bold = true;

            //    coll[i].Cells[1, 2] = "Dates";
            //    coll[i].Range["B1:C1"].Merge();
            //    coll[i].Cells[2, 2] = "Week 1";
            //    coll[i].Cells[3, 2] = "$";
            //    coll[i].Cells[3, 3] = "%";

            //    // coll[i].Columns[B:B];

            //    coll[i].Cells[1, 4] = "Dates";
            //    coll[i].Range["d1:e1"].Merge();
            //    coll[i].Cells[2, 4] = "Week 2";
            //    coll[i].Cells[3, 4] = "$";
            //    coll[i].Cells[3, 5] = "%";

            //    coll[i].Cells[1, 6] = "Dates";
            //    coll[i].Range["f1:g1"].Merge();
            //    coll[i].Cells[2, 6] = "Week 3";
            //    coll[i].Cells[3, 6] = "$";
            //    coll[i].Cells[3, 7] = "%";

            //    coll[i].Cells[1, 8] = "Dates";
            //    coll[i].Range["h1:i1"].Merge();
            //    coll[i].Cells[2, 8] = "Week 4";
            //    coll[i].Cells[3, 8] = "$";
            //    coll[i].Cells[3, 9] = "%";

            //    if (i == 3 || i == 6 || i == 9 || i == 12)  // Extra week
            //    {

            //        coll[i].Cells[1, 10] = "Dates";
            //        coll[i].Range["j1:k1"].Merge();
            //        coll[i].Cells[2, 10] = "Week 5";
            //        coll[i].Cells[3, 10] = "$";
            //        coll[i].Cells[3, 11] = "%";

            //    }

            //    if (checkBox3.Checked == true && i == 12)  // Extra week
            //    {

            //        coll[i].Cells[1, 12] = "Dates";
            //        coll[i].Range["j1:k1"].Merge();
            //        coll[i].Cells[2, 12] = "Week 6";
            //        coll[i].Cells[3, 12] = "$";
            //        coll[i].Cells[3, 13] = "%";

            //    }

            //    coll[i].Cells[4, 1] = "Net Sales";
            //    coll[i].Cells[4, 1].Font.Bold = true;
            //    coll[i].Cells[5, 1] = "Primary Supplier";
            //    coll[i].Cells[6, 1] = "Other Suppliers";
            //    coll[i].Cells[7, 1] = "Bread";
            //    coll[i].Cells[8, 1] = "Produce";
            //    coll[i].Cells[9, 1] = "Carbon Dioxide";
            //    coll[i].Cells[10, 1] = "Food Cost";
            //    coll[i].Cells[10, 1].Font.Bold = true;
            //    coll[i].Cells[11, 1] = "Craft labor";
            //    coll[i].Cells[12, 1] = "Host/Cashier";
            //    coll[i].Cells[13, 1] = "Cooks";
            //    coll[i].Cells[14, 1] = "Servers";
            //    coll[i].Cells[15, 1] = "DMO";
            //    coll[i].Cells[16, 1] = "Supervisors";
            //    coll[i].Cells[17, 1] = "Overtime";
            //    coll[i].Cells[18, 1] = "Management";
            //    coll[i].Cells[19, 1] = "General Manager";
            //    coll[i].Cells[20, 1] = "Manager";
            //    coll[i].Cells[21, 1] = "Bonuses";
            //    coll[i].Cells[22, 1] = "Labor Expenses";
            //    coll[i].Cells[23, 1] = "Payroll Taxes";
            //    coll[i].Cells[24, 1] = "Labor Cost";
            //    coll[i].Cells[24, 1].Font.Bold = true;
            //    coll[i].Cells[25, 1] = "Fees";
            //    coll[i].Cells[26, 1] = "Accounting";
            //    coll[i].Cells[27, 1] = "Bank";
            //    coll[i].Cells[28, 1] = "Credit Card";
            //    coll[i].Cells[29, 1] = "Fuel/Delivery";
            //    coll[i].Cells[30, 1] = "Legal";
            //    coll[i].Cells[31, 1] = "Licenses/Permits";
            //    coll[i].Cells[32, 1] = "Payroll Processing";
            //    coll[i].Cells[33, 1] = "Insurance";
            //    coll[i].Cells[34, 1] = "Insurance";
            //    coll[i].Cells[35, 1] = "Workers Compensation";
            //    coll[i].Cells[36, 1] = "Local Marketing";
            //    coll[i].Cells[37, 1] = "Advertising";
            //    coll[i].Cells[38, 1] = "Charitable Contributions";
            //    coll[i].Cells[39, 1] = "Other";
            //    coll[i].Cells[40, 1] = "Auto/Travel";
            //    coll[i].Cells[41, 1] = "Cash Shortages";
            //    coll[i].Cells[42, 1] = "Repair/Matinenace";
            //    coll[i].Cells[43, 1] = "Eletrical";
            //    coll[i].Cells[44, 1] = "General";
            //    coll[i].Cells[45, 1] = "HVAC";
            //    coll[i].Cells[46, 1] = "Lawn/Parking";
            //    coll[i].Cells[47, 1] = "Painting";
            //    coll[i].Cells[48, 1] = "Plumbing";
            //    coll[i].Cells[49, 1] = "Remodeling";
            //    coll[i].Cells[50, 1] = "Structural";
            //    coll[i].Cells[51, 1] = "Supplies";
            //    coll[i].Cells[52, 1] = "Dish Machine";
            //    coll[i].Cells[53, 1] = "Janitorial";
            //    coll[i].Cells[54, 1] = "Office/Computer";
            //    coll[i].Cells[55, 1] = "Restuarant";
            //    coll[i].Cells[56, 1] = "Uniforms";
            //    coll[i].Cells[57, 1] = "Utilities";
            //    coll[i].Cells[58, 1] = "Data/Telephone";
            //    coll[i].Cells[59, 1] = "Electricity";
            //    coll[i].Cells[60, 1] = "Music";
            //    coll[i].Cells[61, 1] = "Natural Gas";
            //    coll[i].Cells[62, 1] = "Security";
            //    coll[i].Cells[63, 1] = "Trash";
            //    coll[i].Cells[64, 1] = "Water & Sewer";
            //    coll[i].Cells[65, 1] = "Expenses Cost";
            //    coll[i].Cells[65, 1].Font.Bold = true;
            //    coll[i].Cells[66, 1] = "Overhead";
            //    coll[i].Cells[67, 1] = "Mortgage/Rent";
            //    coll[i].Cells[68, 1] = "Loan Payments";
            //    coll[i].Cells[69, 1] = "Association/CAM Fees";
            //    coll[i].Cells[70, 1] = "Property Taxes";
            //    coll[i].Cells[71, 1] = "Advertising Coop";
            //    coll[i].Cells[72, 1] = "National Advertising";
            //    coll[i].Cells[73, 1] = "Licensing Fee";
            //    coll[i].Cells[74, 1] = "Overhead Cost";
            //    coll[i].Cells[74, 1].Font.Bold = true;
            //    coll[i].Cells[76, 1] = "Total Cost";
            //    coll[i].Cells[76, 1].Font.Bold = true;
            //    coll[i].Cells[78, 1] = "Return on Revenue";
            //    coll[i].Cells[78, 1].Font.Bold = true;

            //    coll[i].Columns.AutoFit();
            //    coll[i].Rows.AutoFit();
            //}

            //xlWorkBook.Sheets["Sheet1"].Delete();
            //xlApp.Visible = true;

            //xlWorkBook.SaveAs(lexfile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //// xlWorkBook.Close(true, misValue, misValue);
            //// xlApp.Quit();
            //// xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            //// xlWorkBook.Close(true, misValue, misValue);
            //// xlApp.Quit();

            //ReleaseObject(xlWorkSheet);
            //ReleaseObject(xlWorkBook);
            //ReleaseObject(xlApp);

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
        /// This is the proper code to upload files to AZURE File Cloud Storage.
        /// Make sure the filename and file extension is passed to the code.
        /// </summary>
        /// <param name="fileName"></param>
        private async void UploadFile(string fileName)
        {

            string connectionString = "DefaultEndpointsProtocol=https;AccountName=dynamicfiles;AccountKey=LWb7Fskhw1UOHuKqoMnZMSmbJmVjSz7A77YHVO1ADsKLVKOKQ3vxl0CIvrUQX0buapEu4Oc83ElN/q5lY5M9qQ==;EndpointSuffix=core.windows.net";

            // Name of the share, directory, and file we'll create
            string shareName = "restaurantdocs";
            string dirName = "store" + lcStoreName;
            // string fileName = "RestDataHold.pdf";

            // Path to the local file to upload
            string localFilePath = scanFileDir + fileName;  //  Files.AddBS(baseCurDir + "Scanned_Documents") + fileName;

            // Get a reference to a share and then create it
            ShareClient share = new ShareClient(connectionString, shareName);
            share.CreateIfNotExists();

            // Get a reference to a directory and create it
            ShareDirectoryClient directory = share.GetDirectoryClient(dirName);
            directory.CreateIfNotExists();

            // Get a reference to a file and upload it
            ShareFileClient file = directory.GetFileClient(fileName);
            using (FileStream stream = File.OpenRead(localFilePath))
            {
                file.Create(stream.Length);
                file.UploadRange(new HttpRange(0, stream.Length), stream);
            }
            //  MessageBox.Show("DONE!");
        }




        /// <summary>
        /// Scanner Button
        /// This will handle the scanner feature.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button2_Click(object sender, EventArgs e)
        {

            updateCalculations();

            string lscfolder = scanFileDir;  // Files.AddBS(baseCurDir + "Scanned_Documents");
            try
            {
                // Determine whether the directory exists.
                if (!Directory.Exists(lscfolder))
                {
                    /// If it does not exist then create it. 
                    Directory.CreateDirectory(lscfolder);
                }
            }
            catch { }

            bool adf = false;  // checkBox1
            bool duplex = false;  // checkBox2
            if (checkBox1.Checked)
                adf = true;

            if (checkBox2.Checked)
                duplex = true;

            var path = lscfolder;
            int dpi = 720;  // 150  300  600  720  1200  1270  1440
            WiaWrapper obj = new WiaWrapper();
            obj.SelectScanner();    // Need to look into how to stop scan when no scanner attached. 
            obj.Scan(true, dpi, path, adf, duplex);  //  Scan(bool rotatePage, int DPI, string filepath, bool useAdf, bool duplex)

            FileInfo oldnewestFile = GetNewestFile(new DirectoryInfo(path));
            string value = "Document Name";
            string destinaton = "";
            if (InputBox("New document", "New document name:", ref value) == DialogResult.OK)
            {
                Name = oldnewestFile.DirectoryName + "\\" + value + ".jpeg";
                destinaton = oldnewestFile.DirectoryName + "\\" + value + ".pdf";
            }
            File.Move(oldnewestFile.FullName, Name);

            PdfDocument doc = new PdfDocument();
            doc.Pages.Add(new PdfPage());
            XGraphics xgr = XGraphics.FromPdfPage(doc.Pages[0]);
            XImage img = XImage.FromFile(Name);
            xgr.DrawImage(img, 0, 0);
            doc.Save(destinaton);
            doc.Close();

            var Ocr = new IronOcr.AdvancedOcr()
            {
                CleanBackgroundNoise = true,
                EnhanceContrast = true,
                EnhanceResolution = true,
                Language = IronOcr.Languages.English.OcrLanguagePack,
                Strategy = IronOcr.AdvancedOcr.OcrStrategy.Advanced,
                ColorSpace = AdvancedOcr.OcrColorSpace.GrayScale,
                DetectWhiteTextOnDarkBackgrounds = true,
                InputImageType = AdvancedOcr.InputTypes.Document,
                RotateAndStraighten = true,
                ReadBarCodes = false,
                ColorDepth = 4
            };
            var ResultsPDF = Ocr.ReadPdf(destinaton);
            var TextPDF = ResultsPDF.Text;
            File.WriteAllText(oldnewestFile.DirectoryName + "\\testingPDF.txt", TextPDF);

            var ResultsJPG = Ocr.Read(Name);
            var TextJPG = ResultsJPG.Text;
            File.WriteAllText(oldnewestFile.DirectoryName + "\\testingJPG.txt", TextJPG);  // Looks better

            string lcPDF = Files.JustFName(destinaton);

            UploadFile(lcPDF);

            MessageBox.Show("Finished!");

        }



        /// <summary>
        /// This will change any textbox into a proper formatting for currency. 
        /// </summary>
        /// <param name="txtfile"></param>
        /// <returns></returns>
        private static string currencyFormat(string txtfile)
        {
            decimal val;
            string txtinfo = "";
            string lv = txtfile.Replace(",", "").Replace("$", "");
            if (decimal.TryParse(lv, out val))
                txtinfo = val.ToString("C");

            return txtinfo.ToString();
        }



        /// <summary>
        /// This will make all the fields have correct formatting
        /// </summary>
        private void updateFormat()
        {

            // Left side of the form
            txtNetSales.Text = currencyFormat(txtNetSales.Text);
            txtRetire.Text = currencyFormat(txtRetire.Text);
            txtHealth.Text = currencyFormat(txtHealth.Text);


            // Food
            txtPrimSup.Text = currencyFormat(txtPrimSup.Text);
            txtOtherSupp.Text = currencyFormat(txtOtherSupp.Text);
            txtBread.Text = currencyFormat(txtBread.Text);
            txtBev.Text = currencyFormat(txtBev.Text);
            txtProd.Text = currencyFormat(txtProd.Text);
            txtCarbDio.Text = currencyFormat(txtCarbDio.Text);
            txtFoodTot.Text = currencyFormat(txtFoodTot.Text);


            //// Expenses
            txtAccount.Text = currencyFormat(txtAccount.Text);
            txtBank.Text = currencyFormat(txtBank.Text);
            txtCC.Text = currencyFormat(txtCC.Text);
            txtFuel.Text = currencyFormat(txtFuel.Text);
            txtLegal.Text = currencyFormat(txtLegal.Text);
            txtLicense.Text = currencyFormat(txtLicense.Text);
            txtPayroll.Text = currencyFormat(txtPayroll.Text);
            txtInsur.Text = currencyFormat(txtInsur.Text);
            txtWorkComp.Text = currencyFormat(txtWorkComp.Text);
            txtAdvertising.Text = currencyFormat(txtAdvertising.Text);
            txtCharitableComp.Text = currencyFormat(txtCharitableComp.Text);
            txtAuto.Text = currencyFormat(txtAuto.Text);
            txtCashShort.Text = currencyFormat(txtCashShort.Text);
            txtElectrical.Text = currencyFormat(txtElectrical.Text);
            txtGeneral.Text = currencyFormat(txtGeneral.Text);
            txtHVAC.Text = currencyFormat(txtHVAC.Text);
            txtLawn.Text = currencyFormat(txtLawn.Text);
            txtPaint.Text = currencyFormat(txtPaint.Text);
            txtPlumb.Text = currencyFormat(txtPlumb.Text);
            txtRemodel.Text = currencyFormat(txtRemodel.Text);
            txtStructural.Text = currencyFormat(txtStructural.Text);
            txtDishMach.Text = currencyFormat(txtDishMach.Text);
            txtJanitorial.Text = currencyFormat(txtJanitorial.Text);
            txtOffice.Text = currencyFormat(txtOffice.Text);
            txtRestaurant.Text = currencyFormat(txtRestaurant.Text);
            txtUniform.Text = currencyFormat(txtUniform.Text);
            txtDataTele.Text = currencyFormat(txtDataTele.Text);
            txtElectricity.Text = currencyFormat(txtElectricity.Text);
            txtMusic.Text = currencyFormat(txtMusic.Text);
            txtNatGas.Text = currencyFormat(txtNatGas.Text);
            txtSecurity.Text = currencyFormat(txtSecurity.Text);
            txtTrash.Text = currencyFormat(txtTrash.Text);
            txtWater.Text = currencyFormat(txtWater.Text);
            txtTotExpense.Text = currencyFormat(txtTotExpense.Text);


            //// Labor
            txtHost.Text = currencyFormat(txtHost.Text);
            txtCooks.Text = currencyFormat(txtCooks.Text);
            txtServers.Text = currencyFormat(txtServers.Text);
            txtDMO.Text = currencyFormat(txtDMO.Text);
            txtSupervisor.Text = currencyFormat(txtSupervisor.Text);
            txtOvertime.Text = currencyFormat(txtOvertime.Text);
            txtGenManager.Text = currencyFormat(txtGenManager.Text);
            txtManager.Text = currencyFormat(txtManager.Text);
            txtBonus.Text = currencyFormat(txtBonus.Text);
            txtPayrollTax.Text = currencyFormat(txtPayrollTax.Text);
            txtTotLabor.Text = currencyFormat(txtTotLabor.Text);


            //// Overhead
            txtMortgage.Text = currencyFormat(txtMortgage.Text);
            txtLoan.Text = currencyFormat(txtLoan.Text);
            txtAssociation.Text = currencyFormat(txtAssociation.Text);
            txtPropTax.Text = currencyFormat(txtPropTax.Text);
            txtAdvCoop.Text = currencyFormat(txtAdvCoop.Text);
            txtNationalAdv.Text = currencyFormat(txtNationalAdv.Text);
            txtLicenseFee.Text = currencyFormat(txtLicenseFee.Text);
            txtTotOverhead.Text = currencyFormat(txtTotOverhead.Text);
        }



        /// <summary>
        /// 
        /// </summary>
        private void updateCalculations()
        {
            // This will calculate all the totals of each grouping

            try  //  string txt = textBox.Text.Replace(",", "").Replace("$", "");  Convert.ToDecimal()
            {
                // Food
                decimal totalamtFood = 0m;
                string txt84 = txtPrimSup.Text.Replace(",", "").Replace("$", "");
                string txt77 = txtOtherSupp.Text.Replace(",", "").Replace("$", "");
                string txt76 = txtBread.Text.Replace(",", "").Replace("$", "");
                string txt75 = txtBev.Text.Replace(",", "").Replace("$", "");
                string txt69 = txtProd.Text.Replace(",", "").Replace("$", "");
                string txt68 = txtCarbDio.Text.Replace(",", "").Replace("$", "");

                totalamtFood = Convert.ToDecimal(txt84) + Convert.ToDecimal(txt77) + Convert.ToDecimal(txt76) +
                   Convert.ToDecimal(txt75) + Convert.ToDecimal(txt69) + Convert.ToDecimal(txt68);

                txtFoodTot.Text = totalamtFood.ToString("C");


                // Expenses
                decimal totalamtExpenses = 0m;
                string txt27 = txtAccount.Text.Replace(",", "").Replace("$", "");
                string txt26 = txtBank.Text.Replace(",", "").Replace("$", "");
                string txt25 = txtCC.Text.Replace(",", "").Replace("$", "");
                string txt24 = txtFuel.Text.Replace(",", "").Replace("$", "");
                string txt23 = txtLegal.Text.Replace(",", "").Replace("$", "");
                string txt22 = txtLicense.Text.Replace(",", "").Replace("$", "");
                string txt28 = txtPayroll.Text.Replace(",", "").Replace("$", "");
                string txt30 = txtInsur.Text.Replace(",", "").Replace("$", "");
                string txt29 = txtWorkComp.Text.Replace(",", "").Replace("$", "");
                string txt32 = txtAdvertising.Text.Replace(",", "").Replace("$", "");
                string txt31 = txtCharitableComp.Text.Replace(",", "").Replace("$", "");
                string txt21 = txtAuto.Text.Replace(",", "").Replace("$", "");
                string txt20 = txtCashShort.Text.Replace(",", "").Replace("$", "");
                string txt34 = txtElectrical.Text.Replace(",", "").Replace("$", "");
                string txt33 = txtGeneral.Text.Replace(",", "").Replace("$", "");
                string txt19 = txtHVAC.Text.Replace(",", "").Replace("$", "");
                string txt35 = txtLawn.Text.Replace(",", "").Replace("$", "");
                string txt36 = txtPaint.Text.Replace(",", "").Replace("$", "");
                string txt37 = txtPlumb.Text.Replace(",", "").Replace("$", "");
                string txt38 = txtRemodel.Text.Replace(",", "").Replace("$", "");
                string txt39 = txtStructural.Text.Replace(",", "").Replace("$", "");
                string txt43 = txtDishMach.Text.Replace(",", "").Replace("$", "");
                string txt42 = txtJanitorial.Text.Replace(",", "").Replace("$", "");
                string txt44 = txtOffice.Text.Replace(",", "").Replace("$", "");
                string txt41 = txtRestaurant.Text.Replace(",", "").Replace("$", "");
                string txt40 = txtUniform.Text.Replace(",", "").Replace("$", "");
                string txt18 = txtDataTele.Text.Replace(",", "").Replace("$", "");
                string txt45 = txtElectricity.Text.Replace(",", "").Replace("$", "");
                string txt46 = txtMusic.Text.Replace(",", "").Replace("$", "");
                string txt47 = txtNatGas.Text.Replace(",", "").Replace("$", "");
                string txt48 = txtSecurity.Text.Replace(",", "").Replace("$", "");
                string txt49 = txtTrash.Text.Replace(",", "").Replace("$", "");
                string txt50 = txtWater.Text.Replace(",", "").Replace("$", "");

                totalamtExpenses = Convert.ToDecimal(txt27) + Convert.ToDecimal(txt26) + Convert.ToDecimal(txt25) + Convert.ToDecimal(txt24) + Convert.ToDecimal(txt23) +
                    Convert.ToDecimal(txt22) + Convert.ToDecimal(txt28) + Convert.ToDecimal(txt30) + Convert.ToDecimal(txt29) + Convert.ToDecimal(txt32) +
                    Convert.ToDecimal(txt31) + Convert.ToDecimal(txt21) + Convert.ToDecimal(txt20) + Convert.ToDecimal(txt34) + Convert.ToDecimal(txt33) +
                    Convert.ToDecimal(txt19) + Convert.ToDecimal(txt35) + Convert.ToDecimal(txt36) + Convert.ToDecimal(txt37) + Convert.ToDecimal(txt38) +
                    Convert.ToDecimal(txt39) + Convert.ToDecimal(txt43) + Convert.ToDecimal(txt42) + Convert.ToDecimal(txt44) + Convert.ToDecimal(txt41) +
                    Convert.ToDecimal(txt40) + Convert.ToDecimal(txt18) + Convert.ToDecimal(txt45) + Convert.ToDecimal(txt46) + Convert.ToDecimal(txt47) +
                    Convert.ToDecimal(txt48) + Convert.ToDecimal(txt49) + Convert.ToDecimal(txt50);

                txtTotExpense.Text = totalamtExpenses.ToString("C");


                // Labor
                decimal totalamtLabor = 0m;
                string txt90 = txtHost.Text.Replace(",", "").Replace("$", "");
                string txt89 = txtCooks.Text.Replace(",", "").Replace("$", "");
                string txt88 = txtServers.Text.Replace(",", "").Replace("$", "");
                string txt87 = txtDMO.Text.Replace(",", "").Replace("$", "");
                string txt86 = txtSupervisor.Text.Replace(",", "").Replace("$", "");
                string txt85 = txtOvertime.Text.Replace(",", "").Replace("$", "");
                string txt74 = txtGenManager.Text.Replace(",", "").Replace("$", "");
                string txt72 = txtManager.Text.Replace(",", "").Replace("$", "");
                string txt71 = txtBonus.Text.Replace(",", "").Replace("$", "");
                string txt70 = txtPayrollTax.Text.Replace(",", "").Replace("$", "");

                totalamtLabor = Convert.ToDecimal(txt90) + Convert.ToDecimal(txt89) + Convert.ToDecimal(txt88) + Convert.ToDecimal(txt87) +
                    Convert.ToDecimal(txt86) + Convert.ToDecimal(txt85) + Convert.ToDecimal(txt74) + Convert.ToDecimal(txt72) +
                    Convert.ToDecimal(txt71) + Convert.ToDecimal(txt70);

                txtTotLabor.Text = totalamtLabor.ToString("C");


                // Overhead
                decimal totalamtOverhead = 0m;
                string txt83 = txtMortgage.Text.Replace(",", "").Replace("$", "");
                string txt82 = txtLoan.Text.Replace(",", "").Replace("$", "");
                string txt81 = txtAssociation.Text.Replace(",", "").Replace("$", "");
                string txt80 = txtPropTax.Text.Replace(",", "").Replace("$", "");
                string txt79 = txtAdvCoop.Text.Replace(",", "").Replace("$", "");
                string txt78 = txtNationalAdv.Text.Replace(",", "").Replace("$", "");
                string txt73 = txtLicenseFee.Text.Replace(",", "").Replace("$", "");

                totalamtOverhead = Convert.ToDecimal(txt83) + Convert.ToDecimal(txt82) + Convert.ToDecimal(txt81) + Convert.ToDecimal(txt80) +
                    Convert.ToDecimal(txt79) + Convert.ToDecimal(txt78) + Convert.ToDecimal(txt73);

                txtTotOverhead.Text = totalamtOverhead.ToString("C");

            }
            catch { }

        }



        /// <summary>
        /// Save button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button7_Click(object sender, EventArgs e)
        {

            updateCalculations();  //  .Replace(",", "").Replace("$", "")

            string lcYear = txtYear.Text.Trim();
            string lcEOW = txtWeek.Text.Trim();
            string lcNetSales = txtNetSales.Text.Trim().Replace(",", "").Replace("$", "");
            string lcHealth = txtRetire.Text.Trim().Replace(",", "").Replace("$", "");
            string lcRetire = txtHealth.Text.Trim().Replace(",", "").Replace("$", "");

            string lcfPrimSupp = txtPrimSup.Text.Trim().Replace(",", "").Replace("$", "");
            string lcfOthSupp = txtOtherSupp.Text.Trim().Replace(",", "").Replace("$", "");
            string lcfBread = txtBread.Text.Trim().Replace(",", "").Replace("$", "");
            string lcfBev = txtBev.Text.Trim().Replace(",", "").Replace("$", "");
            string lcfProd = txtProd.Text.Trim().Replace(",", "").Replace("$", "");
            string lcfCarbon = txtCarbDio.Text.Trim().Replace(",", "").Replace("$", "");
            string lcfTotFood = txtFoodTot.Text.Trim().Replace(",", "").Replace("$", "");

            string lcoMort = txtMortgage.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoLoan = txtLoan.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoAssoc = txtAssociation.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoPropTax = txtPropTax.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoAdvCoop = txtAdvCoop.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoNatAdver = txtNationalAdv.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoLicenseFee = txtLicenseFee.Text.Trim().Replace(",", "").Replace("$", "");
            string lcoTotOverhead = txtTotOverhead.Text.Trim().Replace(",", "").Replace("$", "");

            string lceAccount = txtAccount.Text.Trim().Replace(",", "").Replace("$", "");
            string lceBank = txtBank.Text.Trim().Replace(",", "").Replace("$", "");
            string lceCC = txtCC.Text.Trim().Replace(",", "").Replace("$", "");
            string lceFuel = txtFuel.Text.Trim().Replace(",", "").Replace("$", "");
            string lceLegal = txtLegal.Text.Trim().Replace(",", "").Replace("$", "");
            string lceLicensePerm = txtLicense.Text.Trim().Replace(",", "").Replace("$", "");
            string lcePayroll = txtPayroll.Text.Trim().Replace(",", "").Replace("$", "");
            string lceInsur = txtInsur.Text.Trim().Replace(",", "").Replace("$", "");
            string lceWorkComp = txtWorkComp.Text.Trim().Replace(",", "").Replace("$", "");
            string lceAdvertise = txtAdvertising.Text.Trim().Replace(",", "").Replace("$", "");
            string lceCharitable = txtCharitableComp.Text.Trim().Replace(",", "").Replace("$", "");
            string lceAuto = txtAuto.Text.Trim().Replace(",", "").Replace("$", "");
            string lceCash = txtCashShort.Text.Trim().Replace(",", "").Replace("$", "");
            string lceElect = txtElectrical.Text.Trim().Replace(",", "").Replace("$", "");
            string lceGeneral = txtGeneral.Text.Trim().Replace(",", "").Replace("$", "");
            string lceHVAC = txtHVAC.Text.Trim().Replace(",", "").Replace("$", "");
            string lceLawn = txtLawn.Text.Trim().Replace(",", "").Replace("$", "");
            string lcePaint = txtPaint.Text.Trim().Replace(",", "").Replace("$", "");
            string lcePlumb = txtPlumb.Text.Trim().Replace(",", "").Replace("$", "");
            string lceRemodel = txtRemodel.Text.Trim().Replace(",", "").Replace("$", "");
            string lceStruct = txtStructural.Text.Trim().Replace(",", "").Replace("$", "");
            string lceDishMach = txtDishMach.Text.Trim().Replace(",", "").Replace("$", "");
            string lceJanitorial = txtJanitorial.Text.Trim().Replace(",", "").Replace("$", "");
            string lceOfficeComp = txtOffice.Text.Trim().Replace(",", "").Replace("$", "");
            string lceRestaurant = txtRestaurant.Text.Trim().Replace(",", "").Replace("$", "");
            string lceUniform = txtUniform.Text.Trim().Replace(",", "").Replace("$", "");
            string lceData = txtDataTele.Text.Trim().Replace(",", "").Replace("$", "");
            string lceElectric = txtElectricity.Text.Trim().Replace(",", "").Replace("$", "");
            string lceMusic = txtMusic.Text.Trim().Replace(",", "").Replace("$", "");
            string lceNatGas = txtNatGas.Text.Trim().Replace(",", "").Replace("$", "");
            string lceSecurity = txtSecurity.Text.Trim().Replace(",", "").Replace("$", "");
            string lceTrash = txtTrash.Text.Trim().Replace(",", "").Replace("$", "");
            string lceWaterSewer = txtWater.Text.Trim().Replace(",", "").Replace("$", "");
            string lceTotExpense = txtTotExpense.Text.Trim().Replace(",", "").Replace("$", "");

            string lclHost = txtHost.Text.Trim().Replace(",", "").Replace("$", "");
            string lclCook = txtCooks.Text.Trim().Replace(",", "").Replace("$", "");
            string lclServer = txtServers.Text.Trim().Replace(",", "").Replace("$", "");
            string lclDMO = txtDMO.Text.Trim().Replace(",", "").Replace("$", "");
            string lclSuperv = txtSupervisor.Text.Trim().Replace(",", "").Replace("$", "");
            string lclOvertime = txtOvertime.Text.Trim().Replace(",", "").Replace("$", "");
            string lclGenManager = txtGenManager.Text.Trim().Replace(",", "").Replace("$", "");
            string lclManager = txtManager.Text.Trim().Replace(",", "").Replace("$", "");
            string lclBonus = txtBonus.Text.Trim().Replace(",", "").Replace("$", "");
            string lclPayTax = txtPayrollTax.Text.Trim().Replace(",", "").Replace("$", "");
            string lclTotLabor = txtTotLabor.Text.Trim().Replace(",", "").Replace("$", "");

            string lcSQL = "";

            lcSQL = "UPDATE dynamicelements..tb_LaborCost SET LaborCost=" + lclTotLabor + " ,HostCashier=" + lclHost + " ,Cooks=" + lclCook + " ,Servers=" + lclServer + " ," +
                "DMO=" + lclDMO + " ,Supervisor=" + lclSuperv + ", Overtime=" + lclOvertime + ", GeneralManager=" + lclGenManager + ", Manager=" + lclManager + "," +
                " Bonus=" + lclBonus + ", PayrollTax=" + lclPayTax + " WHERE Week='" + lcEOW + "' and IDS=" + lcStoreName;  // 138  158  168  180  192  197  209  218  222
            SQLCommand(lcSQL);

            lcSQL = " UPDATE dynamicelements..tb_ExpenseCost SET ExpenseCost =" + lceTotExpense + ",Accounting =" + lceAccount + ",Bank =" + lceBank + ",CreditCard =" + lceCC + ",Fuel =" + lceFuel + "," +
                "Legal =" + lceLegal + ",License =" + lceLicensePerm + ",PayrollProc =" + lcePayroll + ",Insurance =" + lceInsur + ",WorkersComp =" + lceWorkComp + ",Advertising =" + lceAdvertise + "," +
                "Charitable =" + lceCharitable + ",Auto =" + lceAuto + ",CashShortage =" + lceCash + ",Electrical =" + lceElect + ",General =" + lceGeneral + ",HVAC =" + lceHVAC + ",Lawn =" + lceLawn + ",Painting =" + lcePaint + "," +
                "Plumbing =" + lcePlumb + ",Remodeling =" + lceRemodel + ",Structural =" + lceStruct + ",DishMachine =" + lceDishMach + ",Janitorial =" + lceJanitorial + ",Office =" + lceOfficeComp + "," +
                "Restaurant =" + lceRestaurant + ",Uniforms =" + lceUniform + ",Data =" + lceData + ",Electricity =" + lceElectric + ",Music =" + lceMusic + ",NaturalGas =" + lceNatGas + ",Security =" + lceSecurity + "," +
                "Trash =" + lceTrash + ",WaterSewer =" + lceWaterSewer + " WHERE Week ='" + lcEOW + "' and IDS=" + lcStoreName;  // 138  158  168  180  192  197  209  218  222
            SQLCommand(lcSQL);

            lcSQL = " UPDATE dynamicelements..tb_FoodCost SET FoodCost =" + lcfTotFood + ",PrimSupp =" + lcfPrimSupp + ",OthSupp =" + lcfOthSupp + ",Bread =" + lcfBread + "," +
                "Beverage =" + lcfBev + ",Produce =" + lcfProd + ",CarbonDioxide =" + lcfCarbon + " WHERE Week ='" + lcEOW + "' and IDS=" + lcStoreName;  // 138  158  168  180  192  197  209  218  222
            SQLCommand(lcSQL);

            lcSQL = " UPDATE dynamicelements..tb_NetSales SET NetSales =" + lcNetSales + ",Healthcare =" + lcHealth + ",Retirement =" + lcRetire + " WHERE Week = '" + lcEOW + "' and IDS=" + lcStoreName;  // 138  158  168  180  192  197  209  218  222 
            SQLCommand(lcSQL);

            lcSQL = " UPDATE dynamicelements..tb_OverheadCost set OverheadCost =" + lcoTotOverhead + ",Mortgage =" + lcoMort + ",LoanPayment =" + lcoLoan + ",Association =" + lcoAssoc + ",PropertyTax =" + lcoPropTax + "," +
                "AdvertisingCoop =" + lcoAdvCoop + ",NationalAdvertise =" + lcoNatAdver + ",LicensingFee =" + lcoLicenseFee + " WHERE Week = '" + lcEOW + "' and IDS=" + lcStoreName;  // 138  158  168  180  192  197  209  218  222
            SQLCommand(lcSQL);

            MessageBox.Show("Done!");

        }


        private void textBox3_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtNetSales.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtNetSales.Text = val.ToString("C");
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox84_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtPrimSup.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtPrimSup.Text = val.ToString("C");
        }

        private void textBox84_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox77_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtOtherSupp.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtOtherSupp.Text = val.ToString("C");
        }

        private void textBox77_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox76_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtBread.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtBread.Text = val.ToString("C");
        }

        private void textBox76_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox75_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtBev.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtBev.Text = val.ToString("C");
        }

        private void textBox75_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox69_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtProd.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtProd.Text = val.ToString("C");
        }

        private void textBox69_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox68_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtCarbDio.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtCarbDio.Text = val.ToString("C");
        }

        private void textBox68_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtFoodTot.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtFoodTot.Text = val.ToString("C");
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox90_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtHost.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtHost.Text = val.ToString("C");
        }

        private void textBox90_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox89_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtCooks.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtCooks.Text = val.ToString("C");
        }

        private void textBox89_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox88_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtServers.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtServers.Text = val.ToString("C");
        }

        private void textBox88_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox87_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtDMO.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtDMO.Text = val.ToString("C");
        }

        private void textBox87_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox86_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtSupervisor.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtSupervisor.Text = val.ToString("C");
        }

        private void textBox86_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox85_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtOvertime.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtOvertime.Text = val.ToString("C");
        }

        private void textBox85_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox74_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtGenManager.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtGenManager.Text = val.ToString("C");
        }

        private void textBox74_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox72_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtManager.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtManager.Text = val.ToString("C");
        }

        private void textBox72_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox71_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtBonus.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtBonus.Text = val.ToString("C");
        }

        private void textBox71_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox70_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtPayrollTax.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtPayrollTax.Text = val.ToString("C");
        }

        private void textBox70_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtTotLabor.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtTotLabor.Text = val.ToString("C");
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox83_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtMortgage.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtMortgage.Text = val.ToString("C");
        }

        private void textBox83_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox82_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtLoan.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtLoan.Text = val.ToString("C");
        }

        private void textBox82_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox81_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtAssociation.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtAssociation.Text = val.ToString("C");
        }

        private void textBox81_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox80_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtPropTax.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtPropTax.Text = val.ToString("C");
        }

        private void textBox80_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox79_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtAdvCoop.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtAdvCoop.Text = val.ToString("C");
        }

        private void textBox79_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox78_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtNationalAdv.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtNationalAdv.Text = val.ToString("C");
        }

        private void textBox78_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox73_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtLicenseFee.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtLicenseFee.Text = val.ToString("C");
        }

        private void textBox73_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtTotOverhead.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtTotOverhead.Text = val.ToString("C");
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox27_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtAccount.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtAccount.Text = val.ToString("C");
        }

        private void textBox27_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox26_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtBank.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtBank.Text = val.ToString("C");
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox25_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtCC.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtCC.Text = val.ToString("C");
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox24_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtFuel.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtFuel.Text = val.ToString("C");
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox23_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtLegal.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtLegal.Text = val.ToString("C");
        }

        private void textBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox22_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtLicense.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtLicense.Text = val.ToString("C");
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox28_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtPayroll.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtPayroll.Text = val.ToString("C");
        }

        private void textBox28_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox30_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtInsur.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtInsur.Text = val.ToString("C");
        }

        private void textBox30_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox29_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtWorkComp.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtWorkComp.Text = val.ToString("C");
        }

        private void textBox29_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox32_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtAdvertising.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtAdvertising.Text = val.ToString("C");
        }

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox31_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtCharitableComp.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtCharitableComp.Text = val.ToString("C");
        }

        private void textBox31_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox21_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtAuto.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtAuto.Text = val.ToString("C");
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox20_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtCashShort.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtCashShort.Text = val.ToString("C");
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox34_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtElectrical.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtElectrical.Text = val.ToString("C");
        }

        private void textBox34_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox33_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtGeneral.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtGeneral.Text = val.ToString("C");
        }

        private void textBox33_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox19_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtHVAC.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtHVAC.Text = val.ToString("C");
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox35_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtLawn.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtLawn.Text = val.ToString("C");
        }

        private void textBox35_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox36_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtPaint.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtPaint.Text = val.ToString("C");
        }

        private void textBox36_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox37_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtPlumb.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtPlumb.Text = val.ToString("C");
        }

        private void textBox37_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox38_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtRemodel.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtRemodel.Text = val.ToString("C");
        }

        private void textBox38_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox39_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtStructural.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtStructural.Text = val.ToString("C");
        }

        private void textBox39_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox43_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtDishMach.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtDishMach.Text = val.ToString("C");
        }

        private void textBox43_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox42_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtJanitorial.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtJanitorial.Text = val.ToString("C");
        }

        private void textBox42_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox44_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtOffice.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtOffice.Text = val.ToString("C");
        }

        private void textBox44_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox41_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtRestaurant.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtRestaurant.Text = val.ToString("C");
        }

        private void textBox41_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox40_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtUniform.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtUniform.Text = val.ToString("C");
        }

        private void textBox40_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox18_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtDataTele.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtDataTele.Text = val.ToString("C");
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox45_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtElectricity.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtElectricity.Text = val.ToString("C");
        }

        private void textBox45_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox46_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtMusic.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtMusic.Text = val.ToString("C");
        }

        private void textBox46_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox47_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtNatGas.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtNatGas.Text = val.ToString("C");
        }

        private void textBox47_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox48_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtSecurity.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtSecurity.Text = val.ToString("C");
        }

        private void textBox48_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox49_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtTrash.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtTrash.Text = val.ToString("C");
        }

        private void textBox49_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox50_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtWater.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtWater.Text = val.ToString("C");
        }

        private void textBox50_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtTotExpense.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtTotExpense.Text = val.ToString("C");
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtRetire.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtRetire.Text = val.ToString("C");
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {

            updateCalculations();

            string value = txtHealth.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                txtHealth.Text = val.ToString("C");
        }

        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        public static FileInfo GetNewestFile(DirectoryInfo directory)
        {
            return directory.GetFiles()
                .Union(directory.GetDirectories().Select(d => GetNewestFile(d)))
                .OrderByDescending(f => (f == null ? DateTime.MinValue : f.LastWriteTime))
                .FirstOrDefault();
        }



        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(System.Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtInvHold.Text = "FOOD";

            switch (this.tabControl1.SelectedIndex)
            {
                case 0:
                    txtInvHold.Text = "FOOD";
                    break;

                case 1:
                    txtInvHold.Text = "EXPENSE";
                    break;

                case 2:
                    txtInvHold.Text = "LABOR";
                    break;

                case 3:
                    txtInvHold.Text = "OVERHEAD";
                    break;

                default:
                    txtInvHold.Text = "FOOD";
                    break;
            }

            updateCalculations();

        }


        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            decimal totalSalary = 0;
            decimal amt = 0;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                var value = dataGridView1.Rows[i].Cells[1].Value;
                if (value != DBNull.Value)
                {
                    amt = Convert.ToDecimal(value);
                    totalSalary += amt;
                }
            }

            txtTotInvoice.Text = totalSalary.ToString("C");
        }


        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            var ldate = monthCalendar1.SelectionRange.Start.Date;  // .ToShortDateString()
            var nextSunday = Dates.DTOC(ldate.AddDays(7 - (int)ldate.DayOfWeek));
            var lcyear = Dates.CTOD(nextSunday).Year.ToString();

            txtWeek.Text = nextSunday;
            txtYear.Text = lcyear;
            refreshFormFields();
        }



        /// <summary>
        /// Apply Button
        /// Vendor Inv combine
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            string lcSQL = "";
            string lcSQLa = "";
            string lcSQLb = "";
            string lcEOW = txtWeek.Text.Trim();

            var lcInvDate = DateTime.Parse(txtInvDate.Text);

            string lcVendor = vendorIDTextBox.Text.Trim();
            string lcVendorInv = txtInvNumb.Text.Trim();
            string lcCat = cbCategory.Text.Trim();
            decimal lcAmt = 0m;

            string lcVendName = vendorNameTextBox.Text.Trim();
            string lcSalesP = salesPersonTextBox.Text.Trim();
            string lcPhone = phoneTextBox.Text.Trim();
            string lcAddress1 = addressLine1TextBox.Text.Trim();
            string lcAddress2 = addressLine2TextBox.Text.Trim();
            string lcCity = cityTextBox.Text.Trim();
            string lcState = stateProvinceTextBox.Text.Trim();
            string lcCountry = countryRegionTextBox.Text.Trim();
            string lcPostal = postalCodeTextBox.Text.Trim();

            decimal lcTotVal = Convert.ToDecimal(txtTotInvoice.Text.Replace(",", "").Replace("$", ""));

            for (int i = 0; i < (dataGridView1.Rows.Count - 1); i++)
            {
                string lcItem = dataGridView1.Rows[i].Cells[0].Value.ToString();
                lcAmt = Convert.ToDecimal(dataGridView1.Rows[i].Cells[1].Value.ToString());

                lcSQL = " INSERT INTO dynamicelements..tb_VendorInv (Week,IDS,InvDate,VendorID,InvNumber,Category,Item,Amount) VALUES " +
                    "('" + lcEOW + "', " + lcStoreName + ", '" + lcInvDate + "' , '" + lcVendor + "', '" + lcVendorInv + "', '" + lcCat + "', '" + lcItem + "', " + lcAmt + ") "; // 138  158  168  180  192  197  209  218  222
                SQLCommand(lcSQL);
            }

            lcSQLa = " select * from dynamicelements..tb_Vendors where VendorID='%" + lcVendor + "% '";
            OdbcDataReader reader = GetData(lcSQLa);

            if (!reader.HasRows)
            {
                lcSQLa = " INSERT INTO dynamicelements..tb_Vendors (VendorID,VendorName,SalesPerson,Phone,AddressLine1,AddressLine2,City,StateProvince,CountryRegion,PostalCode) VALUES " +
                    " ('" + lcVendor + "','" + lcVendName + "','" + lcSalesP + "','" + lcPhone + "','" + lcAddress1 + "','" + lcAddress2 + "','" + lcCity + "','" + lcState + "','" + lcCountry + "','" + lcPostal + "') ";
                SQLCommand(lcSQLa);

            }
            reader.Close();

            string lcSQLz = "select " + lcCat + " from dynamicelements..vw_OrderLogs where week = '" + lcEOW + "' and AddressID =" + lcStoreName;
            OdbcDataReader readerz = GetData(lcSQLz);

            decimal lcnumb = 0m;
            if (readerz.HasRows)
                lcnumb = Convert.ToDecimal(readerz[lcCat].ToString());

            readerz.Close();
            decimal lcNewTot = lcTotVal + lcnumb;

            switch (txtInvHold.Text.Trim())
            {
                case "FOOD":
                    lcSQLb = " UPDATE dynamicelements..tb_FoodCost SET " + lcCat + " = " + lcNewTot + " WHERE Week='" + lcEOW + "' and IDS =" + lcStoreName;
                    break;

                case "EXPENSES":
                    lcSQLb = " UPDATE dynamicelements..tb_ExpenseCost SET " + lcCat + " = " + lcNewTot + " WHERE Week='" + lcEOW + "' and IDS =" + lcStoreName;
                    break;

                case "LABOR":
                    lcSQLb = " UPDATE dynamicelements..tb_LaborCost SET " + lcCat + " = " + lcNewTot + " WHERE Week='" + lcEOW + "' and IDS =" + lcStoreName;
                    break;

                case "OVERHEAD":
                    lcSQLb = " UPDATE dynamicelements..tb_OverheadCost SET " + lcCat + " = " + lcNewTot + " WHERE Week='" + lcEOW + "' and IDS =" + lcStoreName;
                    break;
            }
            SQLCommand(lcSQLb);

            txtInvDate.Text = "";
            vendorIDTextBox.Text = "";
            txtVndSearch.Text = "";
            txtInvNumb.Text = "";
            cbCategory.Text = "";
            vendorNameTextBox.Text = "";
            salesPersonTextBox.Text = "";
            phoneTextBox.Text = "";
            addressLine1TextBox.Text = "";
            addressLine2TextBox.Text = "";
            cityTextBox.Text = "";
            stateProvinceTextBox.Text = "";
            countryRegionTextBox.Text = "";
            postalCodeTextBox.Text = "";
            txtTotInvoice.Text = "";
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();

            refreshFormFields();

            MessageBox.Show("Done!");

        }



        /// <summary>
        /// This is the Vendor search textbox. Do a fuzzy search in the database to see if it finds the vendor. If not then have user enter info. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Leave(object sender, EventArgs e)
        {

            string lcval = this.txtVndSearch.Text;

            string lcSQL = "select * from dynamicelements..tb_Vendors where VendorID='%" + lcval + "% '";
            OdbcDataReader reader = GetData(lcSQL);

            if (reader.HasRows)
            {
                //  MessageBox.Show("Found something!");
                vendorIDTextBox.Text = reader["VendorID"].ToString();
                vendorNameTextBox.Text = reader["VendorName"].ToString();
                salesPersonTextBox.Text = reader["SalesPerson"].ToString();
                phoneTextBox.Text = reader["Phone"].ToString();
                addressLine1TextBox.Text = reader["AddressLine1"].ToString();
                addressLine2TextBox.Text = reader["AddressLine2"].ToString();
                cityTextBox.Text = reader["City"].ToString();
                stateProvinceTextBox.Text = reader["StateProvince"].ToString();
                countryRegionTextBox.Text = reader["CountryRegion"].ToString();
                postalCodeTextBox.Text = reader["PostalCode"].ToString();
            }
            else
            {
                MessageBox.Show("Found nothing!");
                vendorIDTextBox.Focus();
            }
            reader.Close();
        }


        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column1_KeyPress);
            if (dataGridView1.CurrentCell.ColumnIndex == 1) //Desired Column
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column1_KeyPress);
                }
            }
        }


        private void Column1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // allowed numeric and one dot  ex. 10.23
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if (e.KeyChar == '.' && (sender as TextBox).Text.IndexOf('.') > -1)
            {
                e.Handled = true;
            }
        }
    }
}