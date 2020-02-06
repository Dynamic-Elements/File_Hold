using Syncfusion.OCRProcessor;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using ScanIt;
using IronOcr;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using VFPToolkit;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using WIA;
using System.Linq;

namespace Accounting_PL
{
    public partial class Form1 : Form
    {

        string appPath = AppDomain.CurrentDomain.BaseDirectory;
        string curDir = Files.AddBS(Files.CurDir());
        // MessageBox.Show("here " + curDir);
        string baseCurDir = Files.AddBS(Path.GetFullPath(Path.Combine(Files.CurDir(), @"..\..\..\")));
        //  MessageBox.Show("here " + baseCurDir);
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

            var date = DateTime.Now;
            var lastSunday = Dates.DTOC(date.AddDays(-(int)date.DayOfWeek));  // Grabs the past Sunday for Week End
            var lYear = DateTime.Now.Year.ToString();
            txtWeek.Text = lastSunday;
            txtYear.Text = lYear;   // Yr.Substring(0,4);


            string lcServer = "dynamicelements.database.windows.net";  // playgroup.database.windows.net
            string lcODBC = "ODBC Driver 17 for SQL Server";
            string lcDB = "dynamicelements";
            // string lcPort = "3306";  //  Port=" + lcPort + ";
            string lcUser = "tbmaster";
            string lcProv = "SQLOLEDB";
            string lcPass = "Fzk4pktb";     // Smartman55  Fzk4pktb
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            cnn.Open();


            string lcSQL = "SELECT * from dynamicelements..tb_Config where Year='" + lYear + "'";   // Week='" + textBox1.Text.Trim() + "'";   '12/30/2018'  v" + textBox1.Text.Trim() + "
            OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            OdbcDataReader reader = cmd.ExecuteReader();

            bool fiscialLeapYear;
            if (reader.HasRows)
            {
                fiscialLeapYear = true;
                checkBox3.Checked = true;
            }
            else { }

            txtInvHold.Text = "FOOD";


            //if (Int32.Parse(lYear) % 400 == 0 || (Int32.Parse(lYear) % 4 == 0 && Int32.Parse(lYear) % 100 != 0))
            //    checkBox3.Checked = true;
            // MessageBox.Show("Leap year!");


            // dynamicelements..vw_OrderLogs    //  Will need to create stored procedures
            string lcSQLa = "select * from vw_OrderLogs where week='" + lastSunday + "'";   // Week='" + textBox1.Text.Trim() + "'";   '12/30/2018'  v" + textBox1.Text.Trim() + "  12/30/2018
            OdbcCommand cmda = new OdbcCommand(lcSQLa, cnn);
            OdbcDataReader readera = cmda.ExecuteReader();
            //// MessageBox.Show(Convert.ToString(reader.GetOrdinal("NetSales")));

            if (readera.HasRows)
            {

                txtNetSales.Text = readera["NetSales"].ToString();
                txtRetire.Text = readera["Healthcare"].ToString();
                txtHealth.Text = readera["Retirement"].ToString();

                txtPrimSup.Text = readera["PrimSupp"].ToString();
                txtOtherSupp.Text = readera["OthSupp"].ToString();
                txtBread.Text = readera["Bread"].ToString();
                txtBev.Text = readera["Beverage"].ToString();
                txtProd.Text = readera["Produce"].ToString();
                txtCarbDio.Text = readera["CarbonDioxide"].ToString();
                txtFoodTot.Text = readera["FoodCost"].ToString();

                txtMortgage.Text = readera["Mortgage"].ToString();
                txtLoan.Text = readera["LoanPayment"].ToString();
                txtAssociation.Text = readera["Association"].ToString();
                txtPropTax.Text = readera["PropertyTax"].ToString();
                txtAdvCoop.Text = readera["AdvertisingCoop"].ToString();
                txtNationalAdv.Text = readera["NationalAdvertise"].ToString();
                txtLicenseFee.Text = readera["LicensingFee"].ToString();
                txtTotOverhead.Text = readera["OverheadCost"].ToString();

                txtAccount.Text = readera["Accounting"].ToString();
                txtBank.Text = readera["Bank"].ToString();
                txtCC.Text = readera["CreditCard"].ToString();
                txtFuel.Text = readera["Fuel"].ToString();
                txtLegal.Text = readera["Legal"].ToString();
                txtLicense.Text = readera["License"].ToString();
                txtPayroll.Text = readera["PayrollProc"].ToString();
                txtInsur.Text = readera["Insurance"].ToString();
                txtWorkComp.Text = readera["WorkersComp"].ToString();
                txtAdvertising.Text = readera["Advertising"].ToString();
                txtCharitableComp.Text = readera["Charitable"].ToString();
                txtAuto.Text = readera["Auto"].ToString();
                txtCashShort.Text = readera["CashShortage"].ToString();
                txtElectrical.Text = readera["Electrical"].ToString();
                txtGeneral.Text = readera["General"].ToString();
                txtHVAC.Text = readera["HVAC"].ToString();
                txtLawn.Text = readera["Lawn"].ToString();
                txtPaint.Text = readera["Painting"].ToString();
                txtPlumb.Text = readera["Plumbing"].ToString();
                txtRemodel.Text = readera["Remodeling"].ToString();
                txtStructural.Text = readera["Structural"].ToString();
                txtDishMach.Text = readera["DishMachine"].ToString();
                txtJanitorial.Text = readera["Janitorial"].ToString();
                txtOffice.Text = readera["Office"].ToString();
                txtRestaurant.Text = readera["Restaurant"].ToString();
                txtUniform.Text = readera["Uniforms"].ToString();
                txtDataTele.Text = readera["Data"].ToString();
                txtElectricity.Text = readera["Electricity"].ToString();
                txtMusic.Text = readera["Music"].ToString();
                txtNatGas.Text = readera["NaturalGas"].ToString();
                txtSecurity.Text = readera["Security"].ToString();
                txtTrash.Text = readera["Trash"].ToString();
                txtWater.Text = readera["WaterSewer"].ToString();
                txtTotExpense.Text = readera["ExpenseCost"].ToString();

                txtHost.Text = readera["HostCashier"].ToString();
                txtCooks.Text = readera["Cooks"].ToString();
                txtServers.Text = readera["Servers"].ToString();
                txtDMO.Text = readera["DMO"].ToString();
                txtSupervisor.Text = readera["Supervisor"].ToString();
                txtOvertime.Text = readera["Overtime"].ToString();
                txtGenManager.Text = readera["GeneralManager"].ToString();
                txtManager.Text = readera["Manager"].ToString();
                txtBonus.Text = readera["Bonus"].ToString();
                txtPayrollTax.Text = readera["PayrollTax"].ToString();
                txtTotLabor.Text = readera["LaborCost"].ToString();

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
            cnn.Close();



            // https://cloud.google.com/vision/docs/ocr#vision_text_detection-csharp
            // https://developers.google.com/vision/android/text-overview
            // https://cloud.google.com/vision/docs/pdf


        }


        /// <summary>
        /// Excel Code
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button1_Click(object sender, EventArgs e)
        {

            updateCalculations();


            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;

            string lexfolder = Files.AddBS(baseCurDir + "FinancialFolder");
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

            xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            xlApp.DisplayAlerts = false;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // xlWorkSheet.Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(1);
            //  xlWorkBook.Worksheets.Add();

            //iWeeksPerMonth = 4  sMonth = "January"
            //iWeeksPerMonth = 4  sMonth = "February"
            //iWeeksPerMonth = 5  sMonth = "March"
            //iWeeksPerMonth = 4  sMonth = "April"
            //iWeeksPerMonth = 4  sMonth = "May"
            //iWeeksPerMonth = 5  sMonth = "June"
            //iWeeksPerMonth = 4  sMonth = "July"
            //iWeeksPerMonth = 4  sMonth = "August"
            //iWeeksPerMonth = 5  sMonth = "September"
            //iWeeksPerMonth = 4  sMonth = "October"
            //iWeeksPerMonth = 4  sMonth = "November"
            //iWeeksPerMonth = 5 or iWeeksPerMonth = 6  sMonth = "December"


            var coll = new Excel.Worksheet[14];

            for (int i = 1; i < 14; i++)
            {
                coll[i] = xlWorkBook.Worksheets.Add();
                coll[i].Name = (i == 13) ? "YTD" : CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i);

                coll[i].Cells[1, 1] = "Miami Springs - ###";
                coll[i].Cells[1, 1].Font.Bold = true;

                coll[i].Cells[1, 2] = "Dates";
                coll[i].Range["B1:C1"].Merge();
                coll[i].Cells[2, 2] = "Week 1";
                coll[i].Cells[3, 2] = "$";
                coll[i].Cells[3, 3] = "%";

                coll[i].Cells[1, 4] = "Dates";
                coll[i].Range["d1:e1"].Merge();
                coll[i].Cells[2, 4] = "Week 2";
                coll[i].Cells[3, 4] = "$";
                coll[i].Cells[3, 5] = "%";

                coll[i].Cells[1, 6] = "Dates";
                coll[i].Range["f1:g1"].Merge();
                coll[i].Cells[2, 6] = "Week 3";
                coll[i].Cells[3, 6] = "$";
                coll[i].Cells[3, 7] = "%";

                coll[i].Cells[1, 8] = "Dates";
                coll[i].Range["h1:i1"].Merge();
                coll[i].Cells[2, 8] = "Week 4";
                coll[i].Cells[3, 8] = "$";
                coll[i].Cells[3, 9] = "%";

                if (i == 3 || i == 6 || i == 9 || i == 12)  // Extra week
                {

                    coll[i].Cells[1, 10] = "Dates";
                    coll[i].Range["j1:k1"].Merge();
                    coll[i].Cells[2, 10] = "Week 5";
                    coll[i].Cells[3, 10] = "$";
                    coll[i].Cells[3, 11] = "%";

                }
                else { }

                if (checkBox3.Checked == true && i == 12)  // Extra week
                {

                    coll[i].Cells[1, 12] = "Dates";
                    coll[i].Range["j1:k1"].Merge();
                    coll[i].Cells[2, 12] = "Week 6";
                    coll[i].Cells[3, 12] = "$";
                    coll[i].Cells[3, 13] = "%";

                }
                else { }

                coll[i].Cells[4, 1] = "Net Sales";
                coll[i].Cells[4, 1].Font.Bold = true;
                coll[i].Cells[5, 1] = "Primary Supplier";
                coll[i].Cells[6, 1] = "Other Suppliers";
                coll[i].Cells[7, 1] = "Bread";
                coll[i].Cells[8, 1] = "Produce";
                coll[i].Cells[9, 1] = "Carbon Dioxide";
                coll[i].Cells[10, 1] = "Food Cost";
                coll[i].Cells[10, 1].Font.Bold = true;
                coll[i].Cells[11, 1] = "Craft labor";
                coll[i].Cells[12, 1] = "Host/Cashier";
                coll[i].Cells[13, 1] = "Cooks";
                coll[i].Cells[14, 1] = "Servers";
                coll[i].Cells[15, 1] = "DMO";
                coll[i].Cells[16, 1] = "Supervisors";
                coll[i].Cells[17, 1] = "Overtime";
                coll[i].Cells[18, 1] = "Management";
                coll[i].Cells[19, 1] = "General Manager";
                coll[i].Cells[20, 1] = "Manager";
                coll[i].Cells[21, 1] = "Bonuses";
                coll[i].Cells[22, 1] = "Labor Expenses";
                coll[i].Cells[23, 1] = "Payroll Taxes";
                coll[i].Cells[24, 1] = "Labor Cost";
                coll[i].Cells[24, 1].Font.Bold = true;
                coll[i].Cells[25, 1] = "Fees";
                coll[i].Cells[26, 1] = "Accounting";
                coll[i].Cells[27, 1] = "Bank";
                coll[i].Cells[28, 1] = "Credit Card";
                coll[i].Cells[29, 1] = "Fuel/Delivery";
                coll[i].Cells[30, 1] = "Legal";
                coll[i].Cells[31, 1] = "Licenses/Permits";
                coll[i].Cells[32, 1] = "Payroll Processing";
                coll[i].Cells[33, 1] = "Insurance";
                coll[i].Cells[34, 1] = "Insurance";
                coll[i].Cells[35, 1] = "Workers Compensation";
                coll[i].Cells[36, 1] = "Local Marketing";
                coll[i].Cells[37, 1] = "Advertising";
                coll[i].Cells[38, 1] = "Charitable Contributions";
                coll[i].Cells[39, 1] = "Other";
                coll[i].Cells[40, 1] = "Auto/Travel";
                coll[i].Cells[41, 1] = "Cash Shortages";
                coll[i].Cells[42, 1] = "Repair/Matinenace";
                coll[i].Cells[43, 1] = "Eletrical";
                coll[i].Cells[44, 1] = "General";
                coll[i].Cells[45, 1] = "HVAC";
                coll[i].Cells[46, 1] = "Lawn/Parking";
                coll[i].Cells[47, 1] = "Painting";
                coll[i].Cells[48, 1] = "Plumbing";
                coll[i].Cells[49, 1] = "Remodeling";
                coll[i].Cells[50, 1] = "Structural";
                coll[i].Cells[51, 1] = "Supplies";
                coll[i].Cells[52, 1] = "Dish Machine";
                coll[i].Cells[53, 1] = "Janitorial";
                coll[i].Cells[54, 1] = "Office/Computer";
                coll[i].Cells[55, 1] = "Restuarant";
                coll[i].Cells[56, 1] = "Uniforms";
                coll[i].Cells[57, 1] = "Utilities";
                coll[i].Cells[58, 1] = "Data/Telephone";
                coll[i].Cells[59, 1] = "Electricity";
                coll[i].Cells[60, 1] = "Music";
                coll[i].Cells[61, 1] = "Natural Gas";
                coll[i].Cells[62, 1] = "Security";
                coll[i].Cells[63, 1] = "Trash";
                coll[i].Cells[64, 1] = "Water & Sewer";
                coll[i].Cells[65, 1] = "Expenses Cost";
                coll[i].Cells[65, 1].Font.Bold = true;
                coll[i].Cells[66, 1] = "Overhead";
                coll[i].Cells[67, 1] = "Mortgage/Rent";
                coll[i].Cells[68, 1] = "Loan Payments";
                coll[i].Cells[69, 1] = "Association/CAM Fees";
                coll[i].Cells[70, 1] = "Property Taxes";
                coll[i].Cells[71, 1] = "Advertising Coop";
                coll[i].Cells[72, 1] = "National Advertising";
                coll[i].Cells[73, 1] = "Licensing Fee";
                coll[i].Cells[74, 1] = "Overhead Cost";
                coll[i].Cells[74, 1].Font.Bold = true;
                coll[i].Cells[76, 1] = "Total Cost";
                coll[i].Cells[76, 1].Font.Bold = true;
                coll[i].Cells[78, 1] = "Return on Revenue";
                coll[i].Cells[78, 1].Font.Bold = true;

                coll[i].Columns.AutoFit();
                coll[i].Rows.AutoFit();
            }

            xlWorkBook.Sheets["Sheet1"].Delete();
            xlApp.Visible = true;

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

            updateCalculations();

            string lscfolder = Files.AddBS(baseCurDir + "Scanned_Documents");
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
            int dpi = 600;  // 150  300  600  720  1200  1270  1440
            WiaWrapper obj = new WiaWrapper();
            obj.SelectScanner();
            obj.Scan(true, dpi, path, adf, duplex);  //  Scan(bool rotatePage, int DPI, string filepath, bool useAdf, bool duplex)

            FileInfo oldnewestFile = GetNewestFile(new DirectoryInfo(path));
            string value = "Document Name";
            if (InputBox("New document", "New document name:", ref value) == DialogResult.OK)
            {
                Name = oldnewestFile.DirectoryName + "\\" + value + ".jpeg";
            }
            File.Move(oldnewestFile.FullName, Name);

            //var Ocr = new IronOcr.AutoOcr();
            //var Result = Ocr.Read(@"C:\path\to\image.png");
            //Console.WriteLine(Result.Text);

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

            // var testDocument = @"C:\Users\taylo\Documents\File_Hold\Accounting_PL\Scanned_Documents\test_02.jpg";
            var testDocument = Name;
            var Results = Ocr.Read(testDocument);
            // var Results = Ocr.Read(Name);
            // Console.WriteLine(Results.Text);
            MessageBox.Show(Results.Text);

            //string line = null;
            //TextReader readFile = new StreamReader(oldnewestFile.DirectoryName + "\\" + value + ".txt");
            //line = readFile.ReadToEnd();
            // MessageBox.Show(line);
            //readFile.Close();
            //readFile = null;

            ////Create a new PDF document
            //PdfDocument document = new PdfDocument();
            ////Add a page to the document
            //PdfPage page = document.Pages.Add();
            ////Create PDF graphics for a page
            //PdfGraphics graphics = page.Graphics;
            ////Load the image from the disk
            //PdfBitmap imageFile = new PdfBitmap(Name);   //  "Input.jpg"  path
            ////Draw the image
            //graphics.DrawImage(imageFile, 0, 0, page.GetClientSize().Width, page.GetClientSize().Height);
            ////Save the document into stream
            //MemoryStream stream = new MemoryStream();
            //document.Save(stream);
            ////Initialize the OCR processor by providing the path of tesseract binaries(SyncfusionTesseract.dll and liblept168.dll)
            //using (OCRProcessor processor = new OCRProcessor(@"../../Tesseract Binaries/"))
            //{
            //    //Load a PDF document
            //    PdfLoadedDocument lDoc = new PdfLoadedDocument(stream);

            //    //Set OCR language to process
            //    processor.Settings.Language = Languages.English;

            //    //Enable the AutoDetectRotation
            //    processor.Settings.AutoDetectRotation = true;

            //    //Enable native call  
            //    processor.Settings.EnableNativeCall = true;

            //    //Process OCR by providing the PDF document and Tesseract data
            //    String text = processor.PerformOCR(lDoc, @"..\..\Tessdata\");

            //    // Save the PDF file
            //    string lcNewFile = oldnewestFile.DirectoryName + "\\" + value + ".pdf";  //  lscfolder + "Scan_OCR_File" + rand.Next(10, 100) + ".pdf";  lscfolder + "Scan_OCR_File.pdf";

            //    //Save the OCR processed PDF document in the disk
            //    lDoc.Save(lcNewFile);

            //    //Writes the text to the file
            //    File.WriteAllText(oldnewestFile.DirectoryName + "\\" + value + ".txt", text);  //  lscfolder + "ExtractedText.txt"

            //    //Close the document
            //    lDoc.Close(true);
            //}
            ////This will open the PDF file so, the result will be seen in default PDF viewer
            ////  Process.Start("OCR.pdf");

            //string line = null;
            //TextReader readFile = new StreamReader(oldnewestFile.DirectoryName + "\\" + value + ".txt");
            //line = readFile.ReadToEnd();
            //// MessageBox.Show(line);
            //readFile.Close();
            //readFile = null;

        }


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

            updateCalculations();

            string lcYear = txtYear.Text.Trim();
            string lcEOW = txtWeek.Text.Trim();
            string lcNetSales = txtNetSales.Text.Trim();
            string lcHealth = txtRetire.Text.Trim();
            string lcRetire = txtHealth.Text.Trim();

            string lcfPrimSupp = txtPrimSup.Text.Trim();
            string lcfOthSupp = txtOtherSupp.Text.Trim();
            string lcfBread = txtBread.Text.Trim();
            string lcfBev = txtBev.Text.Trim();
            string lcfProd = txtProd.Text.Trim();
            string lcfCarbon = txtCarbDio.Text.Trim();
            string lcfTotFood = txtFoodTot.Text.Trim();

            string lcoMort = txtMortgage.Text.Trim();
            string lcoLoan = txtLoan.Text.Trim();
            string lcoAssoc = txtAssociation.Text.Trim();
            string lcoPropTax = txtPropTax.Text.Trim();
            string lcoAdvCoop = txtAdvCoop.Text.Trim();
            string lcoNatAdver = txtNationalAdv.Text.Trim();
            string lcoLicenseFee = txtLicenseFee.Text.Trim();
            string lcoTotOverhead = txtTotOverhead.Text.Trim();

            string lceAccount = txtAccount.Text.Trim();
            string lceBank = txtBank.Text.Trim();
            string lceCC = txtCC.Text.Trim();
            string lceFuel = txtFuel.Text.Trim();
            string lceLegal = txtLegal.Text.Trim();
            string lceLicensePerm = txtLicense.Text.Trim();
            string lcePayroll = txtPayroll.Text.Trim();
            string lceInsur = txtInsur.Text.Trim();
            string lceWorkComp = txtWorkComp.Text.Trim();
            string lceAdvertise = txtAdvertising.Text.Trim();
            string lceCharitable = txtCharitableComp.Text.Trim();
            string lceAuto = txtAuto.Text.Trim();
            string lceCash = txtCashShort.Text.Trim();
            string lceElect = txtElectrical.Text.Trim();
            string lceGeneral = txtGeneral.Text.Trim();
            string lceHVAC = txtHVAC.Text.Trim();
            string lceLawn = txtLawn.Text.Trim();
            string lcePaint = txtPaint.Text.Trim();
            string lcePlumb = txtPlumb.Text.Trim();
            string lceRemodel = txtRemodel.Text.Trim();
            string lceStruct = txtStructural.Text.Trim();
            string lceDishMach = txtDishMach.Text.Trim();
            string lceJanitorial = txtJanitorial.Text.Trim();
            string lceOfficeComp = txtOffice.Text.Trim();
            string lceRestaurant = txtRestaurant.Text.Trim();
            string lceUniform = txtUniform.Text.Trim();
            string lceData = txtDataTele.Text.Trim();
            string lceElectric = txtElectricity.Text.Trim();
            string lceMusic = txtMusic.Text.Trim();
            string lceNatGas = txtNatGas.Text.Trim();
            string lceSecurity = txtSecurity.Text.Trim();
            string lceTrash = txtTrash.Text.Trim();
            string lceWaterSewer = txtWater.Text.Trim();
            string lceTotExpense = txtTotExpense.Text.Trim();

            string lclHost = txtHost.Text.Trim();
            string lclCook = txtCooks.Text.Trim();
            string lclServer = txtServers.Text.Trim();
            string lclDMO = txtDMO.Text.Trim();
            string lclSuperv = txtSupervisor.Text.Trim();
            string lclOvertime = txtOvertime.Text.Trim();
            string lclGenManager = txtGenManager.Text.Trim();
            string lclManager = txtManager.Text.Trim();
            string lclBonus = txtBonus.Text.Trim();
            string lclPayTax = txtPayrollTax.Text.Trim();
            string lclTotLabor = txtTotLabor.Text.Trim();

            string lcServer = "playgroup.database.windows.net";
            string lcODBC = "ODBC Driver 17 for SQL Server";
            string lcDB = "tb_HelpingHand";
            // string lcPort = "3306";  //  Port=" + lcPort + ";
            string lcUser = "tbmaster";
            string lcProv = "SQLOLEDB";
            string lcPass = "Smartman55";
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            OdbcConnection cnn = new OdbcConnection(lcConnectionString);

            cnn.Open();

            string lcSQL = "";
            lcSQL = "SELECT * from tb_datahold where Week='" + lcEOW + "'";      // lcSQL = "SELECT * from ~public~.~tb_Residents~ LIMIT 100".Replace('~', '"');
            OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            int result = cmd.ExecuteNonQuery();
            if (result > 0)
            {
                /// Update records
                // MessageBox.Show(result.ToString());
                lcSQL = " Update tb_datahold set NetSales=@lcNetSales, PrimSupp=@lcfPrimSupp, OthSupp=@lcfOthSupp, Bread=@lcfBread, Beverage=@lcfBev," +
                    " Produce=@lcfProd,CarbonDioxide=@lcfCarbon, FoodCost=@lcfTotFood, HostCashier=@lclHost, Cooks=@lclCook, Servers=@lclServer," +
                    " DMO=@lclDMO, Supervisor=@lclSuperv, Overtime=@lclOvertime,GeneralManager=@lclGenManager, Manager=@lclManager, Bonus=@lclBonus," +
                    " PayrollTax=@lclPayTax, Healthcare=@lcHealth, Retirement=@lcRetire, LaborCost=@lclTotLabor, Accounting=@lceAccount,Bank=@lceBank, CreditCard=@lceCC," +
                    " Fuel=@lceFuel, Legal=@lceLegal, License=@lceLicensePerm, PayrollProc=@lcePayroll, Insurance=@lceInsur,WorkersComp=@lceWorkComp," +
                    " Advertising=@lceAdvertise, Charitable=@lceCharitable, Auto=@lceAuto, CashShortage=@lceCash, Electrical=@lceElect,General=@lceGeneral," +
                    " HVAC=@lceHVAC, Lawn=@lceLawn, Painting=@lcePaint, Plumbing=@lcePlumb, Remodeling=@lceRemodel, Structural=@lceStruct," +
                    " DishMachine=@lceDishMach,Janitorial=@lceJanitorial, Office=@lceOfficeComp, Restaurant=@lceRestaurant, Uniforms=@lceUniform," +
                    " Data=@lceData, Electricity=@lceElectric,Music=@lceMusic, NaturalGas=@lceNatGas, Security=@lceSecurity, Trash=@lceTrash," +
                    " WaterSewer=@lceWaterSewer, ExpenseCost=@lceTotExpense, Mortgage=@lcoMort,LoanPayment=@lcoLoan, Association=@lcoAssoc," +
                    " PropertyTax=@lcoPropTax, AdvertisingCoop=@lcoAdvCoop, NationalAdvertise=@lcoNatAdver, LicensingFee=@lcoLicenseFee," +
                    "OverheadCost=@lcoTotOverhead where Week='@lcEOW'";
            }
            else
            {
                /// Insert records
                // MessageBox.Show("Hello There, no records");
                /// ,IDs
                lcSQL = " Insert into tb_datahold (Week,NetSales,PrimSupp,OthSupp,Bread,Beverage,Produce,CarbonDioxide,FoodCost,HostCashier,Cooks,Servers,DMO,Supervisor," +
                    "Overtime,GeneralManager,Manager,Bonus,PayrollTax,Healthcare,Retirement,LaborCost,Accounting,Bank,CreditCard,Fuel,Legal,License,PayrollProc," +
                    "Insurance,WorkersComp,Advertising,Charitable,Auto,CashShortage,Electrical,General,HVAC,Lawn,Painting,Plumbing,Remodeling,Structural,DishMachine," +
                    "Janitorial,Office,Restaurant,Uniforms,Data,Electricity,Music,NaturalGas,Security,Trash,WaterSewer,ExpenseCost,Mortgage,LoanPayment,Association," +
                    "PropertyTax,AdvertisingCoop,NationalAdvertise,LicensingFee,OverheadCost) " +
                    " values " +
                    " ('@lcEOW',@lcNetSales,@lcfPrimSupp,@lcfOthSupp,@lcfBread,@lcfBev,@lcfProd,@lcfCarbon,@lcfTotFood,@lclHost,@lclCook,@lclServer,@lclDMO," +
                    "@lclSuperv,@lclOvertime,@lclGenManager,@lclManager,@lclBonus,@lclPayTax,@lcHealth,@lcRetire,@lclTotLabor,@lceAccount,@lceBank,@lceCC," +
                    "@lceFuel,@lceLegal,@lceLicensePerm,@lcePayroll,@lceInsur,@lceWorkComp,@lceAdvertise,@lceCharitable,@lceAuto,@lceCash,@lceElect,@lceGeneral," +
                    "@lceHVAC,@lceLawn,@lcePaint,@lcePlumb,@lceRemodel,@lceStruct,@lceDishMach,@lceJanitorial,@lceOfficeComp,@lceRestaurant,@lceUniform,@lceData," +
                    "@lceElectric,@lceMusic,@lceNatGas,@lceSecurity,@lceTrash,@lceWaterSewer,@lceTotExpense,@lcoMort,@lcoLoan,@lcoAssoc,@lcoPropTax," +
                    "@lcoAdvCoop,@lcoNatAdver,@lcoLicenseFee,@lcoTotOverhead)";
            }

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
            cmd.Parameters.AddWithValue("@lcHealth", lcHealth);  //  HealthCare= 
            cmd.Parameters.AddWithValue("@lcRetire", lcRetire);  //  Retire=
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

        //public static string lConn(OdbcConnection conn)
        //{

        //    //string lcServer = "salt.db.elephantsql.com";
        //    //string lcODBC = "PostgreSQL ANSI";
        //    //string lcDB = "pffejyte";
        //    //// string lcPort = "5432";  //  Port=" + lcPort + ";
        //    //string lcUser = "pffejyte";
        //    //string lcPass = "Or2m-sdyDidrOWGaXBD--8b1-itKL92b";
        //    //string lcSQL = "";
        //    //string lcConnectionString = "Driver={" + lcODBC + "};Provider=SQLOLEDB;Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";

        //    //string lcServer = "67.222.39.62";
        //    //string lcODBC = "PostgreSQL ANSI";
        //    //string lcDB = "Tb_Test";
        //    //string lcPort = "3306";  //  Port=" + lcPort + ";
        //    //string lcUser = "dynamkr0_pgtest";
        //    //string lcProv = "SQLOLEDB";
        //    //string lcPass = "fzk4pktb";

        //    /// (New) tb_Play
        //    /// tb_HelpingHand
        //    /// playgroup
        //    /// tbmaster
        //    /// Smartman55
        //    /// (new) playgroup ((US) East US)
        //    /// https://hadoop.apache.org/
        //    /// https://www.digitalocean.com/

        //    string lcServer = "playgroup.database.windows.net";
        //    string lcODBC = "ODBC Driver 17 for SQL Server";
        //    string lcDB = "tb_HelpingHand";
        //    // string lcPort = "3306";  //  Port=" + lcPort + ";
        //    string lcUser = "tbmaster";
        //    string lcProv = "SQLOLEDB";
        //    string lcPass = "Smartman55";
        //    string lcSQL = "";
        //    string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
        //    OdbcConnection cnn = new OdbcConnection(lcConnectionString);

        //    // return cnn;
        //}




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
                    txtInvHold.Text = "EXPENSES";
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
            // Add columns together
            decimal totalSalary = 0;
            decimal amt = 0;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                var value = dataGridView1.Rows[i].Cells[4].Value;
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

            string lcServer = "dynamicelements.database.windows.net";
            string lcODBC = "ODBC Driver 17 for SQL Server";
            string lcDB = "dynamicelements";
            // string lcPort = "3306";  //  Port=" + lcPort + ";
            string lcUser = "tbmaster";
            string lcProv = "SQLOLEDB";
            string lcPass = "Fzk4pktb";     // Smartman55  Fzk4pktb
            string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
            OdbcConnection cnn = new OdbcConnection(lcConnectionString);
            cnn.Open();

            string lcSQL = "SELECT * from dynamicelements..vw_OrderLogs where Week='" + nextSunday + "'";
            OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            OdbcDataReader reader = cmd.ExecuteReader();

            if (reader.HasRows)
            {

            }
            else
            {

            }

            cnn.Close();

        }
    }

    //public class Conn_cl
    //{

    //    //string lcServer = "playgroup.database.windows.net";
    //    //string lcODBC = "ODBC Driver 17 for SQL Server";
    //    //string lcDB = "tb_HelpingHand";
    //    //string lcUser = "tbmaster";
    //    //string lcProv = "SQLOLEDB";
    //    //string lcPass = "Smartman55";
    //    //string lcConnectionString = "Driver={" + lcODBC + "};Provider=" + lcProv + ";Server=" + lcServer + ";DATABASE=" + lcDB + ";Uid=" + lcUser + "; Pwd=" + lcPass + ";";
    //    public static string lcConnectionString = "Driver={ODBC Driver 17 for SQL Server};Provider=SQLOLEDB;Server=playgroup.database.windows.net;DATABASE=tb_HelpingHand;Uid=tbmaster; Pwd=Smartman55;";
    //    public static OdbcConnection con;

    //    public static void OpenConection()
    //    {
    //        // string lcConnectionString = "Driver={ODBC Driver 17 for SQL Server};Provider=SQLOLEDB;Server=playgroup.database.windows.net;DATABASE=tb_HelpingHand;Uid=tbmaster; Pwd=Smartman55;";
    //        // OdbcConnection con;
    //        con = new OdbcConnection(lcConnectionString);
    //        con.Open();
    //    }
    //    public static void CloseConnection()
    //    {
    //        con.Close();
    //    }
    //    public static void ExecuteQueries(string Query_)
    //    {
    //        OdbcCommand cmd = new OdbcCommand(Query_, con);
    //        cmd.ExecuteNonQuery();
    //    }
    //    public static OdbcDataReader DataReader(string Query_)  // SqlDataReader
    //    {
    //        OdbcCommand cmd = new OdbcCommand(Query_, con);
    //        OdbcDataReader dr = cmd.ExecuteReader();  // SqlDataReader
    //        return dr;
    //    }
    //    public static object ShowDataInGridView(string Query_)
    //    {
    //        SqlDataAdapter dr = new SqlDataAdapter(Query_, lcConnectionString);  // SqlDataAdapter  SqlDataAdapter
    //        DataSet ds = new DataSet();
    //        dr.Fill(ds);
    //        object dataum = ds.Tables[0];
    //        return dataum;
    //    }
    //}
}
