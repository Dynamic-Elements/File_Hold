using Syncfusion.OCRProcessor;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tesseract;
using VFPToolkit;
using WIA;
using Excel = Microsoft.Office.Interop.Excel;


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

            ListScanners();

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

            var date = DateTime.Now;
            var lastSunday = Dates.DTOC(date.AddDays(-(int)date.DayOfWeek));  // Grabs the past Sunday for Week End

            textBox1.Text = lastSunday;
            textBox2.Text = lastSunday.Substring(lastSunday.Length - 4, 4);   // Yr.Substring(0,4);

            string lcSQL = "SELECT * from tb_HelpingHand..tb_datahold where Week='" + textBox1.Text.Trim() + "'";   // Week='" + textBox1.Text.Trim() + "'";   '12/30/2018'
            OdbcCommand cmd = new OdbcCommand(lcSQL, cnn);
            OdbcDataReader reader = cmd.ExecuteReader();
            // MessageBox.Show(Convert.ToString(reader.GetOrdinal("NetSales")));

            if (reader.HasRows)
            {

                textBox3.Text = reader["NetSales"].ToString();
                textBox8.Text = reader["Healthcare"].ToString();
                textBox9.Text = reader["Retirement"].ToString();

                textBox84.Text = reader["PrimSupp"].ToString();
                textBox77.Text = reader["OthSupp"].ToString();
                textBox76.Text = reader["Bread"].ToString();
                textBox75.Text = reader["Beverage"].ToString();
                textBox69.Text = reader["Produce"].ToString();
                textBox68.Text = reader["CarbonDioxide"].ToString();
                textBox4.Text = reader["FoodCost"].ToString();

                textBox83.Text = reader["Mortgage"].ToString();
                textBox82.Text = reader["LoanPayment"].ToString();
                textBox81.Text = reader["Association"].ToString();
                textBox80.Text = reader["PropertyTax"].ToString();
                textBox79.Text = reader["AdvertisingCoop"].ToString();
                textBox78.Text = reader["NationalAdvertise"].ToString();
                textBox73.Text = reader["LicensingFee"].ToString();
                textBox6.Text = reader["OverheadCost"].ToString();

                textBox27.Text = reader["Accounting"].ToString();
                textBox26.Text = reader["Bank"].ToString();
                textBox25.Text = reader["CreditCard"].ToString();
                textBox24.Text = reader["Fuel"].ToString();
                textBox23.Text = reader["Legal"].ToString();
                textBox22.Text = reader["License"].ToString();
                textBox28.Text = reader["PayrollProc"].ToString();
                textBox30.Text = reader["Insurance"].ToString();
                textBox29.Text = reader["WorkersComp"].ToString();
                textBox32.Text = reader["Advertising"].ToString();
                textBox31.Text = reader["Charitable"].ToString();
                textBox21.Text = reader["Auto"].ToString();
                textBox20.Text = reader["CashShortage"].ToString();
                textBox34.Text = reader["Electrical"].ToString();
                textBox33.Text = reader["General"].ToString();
                textBox19.Text = reader["HVAC"].ToString();
                textBox35.Text = reader["Lawn"].ToString();
                textBox36.Text = reader["Painting"].ToString();
                textBox37.Text = reader["Plumbing"].ToString();
                textBox38.Text = reader["Remodeling"].ToString();
                textBox39.Text = reader["Structural"].ToString();
                textBox43.Text = reader["DishMachine"].ToString();
                textBox42.Text = reader["Janitorial"].ToString();
                textBox44.Text = reader["Office"].ToString();
                textBox41.Text = reader["Restaurant"].ToString();
                textBox40.Text = reader["Uniforms"].ToString();
                textBox18.Text = reader["Data"].ToString();
                textBox45.Text = reader["Electricity"].ToString();
                textBox46.Text = reader["Music"].ToString();
                textBox47.Text = reader["NaturalGas"].ToString();
                textBox48.Text = reader["Security"].ToString();
                textBox49.Text = reader["Trash"].ToString();
                textBox50.Text = reader["WaterSewer"].ToString();
                textBox7.Text = reader["ExpenseCost"].ToString();

                textBox90.Text = reader["HostCashier"].ToString();
                textBox89.Text = reader["Cooks"].ToString();
                textBox88.Text = reader["Servers"].ToString();
                textBox87.Text = reader["DMO"].ToString();
                textBox86.Text = reader["Supervisor"].ToString();
                textBox85.Text = reader["Overtime"].ToString();
                textBox74.Text = reader["GeneralManager"].ToString();
                textBox72.Text = reader["Manager"].ToString();
                textBox71.Text = reader["Bonus"].ToString();
                textBox70.Text = reader["PayrollTax"].ToString();
                textBox5.Text = reader["LaborCost"].ToString();

                //  Retirement   Healthcare
            }
            else
            {

                textBox3.Text = "0.00";
                textBox8.Text = "0.00";
                textBox9.Text = "0.00";

                textBox84.Text = "0.00";
                textBox77.Text = "0.00";
                textBox76.Text = "0.00";
                textBox75.Text = "0.00";
                textBox69.Text = "0.00";
                textBox68.Text = "0.00";
                textBox4.Text = "0.00";

                textBox83.Text = "0.00";
                textBox82.Text = "0.00";
                textBox81.Text = "0.00";
                textBox80.Text = "0.00";
                textBox79.Text = "0.00";
                textBox78.Text = "0.00";
                textBox73.Text = "0.00";
                textBox6.Text = "0.00";

                textBox27.Text = "0.00";
                textBox26.Text = "0.00";
                textBox25.Text = "0.00";
                textBox24.Text = "0.00";
                textBox23.Text = "0.00";
                textBox22.Text = "0.00";
                textBox28.Text = "0.00";
                textBox30.Text = "0.00";
                textBox29.Text = "0.00";
                textBox32.Text = "0.00";
                textBox31.Text = "0.00";
                textBox21.Text = "0.00";
                textBox20.Text = "0.00";
                textBox34.Text = "0.00";
                textBox33.Text = "0.00";
                textBox19.Text = "0.00";
                textBox35.Text = "0.00";
                textBox36.Text = "0.00";
                textBox37.Text = "0.00";
                textBox38.Text = "0.00";
                textBox39.Text = "0.00";
                textBox43.Text = "0.00";
                textBox42.Text = "0.00";
                textBox44.Text = "0.00";
                textBox41.Text = "0.00";
                textBox40.Text = "0.00";
                textBox18.Text = "0.00";
                textBox45.Text = "0.00";
                textBox46.Text = "0.00";
                textBox47.Text = "0.00";
                textBox48.Text = "0.00";
                textBox49.Text = "0.00";
                textBox50.Text = "0.00";
                textBox7.Text = "0.00";

                textBox90.Text = "0.00";
                textBox89.Text = "0.00";
                textBox88.Text = "0.00";
                textBox87.Text = "0.00";
                textBox86.Text = "0.00";
                textBox85.Text = "0.00";
                textBox74.Text = "0.00";
                textBox72.Text = "0.00";
                textBox71.Text = "0.00";
                textBox70.Text = "0.00";
                textBox5.Text = "0.00";

            }
            cnn.Close();

        }

        /// <summary>
        /// This will loop through to find all the connected Printers & Scanners
        /// </summary>
        private void ListScanners()
        {
            // Clear the ListBox.
            comboBox1.Items.Clear();

            // Create a DeviceManager instance
            var deviceManager = new DeviceManager();

            int ivv = deviceManager.DeviceInfos.Count;

            MessageBox.Show(ivv.ToString());

            // Loop through the list of devices and add the name to the listbox
            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                // Add the device only if it's a scanner
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                {
                    continue;
                }

                // Add the Scanner device to the listbox (the entire DeviceInfos object)
                // Important: we store an object of type scanner (which ToString method returns the name of the scanner)
                comboBox1.Items.Add(new Scanner(deviceManager.DeviceInfos[i]));
            }

            if (comboBox1.Items.Count > 0)
                comboBox1.SelectedIndex = 0;

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
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // xlWorkSheet.Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(1);
            //  xlWorkBook.Worksheets.Add();

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

                //  coll[i].Cells[1, 10] = "Dates";
                //  coll[i].Range["j1:k1"].Merge();
                //  coll[i].Cells[2, 10] = "Week 5";
                //  coll[i].Cells[3, 10] = "$";
                //  coll[i].Cells[3, 11] = "%";

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

            /// Notes
            /// https://www.syncfusion.com/kb/9144/how-to-convert-scanned-image-to-searchable-pdf-by-processing-ocr
            /// https://help.syncfusion.com/file-formats/pdf/working-with-ocr?_ga=2.194924142.216447619.1574224028-344549646.1574224028 
            ///  
            /// 
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
            /// https://docs.microsoft.com/en-us/sql/azure-data-studio/quickstart-postgres?view=sql-server-ver15
            /// https://www.mssqltips.com/sqlservertip/3662/sql-server-and-postgresql-linked-server-configuration--part-2/
            /// https://www.vintasoft.com/vsocr-dotnet-index.html

            /// https://docs.microsoft.com/en-us/azure/sql-database/?view=sql-server-ver15
            /// https://www.vintasoft.com/docs/vsimaging-dotnet/Programming-Ocr-Save_OCR_results.html#SaveOcrResultsToTextFile
            /// file:///C:/Program%20Files/gs/gs9.50/doc/Readme.htm
            /// 


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

            // https://ourcodeworld.com/articles/read/382/creating-a-scanning-application-in-winforms-with-csharp
            // https://ithoughthecamewithyou.com/post/scanning-from-the-adf-using-wia-in-c

            //  https://csharp.hotexamples.com/examples/-/Tesseract/-/php-tesseract-class-examples.html

            // Use scanner/Printer
            Scanner device = null;

            this.Invoke(new MethodInvoker(delegate ()
            {
                device = comboBox1.SelectedItem as Scanner;
            }));

            if (device == null)
            {
                MessageBox.Show("You need to select first an scanner device from the list",
                                "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Provide a filename",
                                "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ImageFile image = new ImageFile();
            image = device.ScanJPEG();

            // Testing Random number for multiple runs
            var rand = new Random();

            // Save the image
            var path = lscfolder + "ScanFile"+ rand.Next(10,100) + ".jpeg";

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            image.SaveFile(path);

            string lcNewFile = lscfolder + "Scan_OCR_File" + rand.Next(10, 100) + ".pdf";

            //Create a new PDF document
            PdfDocument document = new PdfDocument();
            //Add a page to the document
            PdfPage page = document.Pages.Add();
            //Create PDF graphics for a page
            PdfGraphics graphics = page.Graphics;
            //Load the image from the disk
            PdfBitmap image1 = new PdfBitmap(path);   //  "Input.jpg"
            //Draw the image
            graphics.DrawImage(image1, 0, 0, page.GetClientSize().Width, page.GetClientSize().Height);
            //Save the document into stream
            MemoryStream stream = new MemoryStream();
            document.Save(stream);
            //Initialize the OCR processor by providing the path of tesseract binaries(SyncfusionTesseract.dll and liblept168.dll)
            using (OCRProcessor processor = new OCRProcessor(@"../../Tesseract Binaries/"))
            {
                //Load a PDF document
                PdfLoadedDocument lDoc = new PdfLoadedDocument(stream);

                //Set OCR language to process
                processor.Settings.Language = Languages.English;

                //Enable the AutoDetectRotation
                processor.Settings.AutoDetectRotation = true;

                //Enable native call  
                processor.Settings.EnableNativeCall = true;

                //Process OCR by providing the PDF document and Tesseract data
                String text = processor.PerformOCR(lDoc, @"..\..\Tessdata\");

                //Save the OCR processed PDF document in the disk
                lDoc.Save(lcNewFile);

                //Writes the text to the file
                File.WriteAllText(lscfolder + "ExtractedText.txt", text);

                //Close the document
                lDoc.Close(true);
            }
            //This will open the PDF file so, the result will be seen in default PDF viewer
            //  Process.Start("OCR.pdf");

            string line = null;
            TextReader readFile = new StreamReader(lscfolder + "ExtractedText.txt");
            line = readFile.ReadToEnd();
            MessageBox.Show(line);
            readFile.Close();
            readFile = null;


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

            updateCalculations();

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

            updateCalculations();

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

            updateCalculations();

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

            updateCalculations();

        }



        private void updateCalculations()
        {
            // This will calculate all the totals of each grouping
            // textBox6.Text = Convert.ToString((Convert.ToInt32(textBox1.Text) + Convert.ToInt32()));

            try
            {
                // Food
                textBox4.Text = Convert.ToString(Convert.ToInt32(textBox84.Text) + Convert.ToInt32(textBox77.Text) + Convert.ToInt32(textBox76.Text) +
                    Convert.ToInt32(textBox75.Text) + Convert.ToInt32(textBox69.Text) + Convert.ToInt32(textBox68.Text));

                // Expenses
                textBox7.Text = Convert.ToString((Convert.ToInt32(textBox27.Text) + Convert.ToInt32(textBox26.Text) + Convert.ToInt32(textBox25.Text) +
                    Convert.ToInt32(textBox24.Text) + Convert.ToInt32(textBox23.Text) + Convert.ToInt32(textBox22.Text) + Convert.ToInt32(textBox28.Text) +
                    Convert.ToInt32(textBox30.Text) + Convert.ToInt32(textBox29.Text) + Convert.ToInt32(textBox32.Text) + Convert.ToInt32(textBox31.Text) +
                    Convert.ToInt32(textBox21.Text) + Convert.ToInt32(textBox20.Text) + Convert.ToInt32(textBox34.Text) + Convert.ToInt32(textBox33.Text) +
                    Convert.ToInt32(textBox19.Text) + Convert.ToInt32(textBox35.Text) + Convert.ToInt32(textBox36.Text) + Convert.ToInt32(textBox37.Text) +
                    Convert.ToInt32(textBox38.Text) + Convert.ToInt32(textBox39.Text) + Convert.ToInt32(textBox43.Text) + Convert.ToInt32(textBox42.Text) +
                    Convert.ToInt32(textBox44.Text) + Convert.ToInt32(textBox41.Text) + Convert.ToInt32(textBox40.Text) + Convert.ToInt32(textBox18.Text) +
                    Convert.ToInt32(textBox45.Text) + Convert.ToInt32(textBox46.Text) + Convert.ToInt32(textBox47.Text) + Convert.ToInt32(textBox48.Text) +
                    Convert.ToInt32(textBox49.Text) + Convert.ToInt32(textBox50.Text)));

                // Labor
                textBox5.Text = Convert.ToString(Convert.ToInt32(textBox90.Text) + Convert.ToInt32(textBox89.Text) + Convert.ToInt32(textBox88.Text) +
                    Convert.ToInt32(textBox87.Text) + Convert.ToInt32(textBox86.Text) + Convert.ToInt32(textBox85.Text) + Convert.ToInt32(textBox74.Text) +
                    Convert.ToInt32(textBox72.Text) + Convert.ToInt32(textBox71.Text) + Convert.ToInt32(textBox70.Text));

                // Overhead
                textBox6.Text = Convert.ToString(Convert.ToInt32(textBox83.Text) + Convert.ToInt32(textBox82.Text) + Convert.ToInt32(textBox81.Text) +
                    Convert.ToInt32(textBox80.Text) + Convert.ToInt32(textBox79.Text) + Convert.ToInt32(textBox78.Text) + Convert.ToInt32(textBox73.Text));
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

            string lcYear = textBox2.Text.Trim();
            string lcEOW = textBox1.Text.Trim();
            string lcNetSales = textBox3.Text.Trim();
            string lcHealth = textBox8.Text.Trim();
            string lcRetire = textBox9.Text.Trim();

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
            string value = textBox3.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox3.Text = val.ToString("C");
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox84_Leave(object sender, EventArgs e)
        {
            string value = textBox84.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox84.Text = val.ToString("C");
        }

        private void textBox84_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox77_Leave(object sender, EventArgs e)
        {
            string value = textBox77.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox77.Text = val.ToString("C");
        }

        private void textBox77_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox76_Leave(object sender, EventArgs e)
        {
            string value = textBox76.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox76.Text = val.ToString("C");
        }

        private void textBox76_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox75_Leave(object sender, EventArgs e)
        {
            string value = textBox75.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox75.Text = val.ToString("C");
        }

        private void textBox75_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox69_Leave(object sender, EventArgs e)
        {
            string value = textBox69.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox69.Text = val.ToString("C");
        }

        private void textBox69_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox68_Leave(object sender, EventArgs e)
        {
            string value = textBox68.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox68.Text = val.ToString("C");
        }

        private void textBox68_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            string value = textBox4.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox4.Text = val.ToString("C");
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox90_Leave(object sender, EventArgs e)
        {
            string value = textBox90.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox90.Text = val.ToString("C");
        }

        private void textBox90_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox89_Leave(object sender, EventArgs e)
        {
            string value = textBox89.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox89.Text = val.ToString("C");
        }

        private void textBox89_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox88_Leave(object sender, EventArgs e)
        {
            string value = textBox88.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox88.Text = val.ToString("C");
        }

        private void textBox88_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox87_Leave(object sender, EventArgs e)
        {
            string value = textBox87.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox87.Text = val.ToString("C");
        }

        private void textBox87_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox86_Leave(object sender, EventArgs e)
        {
            string value = textBox86.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox86.Text = val.ToString("C");
        }

        private void textBox86_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox85_Leave(object sender, EventArgs e)
        {
            string value = textBox85.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox85.Text = val.ToString("C");
        }

        private void textBox85_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox74_Leave(object sender, EventArgs e)
        {
            string value = textBox74.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox74.Text = val.ToString("C");
        }

        private void textBox74_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox72_Leave(object sender, EventArgs e)
        {
            string value = textBox72.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox72.Text = val.ToString("C");
        }

        private void textBox72_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox71_Leave(object sender, EventArgs e)
        {
            string value = textBox71.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox71.Text = val.ToString("C");
        }

        private void textBox71_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox70_Leave(object sender, EventArgs e)
        {
            string value = textBox70.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox70.Text = val.ToString("C");
        }

        private void textBox70_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {
            string value = textBox5.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox5.Text = val.ToString("C");
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox83_Leave(object sender, EventArgs e)
        {
            string value = textBox83.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox83.Text = val.ToString("C");
        }

        private void textBox83_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox82_Leave(object sender, EventArgs e)
        {
            string value = textBox82.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox82.Text = val.ToString("C");
        }

        private void textBox82_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox81_Leave(object sender, EventArgs e)
        {
            string value = textBox81.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox81.Text = val.ToString("C");
        }

        private void textBox81_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox80_Leave(object sender, EventArgs e)
        {
            string value = textBox80.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox80.Text = val.ToString("C");
        }

        private void textBox80_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox79_Leave(object sender, EventArgs e)
        {
            string value = textBox79.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox79.Text = val.ToString("C");
        }

        private void textBox79_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox78_Leave(object sender, EventArgs e)
        {
            string value = textBox78.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox78.Text = val.ToString("C");
        }

        private void textBox78_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox73_Leave(object sender, EventArgs e)
        {
            string value = textBox73.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox73.Text = val.ToString("C");
        }

        private void textBox73_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            string value = textBox6.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox6.Text = val.ToString("C");
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox27_Leave(object sender, EventArgs e)
        {
            string value = textBox27.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox27.Text = val.ToString("C");
        }

        private void textBox27_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox26_Leave(object sender, EventArgs e)
        {
            string value = textBox26.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox26.Text = val.ToString("C");
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox25_Leave(object sender, EventArgs e)
        {
            string value = textBox25.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox25.Text = val.ToString("C");
        }

        private void textBox25_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox24_Leave(object sender, EventArgs e)
        {
            string value = textBox24.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox24.Text = val.ToString("C");
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox23_Leave(object sender, EventArgs e)
        {
            string value = textBox23.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox23.Text = val.ToString("C");
        }

        private void textBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox22_Leave(object sender, EventArgs e)
        {
            string value = textBox22.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox22.Text = val.ToString("C");
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox28_Leave(object sender, EventArgs e)
        {
            string value = textBox28.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox28.Text = val.ToString("C");
        }

        private void textBox28_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox30_Leave(object sender, EventArgs e)
        {
            string value = textBox30.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox30.Text = val.ToString("C");
        }

        private void textBox30_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox29_Leave(object sender, EventArgs e)
        {
            string value = textBox29.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox29.Text = val.ToString("C");
        }

        private void textBox29_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox32_Leave(object sender, EventArgs e)
        {
            string value = textBox32.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox32.Text = val.ToString("C");
        }

        private void textBox32_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox31_Leave(object sender, EventArgs e)
        {
            string value = textBox31.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox31.Text = val.ToString("C");
        }

        private void textBox31_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox21_Leave(object sender, EventArgs e)
        {
            string value = textBox21.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox21.Text = val.ToString("C");
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox20_Leave(object sender, EventArgs e)
        {
            string value = textBox20.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox20.Text = val.ToString("C");
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox34_Leave(object sender, EventArgs e)
        {
            string value = textBox34.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox34.Text = val.ToString("C");
        }

        private void textBox34_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox33_Leave(object sender, EventArgs e)
        {
            string value = textBox33.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox33.Text = val.ToString("C");
        }

        private void textBox33_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox19_Leave(object sender, EventArgs e)
        {
            string value = textBox19.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox19.Text = val.ToString("C");
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox35_Leave(object sender, EventArgs e)
        {
            string value = textBox35.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox35.Text = val.ToString("C");
        }

        private void textBox35_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox36_Leave(object sender, EventArgs e)
        {
            string value = textBox36.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox36.Text = val.ToString("C");
        }

        private void textBox36_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox37_Leave(object sender, EventArgs e)
        {
            string value = textBox37.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox37.Text = val.ToString("C");
        }

        private void textBox37_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox38_Leave(object sender, EventArgs e)
        {
            string value = textBox38.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox38.Text = val.ToString("C");
        }

        private void textBox38_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox39_Leave(object sender, EventArgs e)
        {
            string value = textBox39.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox39.Text = val.ToString("C");
        }

        private void textBox39_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox43_Leave(object sender, EventArgs e)
        {
            string value = textBox43.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox43.Text = val.ToString("C");
        }

        private void textBox43_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox42_Leave(object sender, EventArgs e)
        {
            string value = textBox42.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox42.Text = val.ToString("C");
        }

        private void textBox42_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox44_Leave(object sender, EventArgs e)
        {
            string value = textBox44.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox44.Text = val.ToString("C");
        }

        private void textBox44_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox41_Leave(object sender, EventArgs e)
        {
            string value = textBox41.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox41.Text = val.ToString("C");
        }

        private void textBox41_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox40_Leave(object sender, EventArgs e)
        {
            string value = textBox40.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox40.Text = val.ToString("C");
        }

        private void textBox40_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox18_Leave(object sender, EventArgs e)
        {
            string value = textBox18.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox18.Text = val.ToString("C");
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox45_Leave(object sender, EventArgs e)
        {
            string value = textBox45.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox45.Text = val.ToString("C");
        }

        private void textBox45_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox46_Leave(object sender, EventArgs e)
        {
            string value = textBox46.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox46.Text = val.ToString("C");
        }

        private void textBox46_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox47_Leave(object sender, EventArgs e)
        {
            string value = textBox47.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox47.Text = val.ToString("C");
        }

        private void textBox47_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox48_Leave(object sender, EventArgs e)
        {
            string value = textBox48.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox48.Text = val.ToString("C");
        }

        private void textBox48_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox49_Leave(object sender, EventArgs e)
        {
            string value = textBox49.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox49.Text = val.ToString("C");
        }

        private void textBox49_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox50_Leave(object sender, EventArgs e)
        {
            string value = textBox50.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox50.Text = val.ToString("C");
        }

        private void textBox50_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            string value = textBox7.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox7.Text = val.ToString("C");
        }

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            string value = textBox8.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox8.Text = val.ToString("C");
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && e.KeyChar != (char)8 && e.KeyChar != (char)46;  // 8 is backspace, 46 is period
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            string value = textBox9.Text.Replace(",", "").Replace("$", "");
            decimal val;
            if (decimal.TryParse(value, out val))
                textBox9.Text = val.ToString("C");
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
