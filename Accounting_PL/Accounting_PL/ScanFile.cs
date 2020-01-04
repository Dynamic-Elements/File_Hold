using System;
using System.Drawing;
using System.Drawing.Imaging;
using WIA;

namespace ScanIt
{
    public struct PageSize
    {
        public double Height;
        public double Width;

        public PageSize(double height, double width)
        {
            this.Height = height;
            this.Width = width;
        }
    }

    public class WiaWrapper
    {
        //Standard Page Sizes  --  Height x Width (in)
        public PageSize A0 = new PageSize(46.8, 33.1);

        public PageSize A1 = new PageSize(33.1, 23.4);

        public PageSize A2 = new PageSize(23.4, 16.5);

        public PageSize A3 = new PageSize(16.5, 11.7);

        public PageSize A4 = new PageSize(11.7, 8.3);

        public PageSize A5 = new PageSize(8.3, 5.8);

        public PageSize A6 = new PageSize(5.8, 4.1);

        public PageSize A7 = new PageSize(4.1, 2.9);

        public PageSize A8 = new PageSize(2.9, 2.0);

        public PageSize A9 = new PageSize(2.0, 1.5);

        public PageSize A10 = new PageSize(1.5, 1.0);

        public string DeviceID;

        //Image Filenames
        private const string wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}";

        private const string wiaFormatGIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}";
        private const string wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}";
        private const string wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}";
        private const string wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}";

        #region Setup/select Scanner

        /// <summary>
        /// Select Scanner.
        /// If you need to save the Scanner, Save WiaWrapper.DeviceID
        /// </summary>
        public void SelectScanner()
        {
            WIA.CommonDialog wiaDiag = new WIA.CommonDialog();

            try
            {
                Device d = wiaDiag.ShowSelectDevice(WiaDeviceType.UnspecifiedDeviceType, true, false);
                if (d != null)
                {
                    DeviceID = d.DeviceID;
                    return;
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception("Error, Is a scanner chosen?");
            }

            throw new System.Exception("No Device Selected");
        }

        /// <summary>
        /// Connect to Scanning Device
        /// </summary>
        /// <param name="deviceID"></param>
        /// <returns></returns>
        private Device Connect()
        {
            try
            {
                Device WiaDev = null;

                DeviceManager manager = new DeviceManager();

                //Iterate through each Device until correct Device found
                foreach (DeviceInfo info in manager.DeviceInfos)
                {
                    if (info.DeviceID == DeviceID)
                    {
                        WIA.Properties infoprop = info.Properties;

                        WiaDev = info.Connect();
                        return WiaDev;
                    }
                }
            }
            catch (Exception ex)
            { }
            return null;
        }

        #endregion Setup/select Scanner

        #region Scanning utilities - hasMorePages, SetupPageSize, SetupADF, DeleteFile

        private void Delete_File(string filename)
        {
            //Overwrite File
            if (System.IO.File.Exists(filename))
            {
                //file exists, delete it
                System.IO.File.Delete(filename);
            }
        }

        /// <summary>
        /// Check to see if ADF has more pages loaded
        /// </summary>
        /// <param name="wia"></param>
        /// <returns></returns>
        private bool HasMorePages(Device wia)
        {
            //determine if there are any more pages waiting
            Property documentHandlingSelect = null;
            Property documentHandlingStatus = null;

            //  bool hasMorePages = true;   // maybe

            string test = string.Empty;

            foreach (Property prop in wia.Properties)
            {
                string propername = prop.Name;
                string propvalue = prop.get_Value().ToString();

                test += propername + " " + propvalue + "<br>";

                if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_SELECT)
                    documentHandlingSelect = prop;
                if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_STATUS)
                    documentHandlingStatus = prop;
            }

            //  hasMorePages = false; //  assume there are no more pages  // maybe
            if (documentHandlingSelect != null)
            //may not exist on flatbed scanner but required for feeder
            {
                //check for document feeder
                if ((Convert.ToUInt32(documentHandlingSelect.get_Value()) & 0x00000001) != 0)  // Convert.ToUInt32(documentHandlingSelect.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_SELECT.FEEDER) != 0  //  0x00000001
                {
                    return ((Convert.ToUInt32(documentHandlingStatus.get_Value()) & 0x00000001) != 0);  //  return ((Convert.ToUInt32(documentHandlingStatus.get_Value()) & 0x00000001) != 0)  WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) != 0
                }
            }

            string tester = test;

            return false;
        }

        /// <summary>
        /// Setup device to Use ADF if required
        /// </summary>
        private void SetupADF(Device wia, bool duplex)
        {
            string adf = string.Empty;

            foreach (WIA.Property deviceProperty in wia.Properties)
            {
                adf += deviceProperty.Name + "<br>";
                if (deviceProperty.Name == "Document Handling Select") //or PropertyID == 3088
                {
                    int value = duplex ? 0x004 : 0x001;
                    deviceProperty.set_Value(value);
                }
            }
        }

        /// <summary>
        /// Setup Page Size
        /// </summary>
        /// <param name="wia"></param>
        private void SetupPageSize(Device wia, bool rotatePage, PageSize pageSize, int DPI, WIA.Item item)
        {
            //Setup Page Size Property
            foreach (WIA.Property itemProperty in item.Properties)
            {
                if (itemProperty.Name.Equals("Horizontal Resolution"))
                {
                    ((IProperty)itemProperty).set_Value(DPI);
                }
                else if (itemProperty.Name.Equals("Vertical Resolution"))
                {
                    ((IProperty)itemProperty).set_Value(DPI);
                }
                else if (itemProperty.Name.Equals("Horizontal Extent"))
                {
                    double extent = DPI * pageSize.Height;

                    if (rotatePage)
                    {
                        extent = DPI * pageSize.Width;
                    }

                    ((IProperty)itemProperty).set_Value(extent);
                }
                else if (itemProperty.Name.Equals("Vertical Extent"))
                {
                    double extent = DPI * pageSize.Width;

                    if (rotatePage)
                    {
                        extent = pageSize.Height * DPI;
                    }

                    ((IProperty)itemProperty).set_Value(extent);
                }
            }
        }

        #endregion Scanning utilities - hasMorePages, SetupPageSize, SetupADF, DeleteFile

        // #region Scan Page - Main Public Method

        public Bitmap ConvertToBitonal(Bitmap original)
        {
            Bitmap source = null;

            // If original bitmap is not already in 32 BPP, ARGB format, then convert
            if (original.PixelFormat != PixelFormat.Format32bppArgb)
            {
                source = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb);
                source.SetResolution(original.HorizontalResolution, original.VerticalResolution);
                using (Graphics g = Graphics.FromImage(source))
                {
                    g.DrawImageUnscaled(original, 0, 0);
                }
            }
            else
            {
                source = original;
            }

            // Lock source bitmap in memory
            BitmapData sourceData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            // Copy image data to binary array
            int imageSize = sourceData.Stride * sourceData.Height;
            byte[] sourceBuffer = new byte[imageSize];
            System.Runtime.InteropServices.Marshal.Copy(sourceData.Scan0, sourceBuffer, 0, imageSize);

            // Unlock source bitmap
            source.UnlockBits(sourceData);

            // Create destination bitmap
            Bitmap destination = new Bitmap(source.Width, source.Height, PixelFormat.Format1bppIndexed);

            // Lock destination bitmap in memory
            BitmapData destinationData = destination.LockBits(new Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);

            // Create destination buffer
            imageSize = destinationData.Stride * destinationData.Height;
            byte[] destinationBuffer = new byte[imageSize];

            int sourceIndex = 0;
            int destinationIndex = 0;
            int pixelTotal = 0;
            byte destinationValue = 0;
            int pixelValue = 128;
            int height = source.Height;
            int width = source.Width;
            int threshold = 500;

            // Iterate lines
            for (int y = 0; y < height; y++)
            {
                sourceIndex = y * sourceData.Stride;
                destinationIndex = y * destinationData.Stride;
                destinationValue = 0;
                pixelValue = 128;

                // Iterate pixels
                for (int x = 0; x < width; x++)
                {
                    // Compute pixel brightness (i.e. total of Red, Green, and Blue values)
                    pixelTotal = sourceBuffer[sourceIndex + 1] + sourceBuffer[sourceIndex + 2] + sourceBuffer[sourceIndex + 3];
                    if (pixelTotal > threshold)
                    {
                        destinationValue += (byte)pixelValue;
                    }
                    if (pixelValue == 1)
                    {
                        destinationBuffer[destinationIndex] = destinationValue;
                        destinationIndex++;
                        destinationValue = 0;
                        pixelValue = 128;
                    }
                    else
                    {
                        pixelValue >>= 1;
                    }
                    sourceIndex += 4;
                }
                if (pixelValue != 128)
                {
                    destinationBuffer[destinationIndex] = destinationValue;
                }
            }

            // Copy binary image data to destination bitmap
            System.Runtime.InteropServices.Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, imageSize);

            // Unlock destination bitmap
            destination.UnlockBits(destinationData);

            // Return
            return destination;
        }

        public void mergeTiffPages(string str_DestinationPath, string[] sourceFiles)
        {
            System.Drawing.Imaging.ImageCodecInfo codec = null;

            foreach (System.Drawing.Imaging.ImageCodecInfo cCodec in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
            {
                if (cCodec.CodecName == "Built-in TIFF Codec")
                    codec = cCodec;
            }

            try
            {
                System.Drawing.Imaging.EncoderParameters imagePararms = new System.Drawing.Imaging.EncoderParameters(1);
                imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.MultiFrame);

                if (sourceFiles.Length == 1)
                {
                    System.IO.File.Copy((string)sourceFiles[0], str_DestinationPath, true);
                }
                else if (sourceFiles.Length >= 1)
                {
                    System.Drawing.Image DestinationImage = (System.Drawing.Image)(new System.Drawing.Bitmap((string)System.Configuration.ConfigurationSettings.AppSettings["Path"] + sourceFiles[0]));

                    DestinationImage.Save(str_DestinationPath, codec, imagePararms);

                    imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.FrameDimensionPage);

                    for (int i = 0; i < sourceFiles.Length - 1; i++)
                    {
                        System.Drawing.Image img = (System.Drawing.Image)(new System.Drawing.Bitmap((string)sourceFiles[i]));

                        DestinationImage.SaveAdd(img, imagePararms);
                        img.Dispose();
                    }

                    imagePararms.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)System.Drawing.Imaging.EncoderValue.Flush);
                    DestinationImage.SaveAdd(imagePararms);
                    imagePararms.Dispose();
                    DestinationImage.Dispose();
                }
            }
            catch (Exception ex)
            { }
        }

        public bool saveMultipage(Image[] bmp, string location, string type)
        {
            if (bmp != null)
            {
                try
                {
                    ImageCodecInfo codecInfo = getCodecForstring(type);

                    for (int i = 0; i < bmp.Length; i++)
                    {
                        if (bmp[i] == null)
                            break;
                        bmp[i] = (Image)ConvertToBitonal((Bitmap)bmp[i]);
                    }

                    if (bmp.Length == 1)
                    {
                        EncoderParameters iparams = new EncoderParameters(1);
                        System.Drawing.Imaging.Encoder iparam = System.Drawing.Imaging.Encoder.Compression;
                        EncoderParameter iparamPara = new EncoderParameter(iparam, (long)(EncoderValue.CompressionCCITT4));
                        iparams.Param[0] = iparamPara;
                        bmp[0].Save(location, codecInfo, iparams);
                    }
                    else if (bmp.Length > 1)
                    {
                        System.Drawing.Imaging.Encoder saveEncoder;
                        System.Drawing.Imaging.Encoder compressionEncoder;
                        EncoderParameter SaveEncodeParam;
                        EncoderParameter CompressionEncodeParam;
                        EncoderParameters EncoderParams = new EncoderParameters(2);

                        saveEncoder = System.Drawing.Imaging.Encoder.SaveFlag;
                        compressionEncoder = System.Drawing.Imaging.Encoder.Compression;

                        // Save the first page (frame).
                        SaveEncodeParam = new EncoderParameter(saveEncoder, (long)EncoderValue.MultiFrame);
                        CompressionEncodeParam = new EncoderParameter(compressionEncoder, (long)EncoderValue.CompressionCCITT4);
                        EncoderParams.Param[0] = CompressionEncodeParam;
                        EncoderParams.Param[1] = SaveEncodeParam;

                        System.IO.File.Delete(location);
                        bmp[0].Save(location, codecInfo, EncoderParams);

                        for (int i = 1; i < bmp.Length; i++)
                        {
                            if (bmp[i] == null)
                                break;

                            SaveEncodeParam = new EncoderParameter(saveEncoder, (long)EncoderValue.FrameDimensionPage);
                            CompressionEncodeParam = new EncoderParameter(compressionEncoder, (long)EncoderValue.CompressionCCITT4);
                            EncoderParams.Param[0] = CompressionEncodeParam;
                            EncoderParams.Param[1] = SaveEncodeParam;
                            bmp[0].SaveAdd(bmp[i], EncoderParams);
                        }

                        SaveEncodeParam = new EncoderParameter(saveEncoder, (long)EncoderValue.Flush);
                        EncoderParams.Param[0] = SaveEncodeParam;
                        bmp[0].SaveAdd(EncoderParams);
                    }
                    return true;
                }
                catch (System.Exception ee)
                {
                    throw new Exception(ee.Message + "  Error in saving as multipage ");
                }
            }
            else
                return false;
        }

        /// <summary>
        /// Scan Page,
        /// </summary>
        /// <param name="wia">Connected Device</param>
        /// <param name="pageSize">Page Size. A4, A3, A2 Etc</param>
        /// <param name="RotatePage">Rotation of page while scanning</param>
        public void Scan(bool rotatePage, int DPI, string filepath, bool useAdf, bool duplex)  //PageSize pageSize,
        {
            int pages = 0;
            bool hasMorePages = false;
            string[] sourceFiles = new string[100];

            WIA.CommonDialog WiaCommonDialog = new WIA.CommonDialog();

            try
            {
                do
                {
                    //  Connect to Device
                    Device wia = Connect();
                    WIA.Item item = wia.Items[1] as WIA.Item;

                    //  Setup ADF
                    if ((useAdf) || (duplex))
                        SetupADF(wia, duplex);

                    //  Setup Page Size
                    //  SetupPageSize(wia, rotatePage, A4, DPI, item);

                    WIA.ImageFile imgFile = null;
                    WIA.ImageFile imgFile_duplex = null; //  if duplex is setup, this will be back page

                    imgFile = (ImageFile)WiaCommonDialog.ShowTransfer(item, wiaFormatJPEG, false);

                    //  If duplex page, get back page now.
                    if (duplex)
                    {
                        imgFile_duplex = (ImageFile)WiaCommonDialog.ShowTransfer(item, wiaFormatJPEG, false);
                    }

                    string varImageFileName = filepath + "\\Scanned" + ".jpeg";

                    Delete_File(varImageFileName); //  if file already exists. delete it.
                    imgFile.SaveFile(varImageFileName);

                    using (var src = new System.Drawing.Bitmap(varImageFileName))
                    using (var bmp = new System.Drawing.Bitmap(1000, 1000, System.Drawing.Imaging.PixelFormat.Format32bppPArgb))
                    using (var gr = System.Drawing.Graphics.FromImage(bmp))
                    {
                        gr.Clear(System.Drawing.Color.Blue);
                        gr.DrawImage(src, new System.Drawing.Rectangle(0, 0, bmp.Width, bmp.Height));
                        gr.DrawString("This is vivek", new System.Drawing.Font("Arial", 15, System.Drawing.FontStyle.Regular), System.Drawing.SystemBrushes.WindowText, new System.Drawing.Point(550, 20));
                        bmp.Save(System.Configuration.ConfigurationSettings.AppSettings["Path"] + "test" + pages + ".jpeg", System.Drawing.Imaging.ImageFormat.Png);

                        ////string imgPath = "test" + pages + ".tiff";
                        ////sourceFiles[pages] = imgPath;
                        //mergeTiffPages(string str_DestinationPath, string[] sourceFiles)
                    }
                    mergeTiffPages(@"D:\Test\", sourceFiles);

                    string varImageFileName_duplex;

                    if (duplex)
                    {
                        varImageFileName_duplex = filepath + "\\Scanned-" + pages.ToString() + ".tiff";
                        Delete_File(varImageFileName_duplex); //if file already exists. delete it.
                        imgFile_duplex.SaveFile(varImageFileName);
                    }

                    //Check with scanner to see if there are more pages.
                    if (useAdf || duplex)
                    {
                        hasMorePages = HasMorePages(wia);
                        pages++;
                    }
                }

                while (hasMorePages);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                //throw new System.Exception(CheckError((uint)ex.ErrorCode));
            }
        }

        private ImageCodecInfo getCodecForstring(string type)
        {
            ImageCodecInfo[] info = ImageCodecInfo.GetImageEncoders();

            for (int i = 0; i < info.Length; i++)
            {
                string EnumName = type.ToString();
                if (info[i].FormatDescription.Equals(EnumName))
                {
                    return info[i];
                }
            }

            return null;
        }
    }

  
    internal class NewScanner
    {
    }

    class WIA_PROPERTIES
    {
        public const uint WIA_DIP_FIRST = 2;
        public const uint WIA_DPA_FIRST = WIA_DIP_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
        public const uint WIA_DPC_FIRST = WIA_DPA_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
        public const uint WIA_DPS_DOCUMENT_HANDLING_SELECT = WIA_DPS_FIRST + 14;
        public const uint WIA_DPS_DOCUMENT_HANDLING_STATUS = WIA_DPS_FIRST + 13;

        //
        // Scanner only device properties (DPS)
        //
        public const uint WIA_DPS_FIRST = WIA_DPC_FIRST + WIA_RESERVED_FOR_NEW_PROPS;

        public const uint WIA_RESERVED_FOR_NEW_PROPS = 1024;
    }
}