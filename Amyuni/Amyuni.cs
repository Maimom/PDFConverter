using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CDIntfEx;
using System.IO;
using System.Drawing.Printing;

namespace PDFConverter
{
    public class Amyuni
    {
        //create PDF printer constants	
        const string PDFprinter = "Amyuni PDF Converter";
        const Int32 sNoPrompt = 0x01;
        const int s = 0x12;
        const Int32 sUseFileName = 0x02;
        const Int32 sBroadCast = 0x20;
        const Int32 iConcat = 0x04;
        const Int32 sEmbedFonts = 0x10;
        const Int32 lExportToRtf = 0x8000000;
        const Int32 lExportToTIFF = 0x8000000;
        const Int32 JPEGExport = 0x10000000;
        const Int32 FullEmbed = 0x200;
        const Int32 MultilingualSupport = 0x80;
        const Int32 JPegLevelMedium = 0x00040000;
        const Int32 PrintWatermark = 0x40;
        const Int32 SendToCreator = 0x2000000;
        const Int32 EmbedStandardFonts = 0x00200000;
        const Int32 ConvertHyperlinks = 0x00100000;
        const Int32 AddIdNumber = 0x00004000;
        const Int32 AddDateTime = 0x00003000;

        //Evaluation Codes of the Amyuni PDF Suite
        const string strLicenseTo = "Amyuni Technologies Eval Version";
        const string strActivationCode = "07EFCDAB0100010031825101B257659266255C64F543DF28853EEACF7AFD9DC81D66B62926B87CEC1161728458C731CE6C2429A6B440";

        /// <summary>
        /// Print MsWord Docs
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PrintWord(object sender, System.EventArgs e)
        {


            //Declare object
            CDIntfEx.CDIntfEx PDF = new CDIntfEx.CDIntfExClass();

            //Initialize Printer
            PDF.DriverInit(PDFprinter);

            //set FileName for resulting PDF document
            PDF.DefaultFileName = getSampleWorkingDirectory() + "Resulting_Docs\\print_word.pdf";
            PDF.FileNameOptionsEx = (int)(sNoPrompt + sUseFileName);

            PDF.EnablePrinter(strLicenseTo, strActivationCode);

            //Print to Word

            PrintToMsWord(getSampleWorkingDirectory() + "Source_Docs\\batchconvert.doc");

            PDF.FileNameOptions = 0;
            PDF.DriverEnd();

        }

        public void PrintToMsWord(string strFileName)
        {
            //object oMissing = System.Reflection.Missing.Value;

            //Start Word and open the test document.

            //    Microsoft.Office.Interop.Word._Application oWord;
            //    Microsoft.Office.Interop.Word._Document oDoc;
            //    oWord = new Microsoft.Office.Interop.Word.Application();
            //    oWord.Visible = false;
            //    object oPath;

            //    string strSaveActivePrinter;

            //    //Save active printer
            //    strSaveActivePrinter = oWord.ActivePrinter;

            //    //assign word a printer
            //    oWord.ActivePrinter = PDFprinter;
            //    oPath = strFileName.ToString();

            //    oDoc = oWord.Documents.Open(ref oPath,
            //        ref oMissing, ref oMissing, ref oMissing,
            //        ref oMissing, ref oMissing, ref oMissing,
            //        ref oMissing, ref oMissing,
            //        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
            //        ref oMissing, ref oMissing, ref oMissing);

            //    //print it
            //    object xcopies = 1;
            //    object xpt = false;
            //    object oFalse = false;

            //    oDoc.PrintOut(ref xpt, ref oMissing, ref oMissing,
            //        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref 
            //        xcopies, ref oMissing, ref oMissing, ref oMissing, ref 
            //        oMissing, ref oMissing, ref oMissing, ref oMissing, ref 
            //        oMissing, ref oMissing, ref oMissing);



            //    //close the document
            //    oDoc.Close(ref oFalse, ref oMissing, ref oMissing);

            //    oWord.Options.SaveNormalPrompt = false;
            //    oWord.Options.SavePropertiesPrompt = false;
            //    oWord.NormalTemplate.Saved = true;

            //    //restore active printer
            //    oWord.ActivePrinter = strSaveActivePrinter;

            //    //close word
            //    oWord.Quit(ref oFalse, ref oMissing, ref oMissing);

            //    //Drop our reference to the COM object
            //    Marshal.ReleaseComObject(oWord);

            //    oDoc = null;
            //    oWord = null;
        }

        /// <summary>
        /// Print existing PDF document
        /// The Print method can be used to print a PDF document to a hardware printer. It is also used to print 
        /// multiple pages on a single sheet of paper.This method is available only in the professional 
        /// version of the Amyuni PDF Converter product.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PrintPDFDoc()
        {

            CDIntfEx.Document PDFDoc = new CDIntfEx.DocumentClass();

          
                PDFDoc.OpenEx(getSampleWorkingDirectory() + "Source_Docs\\fivepages.pdf", "");

                PDFDoc.SetLicenseKey(strLicenseTo, strActivationCode);

                //Parameters
                //PrinterName
                //[in] Name of printer as it shows in the printers control panel. 
                //If this parameter is left empty, the document will print to the default printer
                //StartPage
                //[in] Page number from which to start printing. The index of the first page is 1
                //EndPage
                //[in] Page number at which to stop printing
                //Copies
                //[in] Number of copies to print the document

                PDFDoc.Print("", 1, PDFDoc.PageCount(), 1);

           

        }


        public void PrintWaterMark()
        {
            CDIntfEx.Document oDoc = new CDIntfEx.DocumentClass();
            oDoc.SetLicenseKey(strLicenseTo, strActivationCode);
            oDoc.Open(getSampleWorkingDirectory() + "Source_Docs\\fivepages.pdf");
            int nPageCount = oDoc.PageCount();

            oDoc.Print("", 1, nPageCount, 1);
        }



        /// <summary>
        /// Encrypt document at run-time
        /// The Encryption property can be used to password protect a PDF document and restrict 
        /// users to viewing, modifying or evenprinting the document.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Encrypt()
        {

            CDIntfEx.CDIntfExClass PDF = new CDIntfEx.CDIntfExClass();
            PDF.DriverInit(PDFprinter);

            PDF.DefaultFileName = getSampleWorkingDirectory() + "Resulting_Docs\\encrypt.pdf";

            PDF.SetDefaultPrinter();


            /*////////////////////////////////////////////////////////////
            'Permission                         Permission value
            'Enable Printing                        - 64 + 4
            'Enable document modification           - 64 + 8
            'Enable copying text and graphics       - 64 + 16
            'Enable adding and changing notes       - 64 + 32
            'To combine multiple options, use -64 plus the values 4, 8, 16 or 32. E.g. to enable */

            //1 = 40-bit encryption
            //2 = 128-bit encryption

            PDF.Encryption = 2;
            PDF.OwnerPassword = "aaaaaa";
            PDF.UserPassword = "bbbbbb";
            PDF.Permissions = (-64 + 4);
            PDF.SetDefaultConfig();


            PDF.FileNameOptionsEx = (int)sNoPrompt + sUseFileName;
            PDF.EnablePrinter(strLicenseTo, strActivationCode);

            //Print something
           // printDocument1.Print();

            //Reset
            PDF.Encryption = 0;
            PDF.SetDefaultConfig();

            PDF.RestoreDefaultPrinter();
            PDF.FileNameOptions = 0;

        }

        /// <summary>
        /// The BatchConvert method converts a number of files to PDF, RTF, HTML, Excel or JPeg formats in batch mode.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void BatchConvert()
        {
            CDIntfEx.CDIntfExClass PDF = new CDIntfEx.CDIntfExClass();
            PDF.DriverInit(PDFprinter);  //Amyuni PDF Converter

            PDF.SetDefaultPrinter();
            PDF.DefaultDirectory = getSampleWorkingDirectory() + "Resulting_Docs\\";
            PDF.DefaultFileName = getSampleWorkingDirectory() + "Resulting_Docs\\test.pdf";

            PDF.FileNameOptionsEx = (int)sNoPrompt + sUseFileName;

            PDF.EnablePrinter(strLicenseTo, strActivationCode);

            /*The Document Converter printer should be configured with the destination file name and all other 
             * options before calling this function. The printer should also be set as default printer. 
             * This function launches the application that is associated with a specific file and issues a print 
             * command to convert the document.*/

            PDF.BatchConvert(getSampleWorkingDirectory() + "Source_Docs\\Sample.docx");
            PDF.RestoreDefaultPrinter();
            PDF.FileNameOptions = 0;

        }

        /// <summary>
        /// This method is only used to get the working directory of this sample project.
        /// </summary>
        /// <returns>Directory path to sample project</returns>
        string getSampleWorkingDirectory()
        {
            string currentDirName = System.IO.Directory.GetCurrentDirectory();
            string[] directories = currentDirName.Split(Path.DirectorySeparatorChar);


            StringBuilder strWorkingPath = new StringBuilder();

            if (null != directories)
            {
                foreach (string direc in directories)
                {
                    strWorkingPath.Append(direc + "\\");
                    if (direc == "PDFConverter")
                    {
                        return strWorkingPath.ToString();
                    }
                }
            }
            //Can't find project
            return "";

        }

        public void test()
        {
            StandardPrintController std = new StandardPrintController();
        }
       
    }
}
