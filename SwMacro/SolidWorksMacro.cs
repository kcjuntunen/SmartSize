using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System;

namespace SmartSize.csproj
{
    public partial class SolidWorksMacro
    {

        public void Main()
        {
            SmartPrinter x = new SmartPrinter(swApp);
            x.SetSizes();
            x.PrintAll();
        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        SldWorks swApp;
    }

    public class SmartPrinter
    {
        private SldWorks sw;
        private ModelDoc2 swDoc;
        private ModelDocExtension swExt;
        private Sheet swSheet;
        private DrawingDoc swDrawing;
        private PageSetup myPageSetup;
        private double[] vsheetprops;

        public SmartPrinter(SldWorks sw)
        {        
            swDoc = (ModelDoc2)sw.ActiveDoc;
            swDrawing = (DrawingDoc)sw.ActiveDoc;
            swSheet = (Sheet)swDrawing.GetCurrentSheet();
            swExt = swDoc.Extension;
        }

        public void SetSizes()
        {
            if (!(swExt.UsePageSetup == (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet))
            {
                swExt.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_Document;
                swExt.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;
            }

            foreach (string x in (string[])swDrawing.GetSheetNames())
            {
                swSheet = swDrawing.get_Sheet(x);
                myPageSetup = (PageSetup)swSheet.PageSetup;
                vsheetprops = (double[])swSheet.GetProperties();

                double len = (double)vsheetprops[5] / 0.0254;
                double wid = (double)vsheetprops[6] / 0.0254;

                System.Diagnostics.Debug.Print("{0}: {1} x {2}\nCurrent Paper Size: {3}, Orientation: {4}",
                    x,
                    len,
                    wid,
                    Enum.GetName(typeof(swDwgPaperSizes_e), myPageSetup.PrinterPaperSize),
                    Enum.GetName(typeof(swPageSetupOrientation_e), myPageSetup.Orientation));

                if (len == 17 && wid == 11) // 11 x 17
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                    myPageSetup.PrinterPaperSize = 17;
                }
                else if (len == 11 && wid == 17)
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                    myPageSetup.PrinterPaperSize = 17;
                }
                else if (len == 11 && wid == 8.5) // 8.5 x 11
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                    myPageSetup.PrinterPaperSize = 1;
                }
                else if (len == 8.5 && wid == 11)
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                    myPageSetup.PrinterPaperSize = 1;
                }
                //else if (len == 36 && wid == 24) // Arch D
                //{
                //    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                //    myPageSetup.PrinterPaperSize = 214;
                //}
                //else if (len == 24 && wid == 36)
                //{
                //    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                //    myPageSetup.PrinterPaperSize = 214;
                //}

                System.Diagnostics.Debug.Print("Set to: {0}/{1}, Orientation: {2}",
                    myPageSetup.PrinterPaperSize,
                    Enum.GetName(typeof(swDwgPaperSizes_e), myPageSetup.PrinterPaperSize),
                    Enum.GetName(typeof(swPageSetupOrientation_e), myPageSetup.Orientation));
            }   

        }

        public void SetSizesLetter()
        {
            if (!(swExt.UsePageSetup == (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet))
            {
                swExt.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_Document;
                swExt.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;
            }

            foreach (string x in (string[])swDrawing.GetSheetNames())
            {
                swSheet = swDrawing.get_Sheet(x);
                myPageSetup = (PageSetup)swSheet.PageSetup;
                vsheetprops = (double[])swSheet.GetProperties();

                double len = (double)vsheetprops[5] / 0.0254;
                double wid = (double)vsheetprops[6] / 0.0254;

                System.Diagnostics.Debug.Print("{0}: {1} x {2}\nCurrent Paper Size: {3}, Orientation: {4}",
                    x,
                    len,
                    wid,
                    Enum.GetName(typeof(swDwgPaperSizes_e), myPageSetup.PrinterPaperSize),
                    Enum.GetName(typeof(swPageSetupOrientation_e), myPageSetup.Orientation));

                myPageSetup.PrinterPaperSize = 1;

                if (len == 17 && wid == 11) // 11 x 17
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                }
                else if (len == 11 && wid == 17)
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                }
                else if (len == 11 && wid == 8.5) // 8.5 x 11
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                }
                else if (len == 8.5 && wid == 11)
                {
                    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                }
                //else if (len == 36 && wid == 24) // Arch D
                //{
                //    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
                //    myPageSetup.PrinterPaperSize = 214;
                //}
                //else if (len == 24 && wid == 36)
                //{
                //    myPageSetup.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;
                //    myPageSetup.PrinterPaperSize = 214;
                //}

                System.Diagnostics.Debug.Print("Set to: {0}/{1}, Orientation: {2}",
                    myPageSetup.PrinterPaperSize,
                    Enum.GetName(typeof(swDwgPaperSizes_e), myPageSetup.PrinterPaperSize),
                    Enum.GetName(typeof(swPageSetupOrientation_e), myPageSetup.Orientation));
            }

        }

        public void PrintAll()
        {
            PrinterSettings p = new PrinterSettings();
            System.Diagnostics.Debug.Print("Printing with {0}", p.PrinterName);

            int[] sheets = { 0 };
            swExt.PrintOut3(sheets,                                 // PageArray
                1,                                                  // Copies
                false,                                              // Collate
                p.PrinterName,                                      // Printer
                // "\\\\AMSTORE-SVR-01\\HP Laserjet 5100 Engineering", // Printer
                "",                                                 // PrintFileName
                false);                                             // ConvertToHighQuality
        }
    }
}


