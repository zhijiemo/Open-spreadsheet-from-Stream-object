using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using System.IO;

namespace Open_spreadsheet_from_Stream_object
{
    class Program
    {
        static void Main(string[] args)
        {
            FileStream fs = new FileStream("OpenFromStreamOriginal.xlsx", FileMode.Open);

            MemoryStream msFirstPass = new MemoryStream();
            SLDocument slFirstPass = new SLDocument(fs, "Sheet1");
            slFirstPass.SetCellValue(5, 1, "Got the baton!");
            slFirstPass.SaveAs(msFirstPass);

            MemoryStream msSecondPass = new MemoryStream();
            SLDocument slSecondPass = new SLDocument(msFirstPass, "Sheet2");
            slSecondPass.SetCellValue(5, 1, "Passed it to the second guy!");
            slSecondPass.SaveAs(msSecondPass);

            MemoryStream msThirdPass = new MemoryStream();
            SLDocument slThirdPass = new SLDocument(msSecondPass, "Sheet3");
            slThirdPass.SetCellValue(5, 1, "Handed it over to the third guy!");
            slThirdPass.SaveAs(msThirdPass);

            SLDocument slFinalPass = new SLDocument(msThirdPass, "Sheet3");
            slFinalPass.AddWorksheet("Sheet4");
            slFinalPass.SetCellValue(5, 1, "And we cross the finish line!");

            slFinalPass.SaveAs("OpenFromStreamModified.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
