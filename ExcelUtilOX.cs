using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
namespace CommonLib
{
    public static class ExcelUtilOX
    {

        private static readonly ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static IEnumerable<string> GetAllSheetNamesViaOpenXml(string excelFullPathFile)
        {
            IEnumerable<string> res = null;
            try
            {
                using (var doc = SpreadsheetDocument.Open(excelFullPathFile, false))
                {
                    //var sheetsa = doc.WorkbookPart.Workbook.Sheets.Select(a => a.GetAttribute("name", "").Value).ToList();  //.Select(a=>a.LocalName.).ToList()

                    ////foreach (var sheet in sheets)
                    ////{
                    ////    foreach (var attr in sheet.GetAttributes())
                    ////    {
                    ////        Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                    ////    }
                    ////}

                    ////doc.WorkbookPart.Workbook.Descendants<Sheet>()

                    //var sheets = doc.WorkbookPart.Workbook.Sheets.Cast<Sheet>().ToList();   //.Select(a=>a.Name)
                    //sheets.ForEach(x => log.Debug(
                    //      String.Format("RelationshipId:{0}\n SheetName:{1}\n SheetId:{2}"
                    //      , x.Id.Value, x.Name.Value, x.SheetId.Value)));

                    res = doc.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Select(a => a.Name.Value).ToArray();
                }
            }
            catch (System.IO.FileFormatException ex) {
                log.Error(ex.ToString());
                throw ex;
            }
            catch (System.IO.FileNotFoundException ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }
            catch (System.IO.IOException ex) {
                log.Error(ex.ToString());
            }
            catch (Exception ex)
            {
                log.Error(ex.ToString());
                throw ex;
            }

            return res;
        }


        // The DOM approach.
        // Note that the code below works only for cells that contain numeric values.
        // 
        static void ReadExcelFileDOM(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.Text;
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }

        // The SAX approach.
        static void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }
    }
}
