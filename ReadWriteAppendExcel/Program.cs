using System;
using IronXL;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ReadWriteAppendExcel
{
    class Program
    {
        private const string Filename = "TrueID_Responses_Data.xlsx";
        private const string WorkSheetName = "TrueIDResponses";
        private const string Author = "IClearAccountOpeningAPI";
        private const string Title = "TrueID_Responses_Data";

        static void Main(string[] args)
        {
            //Using IronXL nuget package
            //try
            //{
            //    WorkBook workbook = WorkBook.Load(Filename);
            //    WorkSheet sheet = workbook.GetWorkSheet(WorkSheetName);
            //    AppendDataInFile(sheet, "321654321654", "321654321654", "passed", "{abc: asd}");
            //    workbook.Save();
            //}
            //catch (Exception ex)
            //{
            //    if (ex is System.IO.FileNotFoundException)
            //    {
            //        var newWorkBook = WorkBook.Create(ExcelFileFormat.XLSX);
            //        newWorkBook.Metadata.Author = Author;
            //        newWorkBook.Metadata.Title = Title;
            //        newWorkBook.Metadata.Created = DateTime.UtcNow;
            //        newWorkBook.Metadata.Modified = DateTime.UtcNow;

            //        WorkSheet sheet = newWorkBook.CreateWorkSheet(WorkSheetName);
            //        CreateHeaderForNewFile(sheet);
            //        AppendDataInFile(sheet, "321654321654", "321654321654", "passed", "{abc: asd}");
            //        newWorkBook.SaveAs(Filename);
            //    }

            //    else
            //    {
            //        Console.WriteLine("There was an error in reading the file.");
            //    }
            //}


            //Using ClosedXML nuget package
            //using (var workbook = new XLWorkbook())
            //{
            //    var worksheet = workbook.Worksheets.Add("Sample Sheet");
            //    worksheet.Cell("A1").Value = "Hello World!";
            //    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
            //    workbook.SaveAs("HelloWorld.xlsx");
            //}
            try
            {
                XLWorkbook workbook = new XLWorkbook(Filename);
                bool isSheetAvailable = workbook.Worksheets.TryGetWorksheet(WorkSheetName, out IXLWorksheet sheet);
                if (isSheetAvailable)
                    AppendDataInFileUsingClosedXML(sheet, "321654321654", "321654321654", "passed", "{abc: asd}");
                workbook.Save();
            }
            catch (Exception ex)
            {
                if (ex is System.IO.FileNotFoundException)
                {
                    var newWorkBook = new XLWorkbook();
                    newWorkBook.Author = Author;

                    IXLWorksheet sheet = newWorkBook.AddWorksheet(WorkSheetName);
                    CreateHeaderForNewFileUsingClosedXML(sheet);
                    AppendDataInFileUsingClosedXML(sheet, "321654321654", "321654321654", "passed", "{abc: asd}");
                    newWorkBook.SaveAs(Filename);
                }

                else
                {
                    Console.WriteLine("There was an error in reading the file.");
                }
            }
        }

        private static void AppendDataInFileUsingClosedXML(IXLWorksheet sheet, string reference, string conversationId, string transactionStatus, string stringifiedResponse)
        {
            var dataWritePointerCount = sheet.RangeUsed().RowCount() + 1;
            sheet.Cell("A" + dataWritePointerCount).Value = dataWritePointerCount - 1;
            sheet.Cell("B" + dataWritePointerCount).Value = DateTime.UtcNow;
            
            sheet.Cell("C" + dataWritePointerCount).Value = reference;
            sheet.Cell("C" + dataWritePointerCount).DataType = XLDataType.Text;

            sheet.Cell("D" + dataWritePointerCount).Value = conversationId;
            sheet.Cell("D" + dataWritePointerCount).DataType = XLDataType.Text;

            sheet.Cell("E" + dataWritePointerCount).Value = transactionStatus;
            sheet.Cell("E" + dataWritePointerCount).DataType = XLDataType.Text;

            sheet.Cell("F" + dataWritePointerCount).Value = stringifiedResponse;
        }
        private static void CreateHeaderForNewFileUsingClosedXML(IXLWorksheet sheet)
        {
            var cells = new Dictionary<string, string>() {
                { "A1", "S.No"},
                { "B1", "Recording Date"},
                { "C1", "Reference Number"},
                { "D1", "Conversation Id"},
                { "E1", "Transaction Status"},
                { "F1", "Full Stringified Response"}
            };

            foreach (var cell in cells.Keys)
            {
                sheet.Cell(cell).Value = cells[cell];
                sheet.Cell(cell).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                sheet.Cell(cell).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                sheet.Cell(cell).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                sheet.Cell(cell).Style.Border.RightBorder = XLBorderStyleValues.Thick;
            }
        }


        private static void AppendDataInFile(WorkSheet sheet, string reference, string conversationId, string transactionStatus, string stringifiedResponse)
        {
            var dataWritePointerCount = sheet.RowCount + 1;
            sheet["A" + dataWritePointerCount].Value = dataWritePointerCount - 1;
            sheet["B" + dataWritePointerCount].Value = DateTime.UtcNow;
            sheet["C" + dataWritePointerCount].Value = reference;
            sheet["D" + dataWritePointerCount].Value = conversationId;
            sheet["E" + dataWritePointerCount].Value = transactionStatus;
            sheet["F" + dataWritePointerCount].Value = stringifiedResponse;
        }
        private static void CreateHeaderForNewFile(WorkSheet sheet)
        {
            var cells = new Dictionary<string, string>() {
                { "A1", "S.No"},
                { "B1", "Recording Date"},
                { "C1", "Reference Number"},
                { "D1", "Conversation Id"},
                { "E1", "Transaction Status"},
                { "F1", "Full Stringified Response"}
            };

            foreach (var cell in cells.Keys)
            {
                sheet[cell].Value = cells[cell];
                sheet[cell].Style.TopBorder.Type = IronXL.Styles.BorderType.Thick;
                sheet[cell].Style.BottomBorder.Type = IronXL.Styles.BorderType.Thick;
                sheet[cell].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thick;
                sheet[cell].Style.RightBorder.Type = IronXL.Styles.BorderType.Thick;
                sheet[cell].Style.SetBackgroundColor("999999");
                sheet[cell].Style.ShrinkToFit = false;
            }

            sheet.Columns[0].Width = 3000;
            sheet.Columns[1].Width = 6000;
            sheet.Columns[2].Width = 6000;
            sheet.Columns[3].Width = 6000;
            sheet.Columns[4].Width = 6000;
            sheet.Columns[5].Width = 40000;
        }
    }
}
