using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BackEnd
{
    public class XLSWatermark
    {
        public static void WatermarkAllSheets(SpreadsheetDocument spreadsheet)
        {
            WorkbookPart workbookPart = spreadsheet.WorkbookPart;
            foreach (Sheet item in workbookPart.Workbook.Sheets)
            {
                WorksheetPart worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(item.Id);

                HeaderFooter headerFooter = new HeaderFooter();
                OddHeader oddHeader1 = new OddHeader();
                oddHeader1.Text = "&C&35&K03+079Microsoft Corporation\nOne Microsoft Way\nRedmond, WA 98052\ndcxhelp@micrsoft.com";
                headerFooter.Append(oddHeader1);

                worksheetPart.Worksheet.Append(headerFooter);
            }

        }
    }
}
