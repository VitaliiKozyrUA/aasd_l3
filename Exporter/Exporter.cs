using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word.Document;

namespace Exporter
{
    public class Exporter
    {
        public static void exportToExcel(List<List<object>> ordersData, List<List<object>> productsData)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            var workBook = excelApp.Workbooks.Add();
            _Worksheet workSheetOrders = workBook.Worksheets.Add();
            _Worksheet workSheetProducts = workBook.Worksheets.Add();

            exportDataToExcel(workSheetOrders, "Orders", ordersData);
            exportDataToExcel(workSheetProducts, "Products", productsData);

            Marshal.ReleaseComObject(workSheetOrders);
            Marshal.ReleaseComObject(workSheetProducts);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void exportDataToExcel(_Worksheet workSheet, string sheetName, List<List<object>> data)
        {
            workSheet.Name = sheetName;

            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    workSheet.Cells[i + 1, j + 1] = data[i][j];
                }
            }
        }

        public static void exportToWordDoc(List<List<object>> ordersData, List<List<object>> productsData)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            var document = wordApp.Documents.Add();

            var ordersSection = document.Content.Paragraphs.Add();
            ordersSection.Range.Text = "Orders";
            ordersSection.Range.InsertParagraphAfter();
            exportDataToWord(document, ordersData);

            var productsSection = document.Content.Paragraphs.Add();
            productsSection.Range.Text = "Products";
            productsSection.Range.InsertParagraphAfter();
            exportDataToWord(document, productsData);

            Marshal.ReleaseComObject(ordersSection);
            Marshal.ReleaseComObject(productsSection);
            Marshal.ReleaseComObject(document);
            Marshal.ReleaseComObject(wordApp);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void exportDataToWord(Document document, List<List<object>> data)
        {
            var dataTable = document.Tables.Add(document.Content.Paragraphs.Last.Range, data.Count, data[0].Count);

            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Count; j++)
                {
                    dataTable.Cell(i + 1, j + 1).Range.Text = data[i][j].ToString();
                }
            }
        }

    }
}
