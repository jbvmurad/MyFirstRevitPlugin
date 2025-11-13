using ClosedXML.Excel;
using BIMQuantityManager.Model;
#pragma warning disable CS0618

namespace BIMQuantityManager.Helpers;

public sealed class BimQuantityManagerExcelHelper
{
    public static void ExportToExcel(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        using (var workbook = new XLWorkbook())
        {
            var ws = workbook.Worksheets.Add("BIM Quantity Manager");

            int totalCols = 4 + selectedViews.Count;

            ws.Range(1, 1, 1, totalCols).Merge();
            ws.Cell(1, 1).Value = "BIM QUANTITY MANAGER";
            ws.Cell(1, 1).Style.Font.Bold = true;
            ws.Cell(1, 1).Style.Font.FontSize = 20;
            ws.Cell(1, 1).Style.Font.FontName = "Aptos Narrow";
            ws.Cell(1, 1).Style.Font.FontColor = XLColor.White;
            ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(41, 128, 185);
            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Cell(1, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            ws.Row(1).Height = 35;

            ws.Range(2, 1, 2, totalCols).Merge();
            ws.Cell(2, 1).Value = "Normal Table - Complete Data Export";
            ws.Cell(2, 1).Style.Font.FontSize = 12;
            ws.Cell(2, 1).Style.Font.FontName = "Aptos Narrow";
            ws.Cell(2, 1).Style.Font.Italic = true;
            ws.Cell(2, 1).Style.Font.FontColor = XLColor.FromArgb(127, 140, 141);
            ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(2).Height = 20;

            ws.Range(3, 1, 3, totalCols).Merge();
            ws.Cell(3, 1).Value = $"Generated: {DateTime.Now:dd MMMM yyyy - HH:mm:ss}";
            ws.Cell(3, 1).Style.Font.FontSize = 10;
            ws.Cell(3, 1).Style.Font.FontName = "Aptos Narrow";
            ws.Cell(3, 1).Style.Font.FontColor = XLColor.FromArgb(149, 165, 166);
            ws.Cell(3, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(3).Height = 18;

            int col = 1;
            ws.Cell(5, col).Value = "Category";
            ws.Cell(5, col).Style.Font.FontName = "Aptos Narrow";
            col++;

            ws.Cell(5, col).Value = "Type Name";
            ws.Cell(5, col).Style.Font.FontName = "Aptos Narrow";
            col++;

            ws.Cell(5, col).Value = "Total Quantity";
            ws.Cell(5, col).Style.Font.FontName = "Aptos Narrow";
            col++;

            ws.Cell(5, col).Value = activeViewName;
            ws.Cell(5, col).Style.Font.FontName = "Aptos Narrow";
            col++;

            int dynamicStartCol = col;
            foreach (var view in selectedViews)
            {
                ws.Cell(5, col).Value = view.DisplayName;
                ws.Cell(5, col).Style.Font.FontName = "Aptos Narrow";
                col++;
            }

            var headerRange = ws.Range(5, 1, 5, totalCols);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontSize = 12;
            headerRange.Style.Font.FontName = "Aptos Narrow";
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(52, 73, 94);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            headerRange.Style.Border.OutsideBorderColor = XLColor.FromArgb(44, 62, 80);
            ws.Row(5).Height = 25;

            int row = 6;
            int dataCount = 0;
            foreach (var type in typeData)
            {
                col = 1;

                if (!string.IsNullOrEmpty(type.Category))
                {
                    ws.Cell(row, col).Value = type.Category;
                    ws.Cell(row, col).Style.Font.Bold = true;
                    ws.Cell(row, col).Style.Font.FontName = "Aptos Narrow";
                    ws.Cell(row, col).Style.Font.FontColor = XLColor.FromArgb(44, 62, 80);
                }
                col++;

                ws.Cell(row, col).Value = type.TypeName;
                ws.Cell(row, col).Style.Font.FontName = "Aptos Narrow";
                ws.Cell(row, col).Style.Font.FontColor = XLColor.FromArgb(52, 73, 94);
                col++;

                ws.Cell(row, col).Value = type.TotalQuantity;
                ws.Cell(row, col).Style.Font.Bold = true;
                ws.Cell(row, col).Style.Font.FontName = "Aptos Narrow";
                ws.Cell(row, col).Style.Font.FontColor = XLColor.FromArgb(41, 128, 185);
                col++;

                ws.Cell(row, col).Value = type.ActiveViewQuantity;
                ws.Cell(row, col).Style.Font.Bold = true;
                ws.Cell(row, col).Style.Font.FontName = "Aptos Narrow";
                ws.Cell(row, col).Style.Font.FontColor = XLColor.FromArgb(41, 128, 185);
                col++;

                foreach (var view in selectedViews)
                {
                    long viewKey = view.ViewId.Value;
                    int quantity = type.ViewQuantities.ContainsKey(viewKey)
                        ? type.ViewQuantities[viewKey]
                        : 0;
                    ws.Cell(row, col).Value = quantity;
                    ws.Cell(row, col).Style.Font.Bold = true;
                    ws.Cell(row, col).Style.Font.FontName = "Aptos Narrow";
                    ws.Cell(row, col).Style.Font.FontColor = XLColor.FromArgb(41, 128, 185);
                    col++;
                }

                for (int c = 3; c <= totalCols; c++)
                {
                    ws.Cell(row, c).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                var rowRange = ws.Range(row, 1, row, totalCols);

                if (dataCount % 2 == 0)
                {
                    rowRange.Style.Fill.BackgroundColor = XLColor.FromArgb(236, 240, 241);
                }
                else
                {
                    rowRange.Style.Fill.BackgroundColor = XLColor.White;
                }

                rowRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rowRange.Style.Border.OutsideBorderColor = XLColor.FromArgb(189, 195, 199);

                row++;
                dataCount++;
            }
 
            ws.Column(1).Width = 25; 
            ws.Column(2).Width = 55; 
            ws.Column(3).Width = 18; 
            ws.Column(4).Width = 28; 

            for (int c = dynamicStartCol; c <= totalCols; c++)
            {
                ws.Column(c).Width = 28; 
            }

            workbook.SaveAs(path);
        }
    }
}