using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Geom;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using BIMQuantityManager.Model;
using System.IO;

namespace BIMQuantityManager.Helpers;

public sealed class BimQuantityManagerPdfHelper
{
    public static void ExportToPdf(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        try
        {
            var props = new WriterProperties();
            typeof(WriterProperties)
                .GetField("smartMode", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.SetValue(props, false);

            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var writer = new PdfWriter(stream, props))
            using (var pdf = new PdfDocument(writer))
            {
                PageSize pageSize = GetOptimalPageSize(selectedViews.Count);
                pdf.SetDefaultPageSize(pageSize);

                using (var document = new Document(pdf))
                {
                    PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.TIMES_BOLD);
                    PdfFont normalFont = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);

                    var title = new Paragraph("BIM Quantity Manager")
                        .SetFont(boldFont)
                        .SetFontSize(18)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetMarginBottom(10);
                    document.Add(title);

                    var dateText = $"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                    var date = new Paragraph(dateText)
                        .SetFont(normalFont)
                        .SetFontSize(13)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetMarginBottom(10);
                    document.Add(date);

                    var columnWidths = GetDynamicColumnWidths(selectedViews.Count, pageSize);

                    var table = new Table(UnitValue.CreatePointArray(columnWidths.ToArray()))
                        .UseAllAvailableWidth()
                        .SetMarginTop(5);

                    AddHeaderCell(table, "Category", boldFont);
                    AddHeaderCell(table, "Type Name", boldFont);
                    AddHeaderCell(table, "Total Quantity", boldFont);
                    AddHeaderCell(table, activeViewName, boldFont);

                    foreach (var view in selectedViews)
                    {
                        AddHeaderCell(table, view.DisplayName, boldFont);
                    }

                    bool isAlternate = false;
                    foreach (var type in typeData)
                    {
                        AddDataCell(table, type.Category ?? string.Empty, boldFont, normalFont, true, TextAlignment.LEFT, isAlternate);
                        AddDataCell(table, type.TypeName ?? string.Empty, boldFont, normalFont, false, TextAlignment.LEFT, isAlternate);
                        AddDataCell(table, type.TotalQuantity.ToString(), boldFont, normalFont, false, TextAlignment.CENTER, isAlternate);
                        AddDataCell(table, type.ActiveViewQuantity.ToString(), boldFont, normalFont, false, TextAlignment.CENTER, isAlternate);

                        foreach (var view in selectedViews)
                        {
                            long viewKey = view.ViewId.Value;
                            int quantity = type.ViewQuantities.ContainsKey(viewKey)
                                ? type.ViewQuantities[viewKey]
                                : 0;
                            AddDataCell(table, quantity.ToString(), boldFont, normalFont, false, TextAlignment.CENTER, isAlternate);
                        }

                        isAlternate = !isAlternate;
                    }

                    document.Add(table);

                    int totalCount = typeData.Count();
                    var footerText = $"Total Types: {totalCount} | Page Size: {GetPageSizeName(pageSize)}";
                    var footer = new Paragraph(footerText)
                        .SetFont(boldFont)
                        .SetFontSize(13)
                        .SetTextAlignment(TextAlignment.CENTER)
                        .SetMarginTop(10)
                        .SetFontColor(new DeviceRgb(44, 62, 80));
                    document.Add(footer);

                    document.Close();
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"PDF Export Error: {ex.Message}\nInner: {ex.InnerException?.Message ?? "None"}\nStack: {ex.StackTrace}", ex);
        }
    }

    private static PageSize GetOptimalPageSize(int selectedViewCount)
    {
        int totalColumns = 4 + selectedViewCount; 

        if (totalColumns <= 6)
            return PageSize.A4.Rotate(); 
        else if (totalColumns <= 8)
            return PageSize.A3.Rotate(); 
        else if (totalColumns <= 10)
            return PageSize.A2.Rotate(); 
        else if (totalColumns <= 12)
            return PageSize.A1.Rotate(); 
        else
            return PageSize.A0.Rotate(); 
    }


    private static List<float> GetDynamicColumnWidths(int selectedViewCount, PageSize pageSize)
    {
        float availableWidth = pageSize.GetWidth() - 80; 
        int totalColumns = 4 + selectedViewCount;

        var columnWidths = new List<float>();

        float categoryWidth = availableWidth * 0.15f;
        float typeNameWidth = availableWidth * 0.30f;
        float totalQtyWidth = availableWidth * 0.075f;
        float activeViewWidth = availableWidth * 0.075f;

        columnWidths.Add(categoryWidth);
        columnWidths.Add(typeNameWidth);
        columnWidths.Add(totalQtyWidth);
        columnWidths.Add(activeViewWidth);

        float remainingWidth = availableWidth * 0.40f;
        float dynamicColumnWidth = selectedViewCount > 0 ? remainingWidth / selectedViewCount : 0;

        for (int i = 0; i < selectedViewCount; i++)
        {
            columnWidths.Add(dynamicColumnWidth);
        }

        return columnWidths;
    }

    private static string GetPageSizeName(PageSize pageSize)
    {
        if (pageSize.Equals(PageSize.A4.Rotate())) return "A4 Landscape";
        if (pageSize.Equals(PageSize.A3.Rotate())) return "A3 Landscape";
        if (pageSize.Equals(PageSize.A2.Rotate())) return "A2 Landscape";
        if (pageSize.Equals(PageSize.A1.Rotate())) return "A1 Landscape";
        if (pageSize.Equals(PageSize.A0.Rotate())) return "A0 Landscape";
        return "Custom";
    }

    private static void AddHeaderCell(Table table, string text, PdfFont boldFont)
    {
        var para = new Paragraph(text)
            .SetFont(boldFont)
            .SetFontColor(ColorConstants.WHITE)
            .SetFontSize(11);

        var cell = new Cell()
            .Add(para)
            .SetBackgroundColor(new DeviceRgb(44, 62, 80))
            .SetTextAlignment(TextAlignment.CENTER)
            .SetPadding(5);

        table.AddHeaderCell(cell);
    }

    private static void AddDataCell(Table table, string text, PdfFont boldFont, PdfFont normalFont,
        bool isBold, TextAlignment alignment, bool isAlternate)
    {
        string cellText = text ?? string.Empty;
        var paragraph = new Paragraph(cellText)
            .SetFontSize(11);

        paragraph.SetFont(isBold ? boldFont : normalFont);

        var backgroundColor = isAlternate
            ? new DeviceRgb(249, 249, 249)
            : new DeviceRgb(255, 255, 255);

        var cell = new Cell()
            .Add(paragraph)
            .SetBackgroundColor(backgroundColor)
            .SetTextAlignment(alignment)
            .SetPadding(5);

        table.AddCell(cell);
    }
}