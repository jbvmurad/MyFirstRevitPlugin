using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using BIMQuantityManager.Model;
using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WordTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace BIMQuantityManager.Helpers;

public sealed class BimQuantityManagerWordHelper
{
    public static void ExportToWord(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            SetPageSize(body, selectedViews.Count);

            var titlePara = body.AppendChild(new WordParagraph());
            var titleRun = titlePara.AppendChild(new Run());
            titleRun.AppendChild(new Text("BIM Quantity Manager"));
            ApplyTitleStyle(titlePara);

            var datePara = body.AppendChild(new WordParagraph());
            var dateRun = datePara.AppendChild(new Run());
            dateRun.AppendChild(new Text($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}"));
            ApplyDateStyle(datePara);

            body.AppendChild(new WordParagraph());

            var columnWidths = GetDynamicColumnWidths(selectedViews.Count);

            var table = new WordTable();

            var tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                ),
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableJustification { Val = TableRowAlignmentValues.Center }
            );
            table.AppendChild(tblProp);

            var headerRow = new TableRow();
            headerRow.Append(
                CreateHeaderCell("Category", columnWidths[0]),
                CreateHeaderCell("Type Name", columnWidths[1]),
                CreateHeaderCell("Total Quantity", columnWidths[2]),
                CreateHeaderCell(activeViewName, columnWidths[3])
            );

            for (int i = 0; i < selectedViews.Count; i++)
            {
                headerRow.Append(CreateHeaderCell(selectedViews[i].DisplayName, columnWidths[4 + i]));
            }

            table.Append(headerRow);

            bool isAlternate = false;
            foreach (var type in typeData)
            {
                var dataRow = new TableRow();

                dataRow.Append(
                    CreateDataCell(type.Category ?? string.Empty, true, isAlternate, JustificationValues.Left, columnWidths[0]),
                    CreateDataCell(type.TypeName ?? string.Empty, false, isAlternate, JustificationValues.Left, columnWidths[1]),
                    CreateDataCell(type.TotalQuantity.ToString(), false, isAlternate, JustificationValues.Center, columnWidths[2]),
                    CreateDataCell(type.ActiveViewQuantity.ToString(), false, isAlternate, JustificationValues.Center, columnWidths[3])
                );

                for (int i = 0; i < selectedViews.Count; i++)
                {
                    long viewKey = selectedViews[i].ViewId.Value;
                    int quantity = type.ViewQuantities.ContainsKey(viewKey)
                        ? type.ViewQuantities[viewKey]
                        : 0;
                    dataRow.Append(CreateDataCell(quantity.ToString(), false, isAlternate, JustificationValues.Center, columnWidths[4 + i]));
                }

                table.Append(dataRow);
                isAlternate = !isAlternate;
            }

            body.Append(table);

            body.AppendChild(new WordParagraph());
            var footerPara = body.AppendChild(new WordParagraph());
            var footerRun = footerPara.AppendChild(new Run());
            string pageSizeName = GetPageSizeName(selectedViews.Count);
            footerRun.AppendChild(new Text($"Total Types: {typeData.Count()} | Page Size: {pageSizeName}"));
            ApplyFooterStyle(footerPara);

            mainPart.Document.Save();
        }
    }

    private static void SetPageSize(Body body, int selectedViewCount)
    {
        int totalColumns = 4 + selectedViewCount;

        uint width, height;

        if (totalColumns <= 6) 
        {
            width = 16838; 
            height = 11906;
        }
        else if (totalColumns <= 8) 
        {
            width = 23811; 
            height = 16838;
        }
        else if (totalColumns <= 10) 
        {
            width = 33676; 
            height = 23811;
        }
        else if (totalColumns <= 12) 
        {
            width = 47622; 
            height = 33676;
        }
        else 
        {
            width = 67408; 
            height = 47622;
        }

        var sectionProps = new SectionProperties(
            new PageSize
            {
                Width = width,
                Height = height,
                Orient = PageOrientationValues.Landscape
            },
            new PageMargin
            {
                Top = 720,    
                Right = 720,
                Bottom = 720,
                Left = 720,
                Header = 720,
                Footer = 720,
                Gutter = 0
            }
        );

        body.Append(sectionProps);
    }

    private static string[] GetDynamicColumnWidths(int selectedViewCount)
    {
        int totalColumns = 4 + selectedViewCount;
        var widths = new List<string>();

        int totalWidth = 5000;

        widths.Add(Math.Round(totalWidth * 0.15).ToString()); 
        widths.Add(Math.Round(totalWidth * 0.30).ToString()); 
        widths.Add(Math.Round(totalWidth * 0.075).ToString());
        widths.Add(Math.Round(totalWidth * 0.075).ToString());

        double dynamicWidth = selectedViewCount > 0 ? (totalWidth * 0.40) / selectedViewCount : 0;

        for (int i = 0; i < selectedViewCount; i++)
        {
            widths.Add(Math.Round(dynamicWidth).ToString());
        }

        return widths.ToArray();
    }

    private static string GetPageSizeName(int selectedViewCount)
    {
        int totalColumns = 4 + selectedViewCount;

        if (totalColumns <= 6) return "A4 Landscape";
        if (totalColumns <= 8) return "A3 Landscape";
        if (totalColumns <= 10) return "A2 Landscape";
        if (totalColumns <= 12) return "A1 Landscape";
        return "A0 Landscape";
    }

    private static void ApplyTitleStyle(WordParagraph paragraph)
    {
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { After = "200" }
        );
        paragraph.PrependChild(pPr);

        var run = paragraph.Descendants<Run>().First();
        var rPr = new RunProperties(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
            new Bold(),
            new FontSize { Val = "36" }
        );
        run.PrependChild(rPr);
    }

    private static void ApplyDateStyle(WordParagraph paragraph)
    {
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { After = "200" }
        );
        paragraph.PrependChild(pPr);

        var run = paragraph.Descendants<Run>().First();
        var rPr = new RunProperties(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
            new FontSize { Val = "26" }
        );
        run.PrependChild(rPr);
    }

    private static void ApplyFooterStyle(WordParagraph paragraph)
    {
        var pPr = new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new SpacingBetweenLines { Before = "200" }
        );
        paragraph.PrependChild(pPr);

        var run = paragraph.Descendants<Run>().First();
        var rPr = new RunProperties(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
            new Bold(),
            new Color { Val = "2C3E50" },
            new FontSize { Val = "26" }
        );
        run.PrependChild(rPr);
    }

    private static WordTableCell CreateHeaderCell(string text, string width)
    {
        var cell = new WordTableCell();

        var cellProps = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = width },
            new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "2C3E50" }
        );
        cell.Append(cellProps);

        var para = new WordParagraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center })
        );

        var run = new Run(
            new RunProperties(
                new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                new Bold(),
                new Color { Val = "FFFFFF" },
                new FontSize { Val = "22" }
            ),
            new Text(text)
        );

        para.Append(run);
        cell.Append(para);
        return cell;
    }

    private static WordTableCell CreateDataCell(string text, bool isBold, bool isAlternate, JustificationValues alignment, string width)
    {
        var cell = new WordTableCell();

        var cellProps = new TableCellProperties(
            new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = width },
            new Shading
            {
                Val = ShadingPatternValues.Clear,
                Color = "auto",
                Fill = isAlternate ? "F9F9F9" : "FFFFFF"
            }
        );
        cell.Append(cellProps);

        var para = new WordParagraph(
            new ParagraphProperties(new Justification { Val = alignment })
        );

        var runProps = new RunProperties(
            new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
            new FontSize { Val = "22" }
        );
        if (isBold) runProps.Append(new Bold());

        var run = new Run(runProps, new Text(text));
        para.Append(run);
        cell.Append(para);

        return cell;
    }
}