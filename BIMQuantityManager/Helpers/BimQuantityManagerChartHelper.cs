using BIMQuantityManager.Model;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace BIMQuantityManager.Helpers;

public sealed class BimQuantityManagerChartHelper
{
    static BimQuantityManagerChartHelper()
    {
        ExcelPackage.License.SetNonCommercialPersonal("BIM Quantity Manager");
    }

    public static void ExportPieChart(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        var fileInfo = new FileInfo(path);
        if (fileInfo.Exists) fileInfo.Delete();

        using (var package = new ExcelPackage(fileInfo))
        {
            var wsData = package.Workbook.Worksheets.Add("Data Table");
            CreateDataTable(wsData, typeData, selectedViews, activeViewName);

            // ============ SHEET 2: PIE CHARTS ============
            var wsChart = package.Workbook.Worksheets.Add("Pie Charts");
            CreatePieCharts(wsChart, wsData, typeData.Count(), selectedViews, activeViewName);

            package.Save();
        }
    }

    public static void ExportBarChart(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        var fileInfo = new FileInfo(path);
        if (fileInfo.Exists) fileInfo.Delete();

        using (var package = new ExcelPackage(fileInfo))
        {
            var ws = package.Workbook.Worksheets.Add("Bar Chart Analysis");
            int totalCols = 4 + selectedViews.Count;
            int dataCount = typeData.Count();

            CreateDataTable(ws, typeData, selectedViews, activeViewName);

            int headerRow = 1;
            int firstDataRow = 2;
            int lastDataRow = firstDataRow + dataCount - 1;
            int chartStartCol = totalCols + 2;

            var chart = ws.Drawings.AddBarChart("BarChart", eBarChartType.ColumnClustered);
            chart.Title.Text = "BIM Quantity Manager";
            chart.Title.Font.Bold = true;
            chart.Title.Font.Size = 14;

            chart.SetPosition(headerRow - 1, 0, chartStartCol, 0);

            int totalTableRows = (lastDataRow - headerRow + 1); 
            int rowHeightPixels = 26; 
            int tableHeight = totalTableRows * rowHeightPixels;

            chart.SetSize(1300, tableHeight);

            var series1 = chart.Series.Add(ws.Cells[firstDataRow, 3, lastDataRow, 3], ws.Cells[firstDataRow, 2, lastDataRow, 2]);
            series1.Header = "Total Quantity";

            var series2 = chart.Series.Add(ws.Cells[firstDataRow, 4, lastDataRow, 4], ws.Cells[firstDataRow, 2, lastDataRow, 2]);
            series2.Header = activeViewName;

            int col = 5;
            foreach (var view in selectedViews)
            {
                var series = chart.Series.Add(ws.Cells[firstDataRow, col, lastDataRow, col], ws.Cells[firstDataRow, 2, lastDataRow, 2]);
                series.Header = view.DisplayName;
                col++;
            }

            chart.DataLabel.ShowValue = false;
            chart.Legend.Position = eLegendPosition.Right;
            chart.Legend.Font.Size = 9;

            package.Save();
        }
    }

    public static void ExportLineChart(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        var fileInfo = new FileInfo(path);
        if (fileInfo.Exists) fileInfo.Delete();

        using (var package = new ExcelPackage(fileInfo))
        {
            var ws = package.Workbook.Worksheets.Add("Line Chart Analysis");
            int totalCols = 4 + selectedViews.Count;
            int dataCount = typeData.Count();

            CreateDataTable(ws, typeData, selectedViews, activeViewName);

            int headerRow = 1;
            int firstDataRow = 2;
            int lastDataRow = firstDataRow + dataCount - 1;
            int chartStartCol = totalCols + 2; 

            var chart = ws.Drawings.AddLineChart("LineChart", eLineChartType.Line);
            chart.Title.Text = "BIM Quantity Manager";
            chart.Title.Font.Bold = true;
            chart.Title.Font.Size = 14;

            chart.SetPosition(headerRow - 1, 0, chartStartCol, 0);

            int totalTableRows = (lastDataRow - headerRow + 1); 
            int rowHeightPixels = 26; 
            int tableHeight = totalTableRows * rowHeightPixels;

            chart.SetSize(1300, tableHeight);

            var series1 = chart.Series.Add(ws.Cells[firstDataRow, 3, lastDataRow, 3], ws.Cells[firstDataRow, 2, lastDataRow, 2]);
            series1.Header = "Total Quantity";

            var series2 = chart.Series.Add(ws.Cells[firstDataRow, 4, lastDataRow, 4], ws.Cells[firstDataRow, 2, lastDataRow, 2]);
            series2.Header = activeViewName;

            int col = 5;
            foreach (var view in selectedViews)
            {
                var series = chart.Series.Add(ws.Cells[firstDataRow, col, lastDataRow, col], ws.Cells[firstDataRow, 2, lastDataRow, 2]);
                series.Header = view.DisplayName;
                col++;
            }

            chart.DataLabel.ShowValue = false;
            chart.Legend.Position = eLegendPosition.Bottom;
            chart.Legend.Font.Size = 10;

            package.Save();
        }
    }

    private static void CreateDataTable(ExcelWorksheet ws, IEnumerable<TypeInfo> typeData, List<ViewInfo> selectedViews, string activeViewName)
    {
        int totalCols = 4 + selectedViews.Count;

        ws.Cells[1, 1, 1, totalCols].Merge = true;
        ws.Cells[1, 1].Value = "BIM QUANTITY MANAGER";
        ws.Cells[1, 1].Style.Font.Bold = true;
        ws.Cells[1, 1].Style.Font.Size = 20;
        ws.Cells[1, 1].Style.Font.Name = "Aptos Narrow";
        ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
        ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(41, 128, 185));
        ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        ws.Cells[1, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.FromArgb(41, 128, 185));
        ws.Row(1).Height = 35;

        ws.Cells[2, 1, 2, totalCols].Merge = true;
        ws.Cells[2, 1].Value = "Data Table - Complete Element Quantities";
        ws.Cells[2, 1].Style.Font.Size = 12;
        ws.Cells[2, 1].Style.Font.Name = "Aptos Narrow";
        ws.Cells[2, 1].Style.Font.Italic = true;
        ws.Cells[2, 1].Style.Font.Color.SetColor(Color.FromArgb(127, 140, 141));
        ws.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Cells[2, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(189, 195, 199));
        ws.Row(2).Height = 20;

        ws.Cells[3, 1, 3, totalCols].Merge = true;
        ws.Cells[3, 1].Value = $"Generated: {DateTime.Now:dd MMMM yyyy - HH:mm:ss}";
        ws.Cells[3, 1].Style.Font.Size = 10;
        ws.Cells[3, 1].Style.Font.Name = "Aptos Narrow";
        ws.Cells[3, 1].Style.Font.Color.SetColor(Color.FromArgb(149, 165, 166));
        ws.Cells[3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        ws.Cells[3, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(189, 195, 199));
        ws.Row(3).Height = 18;

        ws.Row(4).Height = 8;

        int col = 1;
        ws.Cells[5, col++].Value = "Category";
        ws.Cells[5, col++].Value = "Type Name";
        ws.Cells[5, col++].Value = "Total Quantity";
        ws.Cells[5, col++].Value = activeViewName;
        foreach (var view in selectedViews)
            ws.Cells[5, col++].Value = view.DisplayName;

        var headerRange = ws.Cells[5, 1, 5, totalCols];
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Font.Size = 12;
        headerRange.Style.Font.Name = "Aptos Narrow";
        headerRange.Style.Font.Color.SetColor(Color.White);
        headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
        headerRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(52, 73, 94));
        headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        headerRange.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.FromArgb(44, 62, 80));
        ws.Row(5).Height = 25;

        int row = 6;
        int dataCount = 0;

        foreach (var type in typeData)
        {
            col = 1;
            ws.Cells[row, col++].Value = type.Category ?? string.Empty;
            ws.Cells[row, col++].Value = type.TypeName ?? string.Empty;
            ws.Cells[row, col++].Value = type.TotalQuantity;
            ws.Cells[row, col++].Value = type.ActiveViewQuantity;

            foreach (var view in selectedViews)
            {
                long viewKey = view.ViewId.Value;
                int quantity = type.ViewQuantities.ContainsKey(viewKey) ? type.ViewQuantities[viewKey] : 0;
                ws.Cells[row, col++].Value = quantity;
            }

            var rowRange = ws.Cells[row, 1, row, totalCols];
            if (dataCount % 2 == 0)
            {
                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rowRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(236, 240, 241));
            }

            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 1].Style.Font.Name = "Aptos Narrow";
            ws.Cells[row, 1].Style.Font.Color.SetColor(Color.FromArgb(44, 62, 80));

            ws.Cells[row, 2].Style.Font.Name = "Aptos Narrow";
            ws.Cells[row, 2].Style.Font.Color.SetColor(Color.FromArgb(52, 73, 94));

            for (int c = 3; c <= totalCols; c++)
            {
                ws.Cells[row, c].Style.Font.Bold = true;
                ws.Cells[row, c].Style.Font.Name = "Aptos Narrow";
                ws.Cells[row, c].Style.Font.Color.SetColor(Color.FromArgb(41, 128, 185));
                ws.Cells[row, c].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            rowRange.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(189, 195, 199));
            row++;
            dataCount++;
        }

        ws.Column(1).Width = 25;
        ws.Column(2).Width = 55;
        ws.Column(3).Width = 18;
        ws.Column(4).Width = 28;
        for (int c = 5; c <= totalCols; c++)
            ws.Column(c).Width = 28;
    }

    private static void CreatePieCharts(ExcelWorksheet ws, ExcelWorksheet wsData, int dataCount, List<ViewInfo> selectedViews, string activeViewName)
    {
        int lastDataRow = 6 + dataCount - 1;
        int chartStartRow = 0;
        int chartStartCol = 0;
        int chartCount = 0;
        int chartsPerRow = 3; 
        int horizontalSpacing = 18; 
        int verticalSpacing = 50; 

        var chartTotal = ws.Drawings.AddPieChart($"PieChart_Total", ePieChartType.Pie);
        chartTotal.Title.Text = "Total Quantity";
        chartTotal.Title.Font.Bold = true;
        chartTotal.Title.Font.Size = 12;

        int totalCol = chartStartCol + (chartCount % chartsPerRow) * horizontalSpacing;
        int totalRow = chartStartRow + (chartCount / chartsPerRow) * verticalSpacing;
        chartTotal.SetPosition(totalRow, 0, totalCol, 0);
        chartTotal.SetSize(1000, 900);

        var seriesTotal = chartTotal.Series.Add(wsData.Cells[6, 3, lastDataRow, 3], wsData.Cells[6, 2, lastDataRow, 2]);
        seriesTotal.Header = "Total Quantity";
        chartTotal.DataLabel.ShowValue = false;
        chartTotal.DataLabel.ShowCategory = false;
        chartTotal.DataLabel.ShowPercent = false;
        chartTotal.Legend.Position = eLegendPosition.Right;
        chartTotal.Legend.Font.Size = 9;
        chartCount++;

        var chartActive = ws.Drawings.AddPieChart($"PieChart_Active", ePieChartType.Pie);
        chartActive.Title.Text = activeViewName;
        chartActive.Title.Font.Bold = true;
        chartActive.Title.Font.Size = 12;

        int activeCol = chartStartCol + (chartCount % chartsPerRow) * horizontalSpacing;
        int activeRow = chartStartRow + (chartCount / chartsPerRow) * verticalSpacing;
        chartActive.SetPosition(activeRow, 0, activeCol, 0);
        chartActive.SetSize(1000, 900);

        var seriesActive = chartActive.Series.Add(wsData.Cells[6, 4, lastDataRow, 4], wsData.Cells[6, 2, lastDataRow, 2]);
        seriesActive.Header = activeViewName;
        chartActive.DataLabel.ShowValue = false;
        chartActive.DataLabel.ShowCategory = false;
        chartActive.DataLabel.ShowPercent = false;
        chartActive.Legend.Position = eLegendPosition.Right;
        chartActive.Legend.Font.Size = 9;
        chartCount++;

        for (int i = 0; i < selectedViews.Count; i++)
        {
            var view = selectedViews[i];
            var chartView = ws.Drawings.AddPieChart($"PieChart_View{i}", ePieChartType.Pie);
            chartView.Title.Text = view.DisplayName;
            chartView.Title.Font.Bold = true;
            chartView.Title.Font.Size = 12;

            int viewCol = chartStartCol + (chartCount % chartsPerRow) * horizontalSpacing;
            int viewRow = chartStartRow + (chartCount / chartsPerRow) * verticalSpacing;
            chartView.SetPosition(viewRow, 0, viewCol, 0);
            chartView.SetSize(1000, 900);

            int colIndex = 5 + i;
            var seriesView = chartView.Series.Add(wsData.Cells[6, colIndex, lastDataRow, colIndex], wsData.Cells[6, 2, lastDataRow, 2]);
            seriesView.Header = view.DisplayName;
            chartView.DataLabel.ShowValue = false;
            chartView.DataLabel.ShowCategory = false;
            chartView.DataLabel.ShowPercent = false;
            chartView.Legend.Position = eLegendPosition.Right;
            chartView.Legend.Font.Size = 9;

            chartCount++;
        }
    }
}