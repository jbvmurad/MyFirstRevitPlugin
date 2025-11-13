using BIMQuantityManager.Model;
using System.Text;
using System.IO;

namespace BIMQuantityManager.Helpers;

public sealed class BimQuantityManagerCsvHelper
{
    public static void ExportToCsv(IEnumerable<TypeInfo> typeData, string path, List<ViewInfo> selectedViews, string activeViewName = "Active View")
    {
        var sb = new StringBuilder();

        sb.Append($"Category,Type Name,Total Quantity,{EscapeCsvField(activeViewName)}");

        foreach (var view in selectedViews)
        {
            sb.Append($",{EscapeCsvField(view.DisplayName)}");
        }
        sb.AppendLine();

        foreach (var type in typeData)
        {
            var category = EscapeCsvField(type.Category ?? string.Empty);
            var typeName = EscapeCsvField(type.TypeName ?? string.Empty);

            sb.Append($"{category},{typeName},{type.TotalQuantity},{type.ActiveViewQuantity}");

            foreach (var view in selectedViews)
            {
                long viewKey = view.ViewId.Value;
                int quantity = type.ViewQuantities.ContainsKey(viewKey)
                    ? type.ViewQuantities[viewKey]
                    : 0;
                sb.Append($",{quantity}");
            }

            sb.AppendLine();
        }

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
    }

    private static string EscapeCsvField(string field)
    {
        if (string.IsNullOrEmpty(field))
            return string.Empty;

        if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }

        return field;
    }
}