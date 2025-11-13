using Autodesk.Revit.DB;

namespace BIMQuantityManager.Model;

public sealed class TypeInfo
{
    public required string Category { get; set; }
    public required string TypeName { get; set; }
    public int TotalQuantity { get; set; }
    public int ActiveViewQuantity { get; set; }

    public Dictionary<long, int> ViewQuantities { get; set; } = new Dictionary<long, int>();

    public int SelectedViewQuantity { get; set; }
    public string ViewName { get; set; } = string.Empty;
}