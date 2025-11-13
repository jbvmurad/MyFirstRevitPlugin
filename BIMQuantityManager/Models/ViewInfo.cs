using Autodesk.Revit.DB;

namespace BIMQuantityManager.Model;

public sealed class ViewInfo
{
    public required string Name { get; set; }
    public required ElementId ViewId { get; set; }
    public ViewType ViewType { get; set; }

    public string DisplayName => $"{Name} ({GetViewTypeString()})";

    private string GetViewTypeString()
    {
        return ViewType switch
        {
            ViewType.FloorPlan => "Floor Plan",
            ViewType.CeilingPlan => "Ceiling Plan",
            ViewType.Elevation => "Elevation",
            ViewType.Section => "Section",
            ViewType.ThreeD => "3D View",
            ViewType.Schedule => "Schedule",
            _ => ViewType.ToString()
        };
    }
}