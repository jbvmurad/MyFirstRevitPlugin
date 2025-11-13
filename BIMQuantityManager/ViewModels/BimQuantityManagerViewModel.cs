using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BIMQuantityManager.Model;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace BIMQuantityManager.ViewModels;

public class BimQuantityManagerViewModel : INotifyPropertyChanged
{
    private Document _document;
    private UIDocument _uiDocument;
    private ViewInfo _activeView;
    private ObservableCollection<ViewInfo> _selectedViews = new();
    private ObservableCollection<TypeInfo> _typeData = new();
    private ObservableCollection<TypeInfo> _allTypeData = new();
    private string _searchText = string.Empty;
    private string _activeViewName = string.Empty;
    private string _activeViewColumnHeader = "Active View";
    private bool _isUpdatingFromRevit = false;

    public ObservableCollection<ViewInfo> AllViews { get; set; } = new();
    public ObservableCollection<ViewInfo> AvailableViews { get; set; } = new();

    public string ActiveViewName
    {
        get => _activeViewName;
        set
        {
            _activeViewName = value;
            OnPropertyChanged();
        }
    }

    public ViewInfo ActiveView
    {
        get => _activeView;
        set
        {
            if (_activeView != value && value != null)
            {
                _activeView = value;
                OnPropertyChanged();

                ActiveViewName = _activeView.DisplayName;
                ActiveViewColumnHeader = _activeView.DisplayName;

                UpdateAvailableViews();
                RefreshData();
            }
        }
    }

    public string ActiveViewColumnHeader
    {
        get => _activeViewColumnHeader;
        set
        {
            _activeViewColumnHeader = value;
            OnPropertyChanged();
        }
    }

    public ObservableCollection<ViewInfo> SelectedViews
    {
        get => _selectedViews;
        set
        {
            _selectedViews = value;
            OnPropertyChanged();
            RefreshData();
        }
    }

    public ObservableCollection<TypeInfo> TypeData
    {
        get => _typeData;
        set
        {
            _typeData = value;
            OnPropertyChanged();
        }
    }

    public string SearchText
    {
        get => _searchText;
        set
        {
            if (_searchText != value)
            {
                _searchText = value;
                OnPropertyChanged();
                ApplyFilter();
            }
        }
    }

    public BimQuantityManagerViewModel(UIDocument uiDoc, List<ViewInfo> allViews)
    {
        _uiDocument = uiDoc;
        _document = uiDoc.Document;

        AllViews = new ObservableCollection<ViewInfo>(allViews);

        var currentActiveView = allViews.FirstOrDefault(v => v.ViewId == uiDoc.ActiveView.Id);
        if (currentActiveView == null)
        {
            currentActiveView = new ViewInfo
            {
                Name = uiDoc.ActiveView.Name,
                ViewId = uiDoc.ActiveView.Id,
                ViewType = uiDoc.ActiveView.ViewType
            };
            AllViews.Insert(0, currentActiveView);
        }

        _activeView = currentActiveView;
        ActiveViewName = _activeView.DisplayName;
        ActiveViewColumnHeader = _activeView.DisplayName;

        UpdateAvailableViews();

        _selectedViews = new ObservableCollection<ViewInfo>();
        TypeData = new ObservableCollection<TypeInfo>();

        RefreshData();
    }

    public void UpdateActiveViewFromRevit(ViewInfo newActiveView)
    {
        if (_isUpdatingFromRevit) return;

        try
        {
            _isUpdatingFromRevit = true;
            ActiveView = newActiveView;
        }
        finally
        {
            _isUpdatingFromRevit = false;
        }
    }

    private void UpdateAvailableViews()
    {
        var filtered = AllViews.Where(v => v.ViewId != _activeView.ViewId).ToList();
        AvailableViews = new ObservableCollection<ViewInfo>(filtered);
        OnPropertyChanged(nameof(AvailableViews));

        var toRemove = _selectedViews.Where(v => v.ViewId == _activeView.ViewId).ToList();
        foreach (var v in toRemove)
        {
            _selectedViews.Remove(v);
        }
    }

    public void ChangeActiveViewInRevit()
    {
        try
        {
            if (_activeView == null || _isUpdatingFromRevit) return;

            _uiDocument.ActiveView = _document.GetElement(_activeView.ViewId) as View;
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                $"Active View dəyişdirilə bilmədi:\n{ex.Message}",
                "Xəta",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
        }
    }

    public void RefreshData()
    {
        if (_document == null)
        {
            TypeData = new ObservableCollection<TypeInfo>();
            return;
        }

        var types = CollectTypeQuantities();
        _allTypeData = new ObservableCollection<TypeInfo>(types);

        ApplyFilter();
    }

    private void ApplyFilter()
    {
        if (_allTypeData == null)
            return;

        if (string.IsNullOrWhiteSpace(_searchText))
        {
            TypeData = new ObservableCollection<TypeInfo>(_allTypeData);
            return;
        }

        var searchLower = _searchText.ToLower().Trim();

        var filtered = _allTypeData
            .Where(t =>
                t.Category.ToLower().Contains(searchLower) ||
                t.TypeName.ToLower().Contains(searchLower))
            .ToList();

        TypeData = new ObservableCollection<TypeInfo>(filtered);
    }

    private List<TypeInfo> CollectTypeQuantities()
    {
        Dictionary<string, TypeQuantityData> typeDict = new Dictionary<string, TypeQuantityData>();

        var activeViewElement = _document.GetElement(_activeView.ViewId) as View;
        if (activeViewElement == null) return new List<TypeInfo>();

        var activeViewElements = new FilteredElementCollector(_document, activeViewElement.Id)
            .WhereElementIsNotElementType()
            .Where(e => IsValidModelElement(e))
            .ToList();

        var allElements = new FilteredElementCollector(_document)
            .WhereElementIsNotElementType()
            .Where(e => IsValidModelElement(e))
            .ToList();

        CountElements(allElements, typeDict, CountType.Total);
        CountElements(activeViewElements, typeDict, CountType.ActiveView);

        if (_selectedViews != null && _selectedViews.Count > 0)
        {
            foreach (var kvp in typeDict)
            {
                foreach (var selectedView in _selectedViews)
                {
                    if (!kvp.Value.ViewQuantities.ContainsKey(selectedView.ViewId))
                    {
                        kvp.Value.ViewQuantities[selectedView.ViewId] = 0;
                    }
                }
            }

            foreach (var selectedView in _selectedViews)
            {
                var selectedViewElements = new FilteredElementCollector(_document, selectedView.ViewId)
                    .WhereElementIsNotElementType()
                    .Where(e => IsValidModelElement(e))
                    .ToList();

                CountElementsForView(selectedViewElements, typeDict, selectedView.ViewId);
            }
        }

        var languageManager = Helpers.LanguageManager.Instance;

        var orderedList = typeDict.Select(kvp => new TypeInfo
        {
            Category = languageManager.TranslateCategory(kvp.Value.Category ?? string.Empty),
            TypeName = kvp.Value.TypeName ?? string.Empty,
            TotalQuantity = kvp.Value.TotalCount,
            ActiveViewQuantity = kvp.Value.ActiveViewCount,
            ViewQuantities = kvp.Value.ViewQuantities.ToDictionary(
                kv => kv.Key.Value,
                kv => kv.Value
            ),
            SelectedViewQuantity = kvp.Value.ViewQuantities.Values.FirstOrDefault(),
            ViewName = _selectedViews?.FirstOrDefault()?.Name ?? string.Empty
        })
        .OrderBy(t => t.Category)
        .ThenBy(t => t.TypeName)
        .ToList();

        return orderedList;
    }

    private bool IsValidModelElement(Element elem)
    {
        if (elem == null || elem.Category == null)
            return false;

        var category = elem.Category;

        if (category.CategoryType != CategoryType.Model)
            return false;

        BuiltInCategory builtInCat;
        try
        {
            builtInCat = (BuiltInCategory)category.Id.Value;
        }
        catch
        {
            return false;
        }

        var allowedCategories = new HashSet<BuiltInCategory>
        {
            BuiltInCategory.OST_Walls,
            BuiltInCategory.OST_Floors,
            BuiltInCategory.OST_Roofs,
            BuiltInCategory.OST_Ceilings,
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Furniture,
            BuiltInCategory.OST_FurnitureSystems,
            BuiltInCategory.OST_Stairs,
            BuiltInCategory.OST_Ramps,
            BuiltInCategory.OST_Railings,
            BuiltInCategory.OST_CurtainWallPanels,
            BuiltInCategory.OST_CurtainWallMullions,
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_SpecialityEquipment,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_ElectricalFixtures,
            BuiltInCategory.OST_LightingFixtures,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_Casework,
            BuiltInCategory.OST_Site,
            BuiltInCategory.OST_Parking,
            BuiltInCategory.OST_Planting
        };

        return allowedCategories.Contains(builtInCat);
    }

    private void CountElements(List<Element> elements, Dictionary<string, TypeQuantityData> typeDict, CountType countType)
    {
        foreach (var elem in elements)
        {
            string category = elem.Category?.Name ?? "Unknown";
            string typeName = GetElementTypeName(elem);

            if (string.IsNullOrEmpty(typeName))
                continue;

            string key = $"{category}|{typeName}";

            if (!typeDict.ContainsKey(key))
            {
                typeDict[key] = new TypeQuantityData
                {
                    Category = category,
                    TypeName = typeName
                };
            }

            switch (countType)
            {
                case CountType.Total:
                    typeDict[key].TotalCount++;
                    break;
                case CountType.ActiveView:
                    typeDict[key].ActiveViewCount++;
                    break;
            }
        }
    }

    private void CountElementsForView(List<Element> elements, Dictionary<string, TypeQuantityData> typeDict, ElementId viewId)
    {
        foreach (var elem in elements)
        {
            string category = elem.Category?.Name ?? "Unknown";
            string typeName = GetElementTypeName(elem);

            if (string.IsNullOrEmpty(typeName))
                continue;

            string key = $"{category}|{typeName}";

            if (!typeDict.ContainsKey(key))
            {
                typeDict[key] = new TypeQuantityData
                {
                    Category = category,
                    TypeName = typeName
                };
            }

            if (!typeDict[key].ViewQuantities.ContainsKey(viewId))
            {
                typeDict[key].ViewQuantities[viewId] = 0;
            }

            typeDict[key].ViewQuantities[viewId]++;
        }
    }

    private string GetElementTypeName(Element elem)
    {
        try
        {
            if (elem is FamilyInstance familyInstance)
            {
                FamilySymbol? symbol = familyInstance.Symbol;
                if (symbol != null)
                {
                    return $"{symbol.FamilyName} : {symbol.Name}";
                }
            }

            ElementId typeId = elem.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                ElementType? elemType = _document.GetElement(typeId) as ElementType;
                if (elemType != null)
                {
                    return elemType.Name;
                }
            }

            return elem.Name ?? "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private enum CountType
    {
        Total,
        ActiveView
    }

    private class TypeQuantityData
    {
        public string Category { get; set; } = string.Empty;
        public string TypeName { get; set; } = string.Empty;
        public int TotalCount { get; set; }
        public int ActiveViewCount { get; set; }
        public Dictionary<ElementId, int> ViewQuantities { get; set; } = new Dictionary<ElementId, int>();
    }
}