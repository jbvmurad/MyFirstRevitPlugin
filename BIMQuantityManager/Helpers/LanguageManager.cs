using System.IO;
using Newtonsoft.Json;

namespace BIMQuantityManager.Helpers;

public sealed class LanguageManager
{
    #region Singleton Pattern

    private static LanguageManager? _instance;
    private static readonly object _lock = new object();

    public static LanguageManager Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    _instance ??= new LanguageManager();
                }
            }
            return _instance;
        }
    }

    #endregion

    #region Fields & Properties

    private Dictionary<string, string> _translations = new Dictionary<string, string>();
    private string _currentLanguage = "en-US";

    public event Action? LanguageChanged;

    public string CurrentLanguage
    {
        get => _currentLanguage;
        set
        {
            if (_currentLanguage != value && IsValidLanguageCode(value))
            {
                _currentLanguage = value;
                LoadLanguage(value);
                LanguageChanged?.Invoke();
            }
        }
    }

    public IReadOnlyList<string> SupportedLanguages { get; } = new List<string>
    {
        "en-US",  
        "de-DE",  
        "ru-RU"   
    };

    #endregion

    #region Constructor

    private LanguageManager()
    {
        LoadLanguage(_currentLanguage);
    }

    #endregion

    #region Public Methods


    public string Translate(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
            return string.Empty;

        return _translations.TryGetValue(key, out var value) ? value : key;
    }

    public string TranslateCategory(string category)
    {
        if (string.IsNullOrWhiteSpace(category))
            return string.Empty;

        string originalCategory = category;

        string normalizedKey = category.Replace(" ", "").Replace("-", "");
        string key = $"Category_{normalizedKey}";

        if (_translations.TryGetValue(key, out var translatedValue))
        {
            return translatedValue;
        }

        return originalCategory;
    }

    /// <summary>
    /// Gets the display name for a language code.
    /// </summary>
    /// <param name="langCode">Language code (e.g., "en-US").</param>
    /// <returns>Display name (e.g., "English").</returns>
    public string GetLanguageDisplayName(string langCode)
    {
        return langCode switch
        {
            "en-US" => "English",
            "de-DE" => "Deutsch",
            "ru-RU" => "Русский",
            _ => langCode
        };
    }

    #endregion

    #region Private Methods

    private void LoadLanguage(string langCode)
    {
        try
        {
            string? assemblyPath = Path.GetDirectoryName(typeof(LanguageManager).Assembly.Location);

            if (string.IsNullOrEmpty(assemblyPath))
            {
                _translations = GetDefaultEnglish();
                return;
            }

            string jsonPath = Path.Combine(assemblyPath, "Resources", "Languages", $"{langCode}.json");

            if (File.Exists(jsonPath))
            {
                string json = File.ReadAllText(jsonPath);

                _translations = JsonConvert.DeserializeObject<Dictionary<string, string>>(json)
                    ?? GetDefaultEnglish();
            }
            else
            {
                _translations = GetDefaultEnglish();
                System.Diagnostics.Debug.WriteLine($"Language file not found: {jsonPath}");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error loading language '{langCode}': {ex.Message}");
            _translations = GetDefaultEnglish();
        }
    }

    private bool IsValidLanguageCode(string langCode)
    {
        return SupportedLanguages.Contains(langCode);
    }

    private Dictionary<string, string> GetDefaultEnglish()
    {
        return new Dictionary<string, string>
        {
            { "Export", "Export" },
            { "Settings", "Settings" },
            { "Theme", "Theme" },
            { "Language", "Language" },
            { "DarkMode", "Dark Mode" },
            { "LightMode", "Light Mode" },
            { "SwitchToDarkMode", "Switch to Dark Mode" },
            { "SwitchToLightMode", "Switch to Light Mode" },
            { "Search", "Search" },
            { "SearchCategoryType", "Search Category / Type:" },
            { "SearchPlaceholder", "e.g., Walls, Doors, Windows..." },
            { "ActiveView", "Active View" },
            { "SelectedViews", "Selected Views" },
            { "SelectViews", "Select views..." },
            { "ViewsSelected", "view(s) selected" },
            { "TotalQuantity", "Total Quantity" },
            { "Category", "Category" },
            { "TypeName", "Type Name" },

            { "Category_Casework", "Casework" },
            { "Category_Ceilings", "Ceilings" },
            { "Category_Columns", "Columns" },
            { "Category_CurtainWallMullions", "Curtain Wall Mullions" },
            { "Category_CurtainWallPanels", "Curtain Wall Panels" },
            { "Category_Doors", "Doors" },
            { "Category_ElectricalEquipment", "Electrical Equipment" },
            { "Category_ElectricalFixtures", "Electrical Fixtures" },
            { "Category_Floors", "Floors" },
            { "Category_Furniture", "Furniture" },
            { "Category_FurnitureSystems", "Furniture Systems" },
            { "Category_GenericModel", "Generic Model" },
            { "Category_LightingFixtures", "Lighting Fixtures" },
            { "Category_MechanicalEquipment", "Mechanical Equipment" },
            { "Category_Parking", "Parking" },
            { "Category_Planting", "Planting" },
            { "Category_PlumbingFixtures", "Plumbing Fixtures" },
            { "Category_Railings", "Railings" },
            { "Category_Ramps", "Ramps" },
            { "Category_Roofs", "Roofs" },
            { "Category_Site", "Site" },
            { "Category_SpecialityEquipment", "Speciality Equipment" },
            { "Category_Stairs", "Stairs" },
            { "Category_StructuralColumns", "Structural Columns" },
            { "Category_StructuralFraming", "Structural Framing" },
            { "Category_Walls", "Walls" },
            { "Category_Windows", "Windows" },
            { "ExportSuccess", "Successfully exported!" },
            { "ExportFailed", "Export failed!" },
            { "NoDataAvailable", "No data available to export!" },
            { "FileLockedError", "File is locked or in use. Please close the file and try again." },
            { "ExportCompleted", "Export Completed" },
            { "DoYouWantToOpenFile", "Do you want to open the file now?" },
            { "File", "File" },
            { "Location", "Location" },
            { "Details", "Details" },
            { "ExportToExcel", "Export to Excel" },
            { "ExportToPdf", "Export to PDF" },
            { "ExportToWord", "Export to Word" },
            { "ExportToCsv", "Export to CSV" },
            { "ExportWithPieChart", "With Pie Chart" },
            { "ExportWithBarChart", "With Bar Chart" },
            { "ExportWithLineChart", "With Line Chart" },
            { "OnlyTable", "Only Table" },
            { "Warning", "Warning" },
            { "Error", "Error" },
            { "Success", "Success" },
            { "Cancel", "Cancel" },
            { "OK", "OK" },
            { "Yes", "Yes" },
            { "No", "No" }
        };
    }

    #endregion
}