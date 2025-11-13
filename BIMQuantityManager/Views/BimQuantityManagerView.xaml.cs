using BIMQuantityManager.Helpers;
using BIMQuantityManager.Model;
using BIMQuantityManager.ViewModels;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace BIMQuantityManager.Views
{
    public partial class TypeQuantitySummaryView : Window
    {
        public BimQuantityManagerViewModel ViewModel { get; set; }
        private bool isDarkMode = true;
        private LanguageManager languageManager;

        public TypeQuantitySummaryView(BimQuantityManagerViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
            DataContext = ViewModel;

            languageManager = LanguageManager.Instance;
            languageManager.LanguageChanged += OnLanguageChanged;

            ViewModel.SelectedViews.CollectionChanged += SelectedViews_CollectionChanged;

            SetDarkMode();

            // ✅ Window yüklendikten sonra icon'ları güncelle
            this.Loaded += Window_Loaded;

            ApplyTranslations();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateSettingsIcon(true);
            UpdateThemeIcon(true);
            UpdateLanguageIcon(true);
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                button.ContextMenu.HorizontalOffset = 0;
                button.ContextMenu.VerticalOffset = 2;
                button.ContextMenu.IsOpen = true;
            }
        }

        private void ThemeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (isDarkMode)
            {
                SetLightMode();
                UpdateSettingsIcon(false);
                UpdateThemeIcon(false);
                UpdateLanguageIcon(false);
            }
            else
            {
                SetDarkMode();
                UpdateSettingsIcon(true);
                UpdateThemeIcon(true);
                UpdateLanguageIcon(true);
            }

            isDarkMode = !isDarkMode;
            UpdateThemeMenuText();
        }

        private void LanguageMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is string langCode)
            {
                languageManager.CurrentLanguage = langCode;
            }
        }

        private void OnLanguageChanged()
        {
            ApplyTranslations();
        }

        private void ApplyTranslations()
        {
            if (ExportButton != null)
            {
                ExportButton.Content = $"{languageManager.Translate("Export")} ▾";
            }

            if (ExcelMenuItem != null) ExcelMenuItem.Header = languageManager.Translate("ExportToExcel");
            if (ExcelOnlyTableItem != null) ExcelOnlyTableItem.Header = languageManager.Translate("OnlyTable");
            if (ExcelPieChartItem != null) ExcelPieChartItem.Header = languageManager.Translate("ExportWithPieChart");
            if (ExcelLineChartItem != null) ExcelLineChartItem.Header = languageManager.Translate("ExportWithLineChart");
            if (ExcelBarChartItem != null) ExcelBarChartItem.Header = languageManager.Translate("ExportWithBarChart");
            if (PdfMenuItem != null) PdfMenuItem.Header = languageManager.Translate("ExportToPdf");
            if (WordMenuItem != null) WordMenuItem.Header = languageManager.Translate("ExportToWord");
            if (CsvMenuItem != null) CsvMenuItem.Header = languageManager.Translate("ExportToCsv");

            if (LanguageMenuItem != null) LanguageMenuItem.Header = languageManager.Translate("Language");

            if (ActiveViewLabel != null) ActiveViewLabel.Text = languageManager.Translate("ActiveView") + " ";
            if (SelectedViewsLabel != null) SelectedViewsLabel.Text = languageManager.Translate("SelectedViews") + ":";
            if (SearchLabel != null) SearchLabel.Text = languageManager.Translate("SearchCategoryType");

            if (CategoryColumn != null) CategoryColumn.Header = languageManager.Translate("Category");
            if (TypeNameColumn != null) TypeNameColumn.Header = languageManager.Translate("TypeName");
            if (TotalQuantityColumn != null) TotalQuantityColumn.Header = languageManager.Translate("TotalQuantity");

            UpdateComboBoxPlaceholder();
            UpdateSearchPlaceholder();

            ViewModel.RefreshData();
            UpdateThemeMenuText();
        }

        private void UpdateThemeMenuText()
        {
            if (ThemeMenuItem != null)
            {
                ThemeMenuItem.Header = isDarkMode
                    ? languageManager.Translate("SwitchToLightMode")
                    : languageManager.Translate("SwitchToDarkMode");
            }
        }

        private void UpdateSettingsIcon(bool isDark)
        {
            if (SettingsButton.Content is Grid grid)
            {
                var darkIcon = grid.FindName("SettingsIconDark") as Image;
                var lightIcon = grid.FindName("SettingsIconLight") as Image;

                if (darkIcon != null && lightIcon != null)
                {
                    if (isDark)
                    {
                        darkIcon.Opacity = 1;
                        lightIcon.Opacity = 0;
                    }
                    else
                    {
                        darkIcon.Opacity = 0;
                        lightIcon.Opacity = 1;
                    }
                }
            }
        }

        private void UpdateThemeIcon(bool isDark)
        {
            if (ThemeMenuItem?.Icon is Grid iconGrid)
            {
                Image? darkIcon = null;
                Image? lightIcon = null;

                foreach (var child in iconGrid.Children)
                {
                    if (child is Image img)
                    {
                        if (img.Source?.ToString()?.Contains("2.png") == true)
                        {
                            darkIcon = img;
                        }
                        else if (img.Source?.ToString()?.Contains("1.png") == true)
                        {
                            lightIcon = img;
                        }
                    }
                }

                if (darkIcon != null && lightIcon != null)
                {
                    if (isDark)
                    {
                        // Dark mode: 2.png göster
                        darkIcon.Visibility = Visibility.Visible;
                        lightIcon.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        // Light mode: 1.png göster
                        darkIcon.Visibility = Visibility.Collapsed;
                        lightIcon.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private void UpdateLanguageIcon(bool isDark)
        {
            if (LanguageMenuItem?.Icon is Grid iconGrid)
            {
                Image? darkIcon = null;
                Image? lightIcon = null;

                foreach (var child in iconGrid.Children)
                {
                    if (child is Image img)
                    {
                        if (img.Source?.ToString()?.Contains("LightGlobe.png") == true)
                        {
                            darkIcon = img;
                        }
                        else if (img.Source?.ToString()?.Contains("BlackGlobe.png") == true)
                        {
                            lightIcon = img;
                        }
                    }
                }

                if (darkIcon != null && lightIcon != null)
                {
                    if (isDark)
                    {
                        darkIcon.Visibility = Visibility.Visible;
                        lightIcon.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        darkIcon.Visibility = Visibility.Collapsed;
                        lightIcon.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private void SelectedViews_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateDynamicColumns();
        }

        private void ViewCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.DataContext is ViewInfo view)
            {
                if (checkBox.IsChecked == true)
                {
                    if (!ViewModel.SelectedViews.Contains(view))
                    {
                        ViewModel.SelectedViews.Add(view);
                    }
                }
                else
                {
                    ViewModel.SelectedViews.Remove(view);
                }

                ViewModel.RefreshData();
            }
        }

        private void UpdateDynamicColumns()
        {
            var columnsToRemove = TypeDataGrid.Columns
                .Where(c => c.Header is TextBlock tb && tb.Tag?.ToString() == "DynamicColumn")
                .ToList();

            foreach (var col in columnsToRemove)
            {
                TypeDataGrid.Columns.Remove(col);
            }

            foreach (var view in ViewModel.SelectedViews)
            {
                long viewKey = view.ViewId.Value;

                var column = new DataGridTemplateColumn
                {
                    Width = new DataGridLength(200, DataGridLengthUnitType.Pixel)
                };

                var headerText = new TextBlock
                {
                    Text = view.DisplayName,
                    FontWeight = FontWeights.SemiBold,
                    FontSize = 14,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextAlignment = TextAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    Padding = new Thickness(13, 0, 8, 12),
                    Tag = "DynamicColumn"
                };

                column.Header = headerText;

                var cellTemplate = new DataTemplate();
                var factory = new FrameworkElementFactory(typeof(TextBlock));

                factory.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                factory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                factory.SetValue(TextBlock.FontSizeProperty, 14.0);
                factory.SetResourceReference(TextBlock.ForegroundProperty, "TextPrimaryBrush");

                var binding = new Binding($"ViewQuantities[{viewKey}]")
                {
                    Mode = BindingMode.OneWay,
                    TargetNullValue = 0,
                    FallbackValue = 0
                };
                factory.SetBinding(TextBlock.TextProperty, binding);

                cellTemplate.VisualTree = factory;
                column.CellTemplate = cellTemplate;

                TypeDataGrid.Columns.Add(column);
            }

            TypeDataGrid.Items.Refresh();
        }

        private void ActiveViewComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems.Count > 0)
            {
                ViewModel.ChangeActiveViewInRevit();
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
                button.ContextMenu.HorizontalOffset = 0;
                button.ContextMenu.VerticalOffset = 2;
                button.ContextMenu.IsOpen = true;
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportData("Excel", ".xlsx", (data, path) =>
                BimQuantityManagerExcelHelper.ExportToExcel(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportToPdf_Click(object sender, RoutedEventArgs e)
        {
            ExportData("PDF", ".pdf", (data, path) =>
                BimQuantityManagerPdfHelper.ExportToPdf(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportToWord_Click(object sender, RoutedEventArgs e)
        {
            ExportData("Word", ".docx", (data, path) =>
                BimQuantityManagerWordHelper.ExportToWord(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportToCsv_Click(object sender, RoutedEventArgs e)
        {
            ExportData("CSV", ".csv", (data, path) =>
                BimQuantityManagerCsvHelper.ExportToCsv(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportToPieChart_Click(object sender, RoutedEventArgs e)
        {
            ExportData("Pie Chart", ".xlsx", (data, path) =>
                BimQuantityManagerChartHelper.ExportPieChart(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportToBarChart_Click(object sender, RoutedEventArgs e)
        {
            ExportData("Bar Chart", ".xlsx", (data, path) =>
                BimQuantityManagerChartHelper.ExportBarChart(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportToLineChart_Click(object sender, RoutedEventArgs e)
        {
            ExportData("Line Chart", ".xlsx", (data, path) =>
                BimQuantityManagerChartHelper.ExportLineChart(data, path, ViewModel.SelectedViews.ToList(), ViewModel.ActiveViewColumnHeader));
        }

        private void ExportData(string formatName, string extension, Action<IEnumerable<TypeInfo>, string> exportAction)
        {
            try
            {
                if (ViewModel.TypeData == null || ViewModel.TypeData.Count == 0)
                {
                    MessageBox.Show(
                        languageManager.Translate("NoDataAvailable"),
                        languageManager.Translate("Warning"),
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"BIM Quantity Manager_{DateTime.Now:yyyyMMdd_HHmmss}{extension}";
                string path = System.IO.Path.Combine(desktopPath, fileName);

                if (System.IO.File.Exists(path))
                {
                    try
                    {
                        System.IO.File.Delete(path);
                    }
                    catch
                    {
                        MessageBox.Show(
                            languageManager.Translate("FileLockedError"),
                            languageManager.Translate("Warning"),
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return;
                    }
                }

                exportAction(ViewModel.TypeData, path);

                var result = MessageBox.Show(
                    $"{languageManager.Translate("ExportSuccess")}\n\n" +
                    $"{languageManager.Translate("File")}: {fileName}\n" +
                    $"{languageManager.Translate("Location")}: Desktop\n\n" +
                    languageManager.Translate("DoYouWantToOpenFile"),
                    languageManager.Translate("ExportCompleted"),
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                string errorDetails = ex.InnerException != null
                    ? $"{ex.Message}\n\n{languageManager.Translate("Details")}: {ex.InnerException.Message}"
                    : ex.Message;

                MessageBox.Show(
                    $"{languageManager.Translate("Error")}:\n\n{errorDetails}",
                    languageManager.Translate("Error"),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SetLightMode()
        {
            LogoImage.Source = new BitmapImage(
              new Uri("pack://application:,,,/BIMQuantityManager;component/Logo/logo2.png", UriKind.Absolute)
            );

            this.Resources["WindowBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(245, 245, 245));
            this.Resources["CardBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["BorderBrush"] = new SolidColorBrush(Color.FromRgb(180, 180, 180));
            this.Resources["TextPrimaryBrush"] = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            this.Resources["DataGridRowBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["DataGridAlternateBrush"] = new SolidColorBrush(Color.FromRgb(248, 248, 248));
            this.Resources["GridLineBrush"] = new SolidColorBrush(Color.FromRgb(221, 221, 221));
            this.Resources["DataGridLineBrush"] = new SolidColorBrush(Colors.Black);
            this.Resources["ButtonBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0, 120, 215));
            this.Resources["ButtonForegroundBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["ButtonHoverBrush"] = new SolidColorBrush(Color.FromRgb(0, 90, 158));
            this.Resources["ComboBackgroundBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["ComboForegroundBrush"] = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            this.Resources["ComboBorderBrush"] = new SolidColorBrush(Color.FromRgb(180, 180, 180));
            this.Resources["ComboItemHoverBrush"] = new SolidColorBrush(Color.FromRgb(240, 240, 240));
            this.Resources["ComboItemSelectedBrush"] = new SolidColorBrush(Color.FromRgb(0, 120, 215));
            this.Resources["MenuBackgroundBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["MenuBorderBrush"] = new SolidColorBrush(Color.FromRgb(218, 220, 224));
            this.Resources["MenuItemHoverBrush"] = new SolidColorBrush(Color.FromRgb(243, 244, 246));
            this.Resources["MenuItemPressedBrush"] = new SolidColorBrush(Color.FromRgb(0, 120, 215));
            this.Resources["MenuSeparatorBrush"] = new SolidColorBrush(Color.FromRgb(229, 231, 235));

            LogoImage.Opacity = 1.0;

            ActiveViewComboBox.Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            ActiveViewComboBox.BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180));

            TypeDataGrid.Background = new SolidColorBrush(Colors.White);
            TypeDataGrid.Foreground = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            TypeDataGrid.RowBackground = new SolidColorBrush(Colors.White);
            TypeDataGrid.AlternatingRowBackground = new SolidColorBrush(Color.FromRgb(248, 248, 248));

            TypeDataGrid.Items.Refresh();
            TypeDataGrid.UpdateLayout();
            TypeDataGrid.InvalidateVisual();
        }

        private void SetDarkMode()
        {
            LogoImage.Source = new BitmapImage(
               new Uri("pack://application:,,,/BIMQuantityManager;component/Logo/logo.png", UriKind.Absolute)
            );

            this.Resources["WindowBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(30, 30, 30));
            this.Resources["CardBrush"] = new SolidColorBrush(Color.FromRgb(46, 46, 46));
            this.Resources["BorderBrush"] = new SolidColorBrush(Color.FromRgb(85, 85, 85));
            this.Resources["TextPrimaryBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["DataGridRowBrush"] = new SolidColorBrush(Color.FromRgb(46, 46, 46));
            this.Resources["DataGridAlternateBrush"] = new SolidColorBrush(Color.FromRgb(51, 51, 51));
            this.Resources["GridLineBrush"] = new SolidColorBrush(Color.FromRgb(85, 85, 85));
            this.Resources["DataGridLineBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["ButtonBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0, 122, 204));
            this.Resources["ButtonForegroundBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["ButtonHoverBrush"] = new SolidColorBrush(Color.FromRgb(0, 92, 180));
            this.Resources["ComboBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(46, 46, 46));
            this.Resources["ComboForegroundBrush"] = new SolidColorBrush(Colors.White);
            this.Resources["ComboBorderBrush"] = new SolidColorBrush(Color.FromRgb(85, 85, 85));
            this.Resources["ComboItemHoverBrush"] = new SolidColorBrush(Color.FromRgb(58, 58, 58));
            this.Resources["ComboItemSelectedBrush"] = new SolidColorBrush(Color.FromRgb(0, 122, 204));
            this.Resources["MenuBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(45, 45, 48));
            this.Resources["MenuBorderBrush"] = new SolidColorBrush(Color.FromRgb(63, 63, 70));
            this.Resources["MenuItemHoverBrush"] = new SolidColorBrush(Color.FromRgb(62, 62, 66));
            this.Resources["MenuItemPressedBrush"] = new SolidColorBrush(Color.FromRgb(0, 122, 204));
            this.Resources["MenuSeparatorBrush"] = new SolidColorBrush(Color.FromRgb(63, 63, 70));

            LogoImage.Opacity = 1.0;

            ActiveViewComboBox.Foreground = new SolidColorBrush(Colors.White);
            ActiveViewComboBox.BorderBrush = new SolidColorBrush(Color.FromRgb(85, 85, 85));

            TypeDataGrid.Background = new SolidColorBrush(Color.FromRgb(46, 46, 46));
            TypeDataGrid.Foreground = new SolidColorBrush(Colors.White);
            TypeDataGrid.RowBackground = new SolidColorBrush(Color.FromRgb(46, 46, 46));
            TypeDataGrid.AlternatingRowBackground = new SolidColorBrush(Color.FromRgb(51, 51, 51));

            TypeDataGrid.Items.Refresh();
            TypeDataGrid.UpdateLayout();
            TypeDataGrid.InvalidateVisual();
        }

        private void UpdateComboBoxPlaceholder()
        {
            // ViewComboBox'un template'inden placeholder'ı bul ve güncelle
            if (ViewComboBox != null && ViewComboBox.Template != null)
            {
                try
                {
                    var toggleButton = ViewComboBox.Template.FindName("ToggleButton", ViewComboBox) as System.Windows.Controls.Primitives.ToggleButton;
                    if (toggleButton != null && toggleButton.Template != null)
                    {
                        var selectViewsPlaceholder = toggleButton.Template.FindName("SelectViewsPlaceholder", toggleButton) as TextBlock;
                        if (selectViewsPlaceholder != null)
                        {
                            selectViewsPlaceholder.Text = languageManager.Translate("SelectViews");
                        }
                    }
                }
                catch { }
            }
        }

        private void UpdateSearchPlaceholder()
        {
            try
            {
                var textBoxes = FindVisualChildren<TextBox>(this);
                foreach (var textBox in textBoxes)
                {
                    if (textBox?.Template != null)
                    {
                        var placeholder = textBox.Template.FindName("SearchPlaceholder", textBox) as TextBlock;
                        if (placeholder != null)
                        {
                            placeholder.Text = languageManager.Translate("SearchPlaceholder");
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject? depObj) where T : DependencyObject
        {
            if (depObj == null)
                yield break;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                DependencyObject? child = VisualTreeHelper.GetChild(depObj, i);
                if (child != null && child is T typedChild)
                {
                    yield return typedChild;
                }

                if (child != null)
                {
                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }
    }
}