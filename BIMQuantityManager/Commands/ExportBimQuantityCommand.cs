using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using BIMQuantityManager.Model;
using BIMQuantityManager.ViewModels;
using RevitTaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace BIMQuantityManager.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ExportBimQuantityCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;

                List<ViewInfo> allViews = GetAllViews(doc);

                if (allViews.Count == 0)
                {
                    RevitTaskDialog.Show("Warning", "Heç bir uyğun view tapılmadı!");
                    return Result.Cancelled;
                }

                var viewModel = new BimQuantityManagerViewModel(uiDoc, allViews);
                var window = new BIMQuantityManager.Views.TypeQuantitySummaryView(viewModel);

                commandData.Application.ViewActivated += (sender, e) => OnViewActivated(e, viewModel);

                commandData.Application.Application.DocumentChanged += (sender, args) =>
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        viewModel.RefreshData();
                    });
                };

                var helper = new System.Windows.Interop.WindowInteropHelper(window);
                helper.Owner = commandData.Application.MainWindowHandle;

                window.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                RevitTaskDialog.Show("Error", $"Xəta baş verdi:\n{ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private void OnViewActivated(ViewActivatedEventArgs e, BimQuantityManagerViewModel viewModel)
        {
            try
            {
                var newActiveView = e.CurrentActiveView;
                if (newActiveView == null) return;

                var matchingView = viewModel.AllViews.FirstOrDefault(v => v.ViewId == newActiveView.Id);

                if (matchingView != null && matchingView.ViewId != viewModel.ActiveView?.ViewId)
                {
                    System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                    {
                        viewModel.UpdateActiveViewFromRevit(matchingView);
                    });
                }
            }
            catch { }
        }

        private List<ViewInfo> GetAllViews(Document doc)
        {
            var views = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && v.CanBePrinted)
                .Where(v => v.ViewType == ViewType.FloorPlan ||
                            v.ViewType == ViewType.CeilingPlan ||
                            v.ViewType == ViewType.Elevation ||
                            v.ViewType == ViewType.Section ||
                            v.ViewType == ViewType.ThreeD)
                .OrderBy(v => v.ViewType)
                .ThenBy(v => v.Name)
                .Select(v => new ViewInfo
                {
                    Name = v.Name,
                    ViewId = v.Id,
                    ViewType = v.ViewType
                })
                .ToList();

            return views;
        }
    }
}
