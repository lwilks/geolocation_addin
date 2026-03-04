using System;
using System.Windows.Interop;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using GeolocationAddin.Config;
using GeolocationAddin.Helpers;
using GeolocationAddin.UI;

namespace GeolocationAddin.Commands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class SettingsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var config = ConfigLoader.Load();
                var window = new SettingsWindow(config);

                // Set Revit main window as owner
                var revitHandle = commandData.Application.MainWindowHandle;
                new WindowInteropHelper(window).Owner = revitHandle;

                if (window.ShowDialog() == true)
                {
                    ConfigLoader.Save(window.Config);
                    LogHelper.Info("Settings saved.");
                    TaskDialog.Show("Geolocation", "Settings saved successfully.");
                }

                return Result.Succeeded;
            }
            catch (ConfigurationException ex)
            {
                LogHelper.Error($"Settings error: {ex.Message}");
                TaskDialog.Show("Geolocation — Error", $"Could not load settings:\n\n{ex.Message}");
                return Result.Failed;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Settings error: {ex}");
                TaskDialog.Show("Geolocation — Error", $"An error occurred:\n\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
