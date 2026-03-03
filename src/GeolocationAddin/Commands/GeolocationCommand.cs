using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using GeolocationAddin.Config;
using GeolocationAddin.Core;
using GeolocationAddin.Helpers;

namespace GeolocationAddin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class GeolocationCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiApp = commandData.Application;

            try
            {
                var config = ConfigLoader.Load();
                var mapping = CsvMapping.Load(config.CsvMappingPath);

                LogHelper.Info($"Config loaded. Site model: {config.SiteModelPath}");
                LogHelper.Info($"CSV mapping loaded: {mapping.EntryCount} entries");

                var workflow = new GeolocationWorkflow(uiApp, config, mapping);
                workflow.Execute();

                return Result.Succeeded;
            }
            catch (ConfigurationException ex)
            {
                LogHelper.Error($"Configuration error: {ex.Message}");
                TaskDialog.Show("Geolocation — Config Error", ex.Message);
                return Result.Failed;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Unexpected error: {ex}");
                TaskDialog.Show("Geolocation — Error", $"An unexpected error occurred:\n\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
