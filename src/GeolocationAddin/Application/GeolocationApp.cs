using System;
using System.Reflection;
using Autodesk.Revit.UI;

namespace GeolocationAddin.Application
{
    public class GeolocationApp : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                var tabName = "Geolocation";
                application.CreateRibbonTab(tabName);

                var panel = application.CreateRibbonPanel(tabName, "Tools");

                var assemblyPath = Assembly.GetExecutingAssembly().Location;
                var buttonData = new PushButtonData(
                    "GeolocationCommand",
                    "Run\nGeolocation",
                    assemblyPath,
                    "GeolocationAddin.Commands.GeolocationCommand");

                buttonData.ToolTip = "Copy linked models, publish shared coordinates, and export to IFC/NWC/DWG";

                panel.AddItem(buttonData);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Geolocation Addin", $"Failed to initialize: {ex.Message}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
