using System;
using Autodesk.Revit.DB;

namespace GeolocationAddin.Helpers
{
    public static class RevitDocumentHelper
    {
        public static Document OpenDocument(Autodesk.Revit.ApplicationServices.Application app, string filePath)
        {
            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            var openOptions = new OpenOptions();

            return app.OpenDocumentFile(modelPath, openOptions);
        }

        public static Document OpenDocumentDetached(Autodesk.Revit.ApplicationServices.Application app, string filePath)
        {
            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            var openOptions = new OpenOptions();

            // Detach from central if workshared
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            // Open all worksets
            var worksetConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksetConfig);

            return app.OpenDocumentFile(modelPath, openOptions);
        }

        public static void CloseDocument(Document doc, bool save)
        {
            if (doc == null || !doc.IsValidObject)
                return;

            try
            {
                if (save && doc.IsModifiable)
                {
                    // Use SaveAs for detached documents
                    var saveOptions = new SaveAsOptions { OverwriteExistingFile = true };
                    doc.SaveAs(doc.PathName, saveOptions);
                }

                doc.Close(false); // false = don't prompt to save
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Error closing document: {ex.Message}");
                try { doc.Close(false); } catch { }
            }
        }
    }
}
