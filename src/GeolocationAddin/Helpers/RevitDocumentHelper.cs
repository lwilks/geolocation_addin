using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace GeolocationAddin.Helpers
{
    public static class RevitDocumentHelper
    {
        public static Document OpenDocument(UIApplication uiApp, string filePath)
        {
            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            var openOptions = new OpenOptions();

            var uiDoc = uiApp.OpenAndActivateDocument(modelPath, openOptions, false);
            return uiDoc.Document;
        }

        public static Document OpenDocumentDetached(UIApplication uiApp, string filePath)
        {
            var modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            var openOptions = new OpenOptions();

            // Detach from central if workshared
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            // Open all worksets
            var worksetConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksetConfig);

            var uiDoc = uiApp.OpenAndActivateDocument(modelPath, openOptions, false);
            return uiDoc.Document;
        }

        public static Document OpenCloudDocumentDetached(UIApplication uiApp, Guid projectGuid, Guid modelGuid)
        {
            var cloudPath = ModelPathUtils.ConvertCloudPath(projectGuid, modelGuid);
            var openOptions = new OpenOptions();

            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            var worksetConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksetConfig);

            LogHelper.Info($"Opening cloud model — project: {projectGuid}, model: {modelGuid}");
            var uiDoc = uiApp.OpenAndActivateDocument(cloudPath, openOptions, false);
            return uiDoc.Document;
        }

        public static void SaveDocumentAs(Document doc, string targetPath)
        {
            var saveAsOptions = new SaveAsOptions
            {
                OverwriteExistingFile = true,
                MaximumBackups = 1
            };
            doc.SaveAs(targetPath, saveAsOptions);
        }

        public static void CloseDocument(Document doc, bool save)
        {
            if (doc == null || !doc.IsValidObject)
                return;

            try
            {
                if (save)
                {
                    var saveOptions = new SaveAsOptions
                    {
                        OverwriteExistingFile = true,
                        MaximumBackups = 1
                    };
                    doc.SaveAs(doc.PathName, saveOptions);
                }

                doc.Close(false);
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Error closing document: {ex.Message}");
                try { doc.Close(false); } catch { }
            }
        }
    }
}
