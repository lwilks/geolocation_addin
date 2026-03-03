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

            // Use Application.OpenDocumentFile (not UIApplication.OpenAndActivateDocument)
            // so the document opens in the background without becoming the active UI document.
            // This allows us to close it later without hitting "active document may not be closed".
            return uiApp.Application.OpenDocumentFile(modelPath, openOptions);
        }

        public static Document OpenCloudDocumentDetached(UIApplication uiApp, string region, Guid projectGuid, Guid modelGuid)
        {
            // Use cloud GUIDs to construct a proper cloud ModelPath
            var modelPath = ModelPathUtils.ConvertCloudGUIDsToCloudPath(region, projectGuid, modelGuid);
            var openOptions = new OpenOptions();

            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            var worksetConfig = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksetConfig);

            LogHelper.Info($"Opening cloud model: region={region}, project={projectGuid}, model={modelGuid}");
            // Use Application.OpenDocumentFile to open in background (not as active UI document)
            return uiApp.Application.OpenDocumentFile(modelPath, openOptions);
        }

        public static void SaveDocumentAs(Document doc, string targetPath)
        {
            var saveAsOptions = new SaveAsOptions
            {
                OverwriteExistingFile = true,
                MaximumBackups = 1
            };

            // Detached workshared documents require SaveAsCentral = true
            if (doc.IsWorkshared)
            {
                var wsOptions = new WorksharingSaveAsOptions { SaveAsCentral = true };
                saveAsOptions.SetWorksharingOptions(wsOptions);
            }

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

                    if (doc.IsWorkshared)
                    {
                        var wsOptions = new WorksharingSaveAsOptions { SaveAsCentral = true };
                        saveOptions.SetWorksharingOptions(wsOptions);
                    }

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
