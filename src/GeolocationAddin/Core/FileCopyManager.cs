using System;
using System.IO;
using Autodesk.Revit.DB;
using GeolocationAddin.Helpers;

namespace GeolocationAddin.Core
{
    public static class FileCopyManager
    {
        public static string ResolveLinkFilePath(RevitLinkType linkType)
        {
            var extRef = linkType.GetExternalFileReference();
            if (extRef == null || extRef.GetLinkedFileStatus() == LinkedFileStatus.NotFound)
                return null;

            var modelPath = extRef.GetAbsolutePath();
            if (modelPath == null)
                return null;

            // For local files, convert ModelPath to a string path
            if (ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath) is string path
                && !string.IsNullOrEmpty(path))
            {
                return path;
            }

            return null;
        }

        public static string CopyLinkedModel(string sourceFilePath, string targetFileName, string outputFolder)
        {
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException($"Source linked model not found: {sourceFilePath}");

            // Ensure target has .rvt extension
            if (!targetFileName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                targetFileName += ".rvt";

            var targetPath = Path.Combine(outputFolder, targetFileName);

            LogHelper.Info($"Copying: {Path.GetFileName(sourceFilePath)} -> {targetFileName}");
            File.Copy(sourceFilePath, targetPath, overwrite: true);

            return targetPath;
        }
    }
}
