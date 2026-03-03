using System;
using System.IO;
using Autodesk.Revit.DB;
using GeolocationAddin.Helpers;

namespace GeolocationAddin.Core
{
    public static class FileCopyManager
    {
        public static string ResolveLinkFilePath(RevitLinkType linkType, string linkSourceFolder)
        {
            var typeName = linkType.Name;

            // Strategy 1: Traditional external file reference (local file links)
            try
            {
                var extRef = linkType.GetExternalFileReference();
                if (extRef != null && extRef.GetLinkedFileStatus() != LinkedFileStatus.NotFound)
                {
                    var modelPath = extRef.GetAbsolutePath();
                    if (modelPath != null)
                    {
                        var path = ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPath);
                        if (!string.IsNullOrEmpty(path) && File.Exists(path))
                        {
                            LogHelper.Info($"Resolved via external file reference: {path}");
                            return path;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Info($"Link type '{typeName}' is not a traditional file link ({ex.Message})");
            }

            // Strategy 2: External resource references (ACC/BIM360 via Desktop Connector)
            try
            {
                var extResources = linkType.GetExternalResourceReferences();
                if (extResources != null && extResources.Count > 0)
                {
                    LogHelper.Info($"Found {extResources.Count} external resource reference(s) for '{typeName}'");
                    foreach (var kvp in extResources)
                    {
                        var resourceRef = kvp.Value;

                        var inSessionPath = resourceRef.InSessionPath;
                        LogHelper.Info($"  InSessionPath: '{inSessionPath}'");
                        if (!string.IsNullOrEmpty(inSessionPath) && File.Exists(inSessionPath))
                            return inSessionPath;

                        var refMap = resourceRef.GetReferenceInformation();
                        if (refMap != null)
                        {
                            foreach (var info in refMap)
                            {
                                LogHelper.Info($"  RefInfo: {info.Key} = '{info.Value}'");
                                if (File.Exists(info.Value))
                                    return info.Value;
                            }
                        }
                    }
                }
                else
                {
                    LogHelper.Info($"No external resource references found for '{typeName}'");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"External resource resolution failed for '{typeName}': {ex.Message}");
            }

            // Strategy 3: Search linkSourceFolder from config
            if (!string.IsNullOrEmpty(linkSourceFolder) && Directory.Exists(linkSourceFolder))
            {
                try
                {
                    var searchName = typeName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase)
                        ? typeName : typeName + ".rvt";

                    LogHelper.Info($"Searching '{linkSourceFolder}' for '{searchName}'...");
                    var matches = Directory.GetFiles(linkSourceFolder, searchName, SearchOption.AllDirectories);
                    if (matches.Length > 0)
                    {
                        LogHelper.Info($"Resolved via linkSourceFolder search: {matches[0]}");
                        return matches[0];
                    }
                    LogHelper.Info($"No match found in linkSourceFolder for '{searchName}'");
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"linkSourceFolder search failed: {ex.Message}");
                }
            }

            return null;
        }

        public static string CopyLinkedModel(string sourceFilePath, string targetFileName, string outputFolder)
        {
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException($"Source linked model not found: {sourceFilePath}");

            if (!targetFileName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                targetFileName += ".rvt";

            var targetPath = Path.Combine(outputFolder, targetFileName);

            LogHelper.Info($"Copying: {Path.GetFileName(sourceFilePath)} -> {targetFileName}");
            File.Copy(sourceFilePath, targetPath, overwrite: true);

            return targetPath;
        }
    }
}
