using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using GeolocationAddin.Helpers;

namespace GeolocationAddin.Core
{
    public static class FileCopyManager
    {
        public static string ResolveLinkFilePath(RevitLinkType linkType)
        {
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
                            return path;
                    }
                }
            }
            catch (InvalidOperationException)
            {
                // "This Element does not represent an external file" — cloud/ACC link
                LogHelper.Info($"Link type '{linkType.Name}' is not a traditional file link, trying external resource path...");
            }

            // Strategy 2: External resource references (ACC/BIM360 via Desktop Connector)
            try
            {
                var extResources = linkType.GetExternalResourceReferences();
                if (extResources != null)
                {
                    foreach (var kvp in extResources)
                    {
                        var resourceRef = kvp.Value;
                        var versionPath = resourceRef.InSessionPath;
                        if (!string.IsNullOrEmpty(versionPath) && File.Exists(versionPath))
                        {
                            LogHelper.Info($"Resolved via external resource InSessionPath: {versionPath}");
                            return versionPath;
                        }

                        // Try the server path which Desktop Connector maps to a local path
                        var refMap = resourceRef.GetReferenceInformation();
                        if (refMap != null)
                        {
                            foreach (var info in refMap)
                            {
                                if (File.Exists(info.Value))
                                {
                                    LogHelper.Info($"Resolved via resource reference info: {info.Value}");
                                    return info.Value;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"External resource resolution failed: {ex.Message}");
            }

            // Strategy 3: Try to find the file via the link type name in known Desktop Connector paths
            try
            {
                var typeName = linkType.Name;
                // Strip ".rvt" suffix if present for searching
                var searchName = typeName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase)
                    ? typeName : typeName + ".rvt";

                var dcRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "DC");

                if (Directory.Exists(dcRoot))
                {
                    var matches = Directory.GetFiles(dcRoot, searchName, SearchOption.AllDirectories);
                    if (matches.Length > 0)
                    {
                        LogHelper.Info($"Resolved via Desktop Connector search: {matches[0]}");
                        return matches[0];
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Desktop Connector path search failed: {ex.Message}");
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
