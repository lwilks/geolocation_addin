using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using GeolocationAddin.Config;
using GeolocationAddin.Helpers;
using GeolocationAddin.Models;
using GeolocationAddin.UI;

namespace GeolocationAddin.Core
{
    public class GeolocationWorkflow
    {
        private readonly UIApplication _uiApp;
        private AddinConfig _config;

        public GeolocationWorkflow(UIApplication uiApp, AddinConfig config)
        {
            _uiApp = uiApp;
            _config = config;
        }

        public void Execute()
        {
            var results = new List<ProcessingResult>();

            // 1. Use the active document as the site model
            var uiDoc = _uiApp.ActiveUIDocument;
            if (uiDoc == null)
            {
                TaskDialog.Show("Geolocation", "No document is open.\n\nOpen the site model first, then run this command.");
                return;
            }

            var siteDoc = uiDoc.Document;
            LogHelper.Info($"Using active document as site model: {siteDoc.PathName}");

            // 2. Collect all RevitLinkInstance elements
            var linkInstances = new FilteredElementCollector(siteDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            LogHelper.Info($"Found {linkInstances.Count} link instances in site model.");

            foreach (var li in linkInstances)
            {
                try { LogHelper.Info($"  Link instance: \"{li.Name}\""); }
                catch { }
            }

            // 3. Build initial link list (empty target names)
            var matchList = BuildInitialLinkList(linkInstances);

            // 4. Auto-import from configured CSV path if available
            if (!string.IsNullOrWhiteSpace(_config.CsvMappingPath) && File.Exists(_config.CsvMappingPath))
            {
                try
                {
                    var imported = MappingSerializer.Import(_config.CsvMappingPath);
                    ApplyImportedMappings(matchList, imported);
                    LogHelper.Info($"Auto-imported {imported.Count} mapping(s) from: {_config.CsvMappingPath}");
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"Failed to auto-import CSV: {ex.Message}");
                }
            }

            // 5. Show link mapping dialog
            var mappingWindow = new LinkMappingWindow(matchList, _config);
            var revitHandle = _uiApp.MainWindowHandle;
            new WindowInteropHelper(mappingWindow).Owner = revitHandle;

            mappingWindow.ShowDialog();

            if (!mappingWindow.Confirmed)
            {
                LogHelper.Info("User cancelled link mapping.");
                return;
            }

            // Apply any settings the user changed in the mapping window
            _config = mappingWindow.Config;

            // 6. Build full info from selected links
            var linkInfos = BuildLinkInfoListFromSelection(siteDoc, matchList);
            LogHelper.Info($"Processing {linkInfos.Count} of {linkInstances.Count} selected link instances.");

            if (linkInfos.Count == 0)
            {
                LogHelper.Info("No links selected for processing. Continuing to show summary.");
            }

            // 7. Group by RevitLinkType for coordinate publishing
            var typeGroups = linkInfos.GroupBy(li => li.TypeId.Value);

            foreach (var group in typeGroups)
            {
                foreach (var linkInfo in group)
                {
                    var result = ProcessLink(siteDoc, linkInfo);
                    results.Add(result);
                }
            }

            // 8. Show summary
            ShowSummary(results);
        }

        private List<LinkMatchInfo> BuildInitialLinkList(List<RevitLinkInstance> instances)
        {
            var list = new List<LinkMatchInfo>();

            foreach (var instance in instances)
            {
                try
                {
                    list.Add(new LinkMatchInfo
                    {
                        InstanceName = instance.Name,
                        MatchType = MatchType.None,
                        IsSelected = false,
                        Instance = instance
                    });
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"Error reading link instance: {ex}");
                }
            }

            return list;
        }

        private void ApplyImportedMappings(List<LinkMatchInfo> links, List<(string linkName, string targetFileName, string label, string exportViewName)> imported)
        {
            var fuzzySettings = _config.FuzzyMatchSettings;

            foreach (var (linkName, targetFileName, label, exportViewName) in imported)
            {
                // Try exact match on InstanceName
                var exact = links.FirstOrDefault(l =>
                    string.Equals(l.InstanceName, linkName, StringComparison.OrdinalIgnoreCase));

                if (exact != null)
                {
                    exact.TargetFileName = targetFileName;
                    exact.Label = label;
                    exact.ExportViewName = exportViewName;
                    exact.MatchedImportKey = linkName;
                    exact.MatchType = MatchType.Exact;
                    exact.IsSelected = true;
                    LogHelper.Info($"Exact match: \"{linkName}\" -> \"{targetFileName}\"");
                    continue;
                }

                // Fuzzy fallback
                if (fuzzySettings.Enabled)
                {
                    var candidates = links
                        .Where(l => !l.HasTargetFileName)
                        .Select(l => l.InstanceName)
                        .ToList();

                    var fuzzyResult = FuzzyMatcher.FindBestMatch(
                        linkName, candidates,
                        fuzzySettings.TokenOverlapThreshold,
                        fuzzySettings.LevenshteinThreshold);

                    if (fuzzyResult != null)
                    {
                        var match = links.First(l =>
                            string.Equals(l.InstanceName, fuzzyResult.MatchedKey, StringComparison.OrdinalIgnoreCase));

                        match.TargetFileName = targetFileName;
                        match.Label = label;
                        match.ExportViewName = exportViewName;
                        match.MatchedImportKey = linkName;
                        match.MatchType = MatchType.Fuzzy;
                        LogHelper.Info($"Fuzzy match: \"{linkName}\" -> \"{fuzzyResult.MatchedKey}\" -> \"{targetFileName}\"");
                        continue;
                    }
                }

                LogHelper.Info($"No match for imported entry: \"{linkName}\"");
            }
        }

        private List<LinkInstanceInfo> BuildLinkInfoListFromSelection(Document siteDoc, List<LinkMatchInfo> matchList)
        {
            var linkInfos = new List<LinkInstanceInfo>();

            foreach (var match in matchList.Where(m => m.IsSelected && m.HasTargetFileName))
            {
                string instanceName = match.InstanceName;
                try
                {
                    var instance = match.Instance;
                    var typeId = instance.GetTypeId();
                    var linkType = siteDoc.GetElement(typeId) as RevitLinkType;
                    if (linkType == null)
                    {
                        LogHelper.Error($"Could not resolve RevitLinkType for: {instanceName}");
                        continue;
                    }

                    var sourcePath = FileCopyManager.ResolveLinkFilePath(linkType, _config.LinkSourceFolder);

                    var resolvedOutputFolder = PathHelper.ResolveLabel(_config.OutputFolder, match.Label);

                    var info = new LinkInstanceInfo
                    {
                        Instance = instance,
                        LinkType = linkType,
                        InstanceId = instance.Id,
                        TypeId = typeId,
                        InstanceName = instanceName,
                        SourceFilePath = sourcePath,
                        TargetFileName = match.TargetFileName,
                        TargetFilePath = Path.Combine(resolvedOutputFolder, match.TargetFileName),
                        Label = match.Label,
                        ExportViewName = match.ExportViewName,
                        TotalTransform = instance.GetTotalTransform()
                    };

                    // Extract cloud GUIDs from external resource references
                    try
                    {
                        var extResources = linkType.GetExternalResourceReferences();
                        if (extResources != null)
                        {
                            foreach (var kvp in extResources)
                            {
                                var resourceRef = kvp.Value;

                                var inSessionPath = resourceRef.InSessionPath;
                                if (!string.IsNullOrEmpty(inSessionPath))
                                {
                                    info.CloudInSessionPath = inSessionPath;
                                    LogHelper.Info($"Cloud path for '{instanceName}': {inSessionPath}");
                                }

                                var refMap = resourceRef.GetReferenceInformation();
                                if (refMap != null)
                                {
                                    foreach (var refEntry in refMap)
                                        LogHelper.Info($"  RefInfo['{instanceName}']: {refEntry.Key} = '{refEntry.Value}'");

                                    if (refMap.TryGetValue("LinkedModelProjectId", out var projectIdStr) &&
                                        refMap.TryGetValue("LinkedModelModelId", out var modelIdStr))
                                    {
                                        if (Guid.TryParse(projectIdStr, out var projectGuid) &&
                                            Guid.TryParse(modelIdStr, out var modelGuid))
                                        {
                                            info.CloudProjectGuid = projectGuid;
                                            info.CloudModelGuid = modelGuid;

                                            if (refMap.TryGetValue("LinkedModelRegion", out var region) &&
                                                !string.IsNullOrEmpty(region))
                                                info.CloudRegion = region;
                                            else
                                                info.CloudRegion = "US";

                                            LogHelper.Info($"Cloud GUIDs for '{instanceName}': project={projectGuid}, model={modelGuid}, region={info.CloudRegion}");
                                        }
                                        else
                                        {
                                            LogHelper.Error($"Could not parse cloud GUIDs for '{instanceName}': project='{projectIdStr}', model='{modelIdStr}'");
                                        }
                                    }
                                }

                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Info($"Could not extract cloud info for '{instanceName}': {ex.Message}");
                    }

                    if (!info.CloudProjectGuid.HasValue && sourcePath == null)
                    {
                        LogHelper.Error($"No cloud GUIDs and no source file path for: {instanceName}");
                        continue;
                    }

                    linkInfos.Add(info);
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"Error processing link '{instanceName ?? "unknown"}': {ex}");
                }
            }

            return linkInfos;
        }

        private ProcessingResult ProcessLink(Document siteDoc, LinkInstanceInfo linkInfo)
        {
            var result = new ProcessingResult
            {
                LinkName = linkInfo.InstanceName,
                TargetFileName = linkInfo.TargetFileName
            };

            LogHelper.Info($"\n--- Processing: {linkInfo.InstanceName} -> {linkInfo.TargetFileName} ---");

            var targetFileName = linkInfo.TargetFileName;
            if (!targetFileName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                targetFileName += ".rvt";

            var label = linkInfo.Label;
            var resolvedOutputFolder = PathHelper.ResolveLabel(_config.OutputFolder, label);
            var resolvedIfcFolder = PathHelper.ResolveLabel(_config.IfcOutputFolder, label);
            var resolvedNwcFolder = PathHelper.ResolveLabel(_config.NwcOutputFolder, label);
            var resolvedDwgFolder = PathHelper.ResolveLabel(_config.DwgOutputFolder, label);

            linkInfo.TargetFilePath = Path.Combine(resolvedOutputFolder, targetFileName);

            Directory.CreateDirectory(resolvedOutputFolder);
            if (_config.ExportSettings.ExportIfc && !string.IsNullOrEmpty(resolvedIfcFolder))
                Directory.CreateDirectory(resolvedIfcFolder);
            if (_config.ExportSettings.ExportNwc && !string.IsNullOrEmpty(resolvedNwcFolder))
                Directory.CreateDirectory(resolvedNwcFolder);
            if (_config.ExportSettings.ExportDwg && !string.IsNullOrEmpty(resolvedDwgFolder))
                Directory.CreateDirectory(resolvedDwgFolder);

            Document exportDoc = null;
            bool isCloudDoc = false;
            try
            {
                // === Phase 1: Create the local copy ===
                if (linkInfo.CloudProjectGuid.HasValue && linkInfo.CloudModelGuid.HasValue)
                {
                    LogHelper.Info("Opening via cloud GUIDs (detached)...");
                    exportDoc = RevitDocumentHelper.OpenCloudDocumentDetached(
                        _uiApp, linkInfo.CloudRegion,
                        linkInfo.CloudProjectGuid.Value, linkInfo.CloudModelGuid.Value);
                    isCloudDoc = true;
                    LogHelper.Info("Cloud model opened successfully (detached).");
                }
                else
                {
                    LogHelper.Info($"Opening local source (detached): {linkInfo.SourceFilePath}");
                    exportDoc = RevitDocumentHelper.OpenDocumentDetached(
                        _uiApp, linkInfo.SourceFilePath);
                    LogHelper.Info("Local source opened successfully.");
                }

                LogHelper.Info($"SaveAs to: {linkInfo.TargetFilePath}");
                RevitDocumentHelper.SaveDocumentAs(exportDoc, linkInfo.TargetFilePath);
                result.CopySucceeded = true;
                LogHelper.Info("SaveAs completed — clean local copy created.");

                // === Phase 2: Set shared coordinates on the still-open primary document ===
                // The copy is still open as a primary doc after SaveAs — transactions work directly.
                result.CoordinatesPublished = CoordinatePublisher.PublishViaTransform(
                    siteDoc, exportDoc, linkInfo.TotalTransform);

                if (!result.CoordinatesPublished)
                    LogHelper.Error("Coordinate publishing (Strategy B) failed.");

                // === Phase 3: Export ===
                var baseName = Path.GetFileNameWithoutExtension(targetFileName);

                if (_config.ExportSettings.ExportIfc)
                    result.IfcExported = ModelExporter.ExportIfc(exportDoc, resolvedIfcFolder, baseName, linkInfo.ExportViewName, _config.ExportSettings.IfcExportConfigPath);

                if (_config.ExportSettings.ExportNwc)
                    result.NwcExported = ModelExporter.ExportNwc(exportDoc, resolvedNwcFolder, baseName, linkInfo.ExportViewName);

                if (_config.ExportSettings.ExportDwg)
                    result.DwgExported = ModelExporter.ExportDwg(exportDoc, resolvedDwgFolder, baseName, linkInfo.ExportViewName);

                RevitDocumentHelper.SaveDocumentAs(exportDoc, linkInfo.TargetFilePath);
                LogHelper.Info("Final save completed.");

                // Cloud docs were opened via OpenAndActivateDocument — activate site doc first.
                // Local docs were opened via OpenDocumentDetached (background) and can close directly.
                if (isCloudDoc)
                    RevitDocumentHelper.ActivateDocument(_uiApp, siteDoc);

                exportDoc.Close(false);
                exportDoc = null;

                // === Phase 4: Publish shared coordinate relationship ===
                // Strategy B set the correct coordinate values. Now use Strategy A
                // (LoadFrom → PublishCoordinates → Reload) to establish the named shared
                // coordinate position that Revit requires for "By Shared Coordinates" linking.
                // The copy must be closed for LoadFrom to access the file on disk.
                LogHelper.Info("Attempting to publish shared coordinate relationship (Strategy A)...");
                bool published = CoordinatePublisher.PublishViaRelink(siteDoc, linkInfo);
                if (published)
                {
                    result.CoordinatesPublished = true;
                    LogHelper.Info("Shared coordinate relationship established.");
                }
                else
                {
                    LogHelper.Info("Strategy A failed — model has correct coordinates from Strategy B " +
                                   "but may not support 'By Shared Coordinates' linking.");
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Processing failed: {ex.Message}";
                LogHelper.Error(result.ErrorMessage);
                if (exportDoc != null)
                {
                    try
                    {
                        if (isCloudDoc)
                            RevitDocumentHelper.ActivateDocument(_uiApp, siteDoc);
                        exportDoc.Close(false);
                    }
                    catch { }
                }
            }

            return result;
        }

        private void ShowSummary(List<ProcessingResult> results)
        {
            var sb = new StringBuilder();

            if (results.Count == 0)
            {
                sb.AppendLine("No links were processed.");
                sb.AppendLine();
                sb.AppendLine("Check the log file for details on link discovery and path resolution:");
            }
            else
            {
                sb.AppendLine($"Processed {results.Count} link(s):\n");

                int succeeded = 0;
                int failed = 0;

                foreach (var r in results)
                {
                    sb.AppendLine($"  {r.TargetFileName}: {r.Summary}");

                    if (r.CopySucceeded && r.CoordinatesPublished)
                        succeeded++;
                    else
                        failed++;
                }

                sb.AppendLine();
                sb.AppendLine($"Succeeded: {succeeded}  |  Failed: {failed}");
                sb.AppendLine();
                sb.AppendLine("Log file:");
            }

            sb.AppendLine(LogHelper.LogFilePath);

            LogHelper.Info(sb.ToString());
            TaskDialog.Show("Geolocation — Results", sb.ToString());
        }
    }
}
