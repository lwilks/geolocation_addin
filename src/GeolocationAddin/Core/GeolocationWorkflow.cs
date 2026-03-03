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
        private readonly AddinConfig _config;
        private readonly CsvMapping _mapping;

        public GeolocationWorkflow(UIApplication uiApp, AddinConfig config, CsvMapping mapping)
        {
            _uiApp = uiApp;
            _config = config;
            _mapping = mapping;
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

            // Log all link instance names for debugging
            foreach (var li in linkInstances)
            {
                try { LogHelper.Info($"  Link instance: \"{li.Name}\""); }
                catch { }
            }

            // 3. Phase A — non-destructive matching
            var matchList = BuildMatchList(siteDoc, linkInstances);

            // 4. Show link selection dialog
            var selectionWindow = new LinkSelectionWindow(matchList);
            var revitHandle = _uiApp.MainWindowHandle;
            new WindowInteropHelper(selectionWindow).Owner = revitHandle;

            selectionWindow.ShowDialog();

            if (!selectionWindow.Confirmed)
            {
                LogHelper.Info("User cancelled link selection.");
                return;
            }

            // 5. Phase B — consume and build full info from selected links
            var linkInfos = BuildLinkInfoListFromSelection(siteDoc, matchList);
            LogHelper.Info($"Processing {linkInfos.Count} of {linkInstances.Count} selected link instances.");

            if (linkInfos.Count == 0)
            {
                LogHelper.Info("No links selected for processing. Continuing to show summary.");
            }

            // 6. Group by RevitLinkType for coordinate publishing
            var typeGroups = linkInfos.GroupBy(li => li.TypeId.Value);

            // Cloud models are opened with OpenAndActivateDocument (becomes active doc).
            // We can't close the active doc from the API, so we defer closing until
            // the next ProcessLink opens a new doc (switching the active doc away).
            Document pendingCloseDoc = null;

            foreach (var group in typeGroups)
            {
                foreach (var linkInfo in group)
                {
                    Document previousPending = pendingCloseDoc;
                    pendingCloseDoc = null;

                    var result = ProcessLink(siteDoc, linkInfo, out Document exportDoc);
                    results.Add(result);

                    // ProcessLink opened a new doc (now active), so the previous one
                    // is no longer active and can be closed.
                    if (previousPending != null)
                    {
                        try { previousPending.Close(false); }
                        catch (Exception ex) { LogHelper.Info($"Deferred close failed: {ex.Message}"); }
                    }

                    if (exportDoc != null)
                        pendingCloseDoc = exportDoc;
                }
            }

            // Close the last export doc by switching back to the site doc first
            if (pendingCloseDoc != null)
            {
                RevitDocumentHelper.ActivateDocument(_uiApp, siteDoc);
                try { pendingCloseDoc.Close(false); }
                catch (Exception ex) { LogHelper.Info($"Final close failed: {ex.Message}"); }
            }

            // 5. Show summary
            ShowSummary(results);
        }

        private List<LinkMatchInfo> BuildMatchList(Document siteDoc, List<RevitLinkInstance> instances)
        {
            var matchList = new List<LinkMatchInfo>();
            var fuzzySettings = _config.FuzzyMatchSettings;

            foreach (var instance in instances)
            {
                string instanceName = null;
                try
                {
                    instanceName = instance.Name;

                    // Try exact match (non-destructive peek)
                    if (_mapping.TryGetTargetName(instanceName, out var targetName))
                    {
                        matchList.Add(new LinkMatchInfo
                        {
                            InstanceName = instanceName,
                            MatchedCsvKey = instanceName,
                            TargetFileName = targetName,
                            MatchType = MatchType.Exact,
                            IsSelected = true,
                            Instance = instance
                        });
                        LogHelper.Info($"Exact match: \"{instanceName}\" -> \"{targetName}\"");
                        continue;
                    }

                    // Try fuzzy match
                    if (fuzzySettings.Enabled)
                    {
                        var candidates = _mapping.GetAvailableKeys().ToList();
                        var fuzzyResult = FuzzyMatcher.FindBestMatch(
                            instanceName, candidates,
                            fuzzySettings.TokenOverlapThreshold,
                            fuzzySettings.LevenshteinThreshold);

                        if (fuzzyResult != null)
                        {
                            _mapping.TryGetTargetName(fuzzyResult.MatchedKey, out var fuzzyTarget);
                            matchList.Add(new LinkMatchInfo
                            {
                                InstanceName = instanceName,
                                MatchedCsvKey = fuzzyResult.MatchedKey,
                                TargetFileName = fuzzyTarget,
                                MatchType = MatchType.Fuzzy,
                                TokenScore = fuzzyResult.TokenOverlapScore,
                                LevenshteinScore = fuzzyResult.LevenshteinScore,
                                IsSelected = false,
                                Instance = instance
                            });
                            LogHelper.Info($"Fuzzy match: \"{instanceName}\" -> \"{fuzzyResult.MatchedKey}\" " +
                                           $"(token: {fuzzyResult.TokenOverlapScore:F2}, lev: {fuzzyResult.LevenshteinScore:F2})");
                            continue;
                        }
                    }

                    // No match
                    matchList.Add(new LinkMatchInfo
                    {
                        InstanceName = instanceName,
                        MatchType = MatchType.None,
                        IsSelected = false,
                        Instance = instance
                    });
                    LogHelper.Info($"No match: \"{instanceName}\"");
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"Error matching link '{instanceName ?? "unknown"}': {ex}");
                }
            }

            return matchList;
        }

        private List<LinkInstanceInfo> BuildLinkInfoListFromSelection(Document siteDoc, List<LinkMatchInfo> matchList)
        {
            var linkInfos = new List<LinkInstanceInfo>();

            foreach (var match in matchList.Where(m => m.IsSelected && m.MatchType != MatchType.None))
            {
                string instanceName = match.InstanceName;
                try
                {
                    // Consume from mapping now that user has confirmed
                    var targetName = _mapping.ConsumeByKey(match.MatchedCsvKey);
                    if (targetName == null)
                    {
                        LogHelper.Error($"Could not consume mapping for: {match.MatchedCsvKey}");
                        continue;
                    }

                    var instance = match.Instance;
                    var typeId = instance.GetTypeId();
                    var linkType = siteDoc.GetElement(typeId) as RevitLinkType;
                    if (linkType == null)
                    {
                        LogHelper.Error($"Could not resolve RevitLinkType for: {instanceName}");
                        continue;
                    }

                    var sourcePath = FileCopyManager.ResolveLinkFilePath(linkType, _config.LinkSourceFolder);

                    var info = new LinkInstanceInfo
                    {
                        Instance = instance,
                        LinkType = linkType,
                        InstanceId = instance.Id,
                        TypeId = typeId,
                        InstanceName = instanceName,
                        SourceFilePath = sourcePath,
                        TargetFileName = targetName,
                        TargetFilePath = Path.Combine(_config.OutputFolder, targetName),
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

                                // Extract cloud GUIDs for ConvertCloudGUIDsToCloudPath
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

        private ProcessingResult ProcessLink(Document siteDoc, LinkInstanceInfo linkInfo, out Document openedDoc)
        {
            openedDoc = null;
            var result = new ProcessingResult
            {
                LinkName = linkInfo.InstanceName,
                TargetFileName = linkInfo.TargetFileName
            };

            LogHelper.Info($"\n--- Processing: {linkInfo.InstanceName} -> {linkInfo.TargetFileName} ---");

            var targetFileName = linkInfo.TargetFileName;
            if (!targetFileName.EndsWith(".rvt", StringComparison.OrdinalIgnoreCase))
                targetFileName += ".rvt";
            linkInfo.TargetFilePath = Path.Combine(_config.OutputFolder, targetFileName);

            Document exportDoc = null;
            bool isCloudDoc = false;
            try
            {
                // Step 1: Open the model (detached from central).
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
                    // Non-cloud link: open from local file path
                    LogHelper.Info($"Opening local source (detached): {linkInfo.SourceFilePath}");
                    exportDoc = RevitDocumentHelper.OpenDocumentDetached(
                        _uiApp, linkInfo.SourceFilePath);
                    LogHelper.Info("Local source opened successfully.");
                }

                // Step 2: SaveAs to target path — creates a clean standalone .rvt file
                LogHelper.Info($"SaveAs to: {linkInfo.TargetFilePath}");
                RevitDocumentHelper.SaveDocumentAs(exportDoc, linkInfo.TargetFilePath);
                result.CopySucceeded = true;
                LogHelper.Info("SaveAs completed — clean local copy created.");

                // Step 3: Apply coordinates via Strategy B (transform-based)
                // Pass siteDoc so we can add the site model's survey position offsets
                result.CoordinatesPublished = CoordinatePublisher.PublishViaTransform(
                    siteDoc, exportDoc, linkInfo.TotalTransform);

                if (!result.CoordinatesPublished)
                    LogHelper.Error("Failed to apply coordinates (Strategy B).");

                // Step 4: Export
                var baseName = Path.GetFileNameWithoutExtension(targetFileName);

                if (_config.ExportSettings.ExportIfc)
                    result.IfcExported = ModelExporter.ExportIfc(exportDoc, _config.IfcOutputFolder, baseName);

                if (_config.ExportSettings.ExportNwc)
                    result.NwcExported = ModelExporter.ExportNwc(exportDoc, _config.NwcOutputFolder, baseName);

                if (_config.ExportSettings.ExportDwg)
                    result.DwgExported = ModelExporter.ExportDwg(exportDoc, _config.DwgOutputFolder, baseName);

                // Step 5: Save the document
                RevitDocumentHelper.SaveDocumentAs(exportDoc, linkInfo.TargetFilePath);
                LogHelper.Info("Final save completed.");

                // Cloud docs opened via OpenAndActivateDocument become the active doc
                // and can't be closed from the API. Defer close to the caller.
                if (isCloudDoc)
                {
                    openedDoc = exportDoc;
                }
                else
                {
                    exportDoc.Close(false);
                    exportDoc = null;
                }
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Processing failed: {ex.Message}";
                LogHelper.Error(result.ErrorMessage);
                if (exportDoc != null)
                {
                    // For cloud docs that became active, defer close
                    if (isCloudDoc)
                        openedDoc = exportDoc;
                    else
                        try { exportDoc.Close(false); } catch { }
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
