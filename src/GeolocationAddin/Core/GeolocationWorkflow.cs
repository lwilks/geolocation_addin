using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using GeolocationAddin.Config;
using GeolocationAddin.Helpers;
using GeolocationAddin.Models;

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

            // 3. Build link info list
            var linkInfos = BuildLinkInfoList(siteDoc, linkInstances);
            LogHelper.Info($"Matched {linkInfos.Count} of {linkInstances.Count} link instances to CSV mapping entries.");

            if (linkInfos.Count == 0)
            {
                LogHelper.Info("No link instances matched. Continuing to show summary.");
            }

            // 4. Group by RevitLinkType for coordinate publishing
            var typeGroups = linkInfos.GroupBy(li => li.TypeId.Value);

            foreach (var group in typeGroups)
            {
                foreach (var linkInfo in group)
                {
                    var result = ProcessLink(siteDoc, linkInfo);
                    results.Add(result);
                }
            }

            // 5. Show summary
            ShowSummary(results);
        }

        private List<LinkInstanceInfo> BuildLinkInfoList(Document siteDoc, List<RevitLinkInstance> instances)
        {
            var linkInfos = new List<LinkInstanceInfo>();

            foreach (var instance in instances)
            {
                string instanceName = null;
                try
                {
                    instanceName = instance.Name;
                    var targetName = _mapping.ConsumeTargetName(instanceName);

                    if (targetName == null)
                    {
                        LogHelper.Info($"Skipping unmapped link: {instanceName}");
                        continue;
                    }

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
                                var refMap = kvp.Value.GetReferenceInformation();
                                if (refMap != null)
                                {
                                    string projId, modId;
                                    if (refMap.TryGetValue("LinkedModelProjectId", out projId))
                                        info.CloudProjectGuid = Guid.Parse(projId);
                                    if (refMap.TryGetValue("LinkedModelModelId", out modId))
                                        info.CloudModelGuid = Guid.Parse(modId);
                                }
                                break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Info($"Could not extract cloud GUIDs for '{instanceName}': {ex.Message}");
                    }

                    if (info.CloudProjectGuid.HasValue && info.CloudModelGuid.HasValue)
                    {
                        LogHelper.Info($"Cloud GUIDs for '{instanceName}': project={info.CloudProjectGuid}, model={info.CloudModelGuid}");
                    }
                    else if (sourcePath == null)
                    {
                        LogHelper.Error($"No cloud GUIDs and no source path for: {instanceName}");
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
            linkInfo.TargetFilePath = Path.Combine(_config.OutputFolder, targetFileName);

            Document exportDoc = null;
            try
            {
                // Step 1: Open the cloud model via its cloud path (detached from central).
                // ACC Desktop Connector files are in a cloud-specific format that cannot be
                // opened via local filesystem paths. We must use ConvertCloudPath with the
                // model's cloud GUIDs to open through Revit's cloud mechanism.
                if (linkInfo.CloudProjectGuid.HasValue && linkInfo.CloudModelGuid.HasValue)
                {
                    LogHelper.Info("Opening via cloud path (detached)...");
                    exportDoc = RevitDocumentHelper.OpenCloudDocumentDetached(
                        _uiApp,
                        linkInfo.CloudProjectGuid.Value,
                        linkInfo.CloudModelGuid.Value);
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
                result.CoordinatesPublished = CoordinatePublisher.PublishViaTransform(
                    exportDoc, linkInfo.TotalTransform);

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

                // Step 5: Save and close
                RevitDocumentHelper.CloseDocument(exportDoc, save: true);
                exportDoc = null;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Processing failed: {ex.Message}";
                LogHelper.Error(result.ErrorMessage);
                if (exportDoc != null)
                {
                    try { RevitDocumentHelper.CloseDocument(exportDoc, save: false); }
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
