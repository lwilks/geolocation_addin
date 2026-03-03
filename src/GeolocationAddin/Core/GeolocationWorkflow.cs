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

            // 1. Open site model
            LogHelper.Info($"Opening site model: {_config.SiteModelPath}");
            var siteDoc = RevitDocumentHelper.OpenDocument(_uiApp.Application, _config.SiteModelPath);

            try
            {
                // 2. Collect all RevitLinkInstance elements
                var linkInstances = new FilteredElementCollector(siteDoc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                LogHelper.Info($"Found {linkInstances.Count} link instances in site model.");

                // 3. Build link info list
                var linkInfos = BuildLinkInfoList(siteDoc, linkInstances);
                LogHelper.Info($"Matched {linkInfos.Count} link instances to CSV mapping entries.");

                if (linkInfos.Count == 0)
                {
                    TaskDialog.Show("Geolocation",
                        "No link instances matched the CSV mapping.\n\n" +
                        "Check that LinkInstanceName values in the CSV match the link instance names in Revit.");
                    return;
                }

                // 4. Group by RevitLinkType for coordinate publishing
                var typeGroups = linkInfos.GroupBy(li => li.TypeId.IntegerValue);

                foreach (var group in typeGroups)
                {
                    foreach (var linkInfo in group)
                    {
                        var result = ProcessLink(siteDoc, linkInfo);
                        results.Add(result);
                    }
                }
            }
            finally
            {
                // Close site model without saving (we restored all links)
                RevitDocumentHelper.CloseDocument(siteDoc, save: false);
            }

            // 5. Show summary
            ShowSummary(results);
        }

        private List<LinkInstanceInfo> BuildLinkInfoList(Document siteDoc, List<RevitLinkInstance> instances)
        {
            var linkInfos = new List<LinkInstanceInfo>();

            foreach (var instance in instances)
            {
                var instanceName = instance.Name;
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

                var sourcePath = FileCopyManager.ResolveLinkFilePath(linkType);
                if (sourcePath == null)
                {
                    LogHelper.Error($"Could not resolve source file path for: {instanceName}");
                    continue;
                }

                linkInfos.Add(new LinkInstanceInfo
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
                });
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

            // Step 1: Copy source file
            try
            {
                linkInfo.TargetFilePath = FileCopyManager.CopyLinkedModel(
                    linkInfo.SourceFilePath, linkInfo.TargetFileName, _config.OutputFolder);
                result.CopySucceeded = true;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Copy failed: {ex.Message}";
                LogHelper.Error(result.ErrorMessage);
                return result;
            }

            // Step 2: Publish shared coordinates (Strategy A, fallback B)
            bool coordsPublished = CoordinatePublisher.PublishViaRelink(siteDoc, linkInfo);

            if (!coordsPublished)
            {
                LogHelper.Info("Strategy A failed, attempting Strategy B (transform-based)...");

                // For Strategy B we need to open the copied model
                Document copiedDoc = null;
                try
                {
                    copiedDoc = RevitDocumentHelper.OpenDocumentDetached(
                        _uiApp.Application, linkInfo.TargetFilePath);
                    coordsPublished = CoordinatePublisher.PublishViaTransform(copiedDoc, linkInfo.TotalTransform);
                    RevitDocumentHelper.CloseDocument(copiedDoc, save: true);
                    copiedDoc = null;
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"Strategy B open/apply failed: {ex.Message}");
                    if (copiedDoc != null)
                        RevitDocumentHelper.CloseDocument(copiedDoc, save: false);
                }
            }

            result.CoordinatesPublished = coordsPublished;

            if (!coordsPublished)
            {
                result.ErrorMessage = "Both coordinate publishing strategies failed.";
                LogHelper.Error(result.ErrorMessage);
                return result;
            }

            // Step 3: Export
            Document exportDoc = null;
            try
            {
                exportDoc = RevitDocumentHelper.OpenDocumentDetached(
                    _uiApp.Application, linkInfo.TargetFilePath);

                var baseName = Path.GetFileNameWithoutExtension(linkInfo.TargetFileName);

                if (_config.ExportSettings.ExportIfc)
                    result.IfcExported = ModelExporter.ExportIfc(exportDoc, _config.IfcOutputFolder, baseName);

                if (_config.ExportSettings.ExportNwc)
                    result.NwcExported = ModelExporter.ExportNwc(exportDoc, _config.NwcOutputFolder, baseName);

                if (_config.ExportSettings.ExportDwg)
                    result.DwgExported = ModelExporter.ExportDwg(exportDoc, _config.DwgOutputFolder, baseName);

                RevitDocumentHelper.CloseDocument(exportDoc, save: false);
                exportDoc = null;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Export failed: {ex.Message}";
                LogHelper.Error(result.ErrorMessage);
                if (exportDoc != null)
                    RevitDocumentHelper.CloseDocument(exportDoc, save: false);
            }

            return result;
        }

        private void ShowSummary(List<ProcessingResult> results)
        {
            var sb = new StringBuilder();
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

            LogHelper.Info(sb.ToString());
            TaskDialog.Show("Geolocation — Results", sb.ToString());
        }
    }
}
