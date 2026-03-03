using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using GeolocationAddin.Helpers;

namespace GeolocationAddin.Core
{
    public static class ModelExporter
    {
        public static bool ExportIfc(Document doc, string outputFolder, string fileNameWithoutExtension)
        {
            try
            {
                using (var tx = new Transaction(doc, "Export IFC"))
                {
                    tx.Start();

                    var options = new IFCExportOptions();

                    doc.Export(outputFolder, fileNameWithoutExtension, options);

                    tx.Commit();
                }

                var expectedPath = Path.Combine(outputFolder, fileNameWithoutExtension + ".ifc");
                if (File.Exists(expectedPath))
                {
                    LogHelper.Info($"IFC exported: {expectedPath}");
                    return true;
                }

                LogHelper.Error("IFC export completed but file not found.");
                return false;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"IFC export failed: {ex.Message}");
                return false;
            }
        }

        public static bool ExportNwc(Document doc, string outputFolder, string fileNameWithoutExtension)
        {
            try
            {
                if (!OptionalFunctionalityUtils.IsNavisworksExporterAvailable())
                {
                    LogHelper.Error("NWC export unavailable — Navisworks NWC Export Utility is not installed.");
                    return false;
                }

                var options = new NavisworksExportOptions
                {
                    ExportScope = NavisworksExportScope.Model
                };

                bool exportResult = doc.Export(outputFolder, fileNameWithoutExtension, options);
                LogHelper.Info($"NWC Export() returned: {exportResult}");

                // Search for the NWC file (exporter may vary naming)
                var expectedPath = Path.Combine(outputFolder, fileNameWithoutExtension + ".nwc");
                if (File.Exists(expectedPath))
                {
                    LogHelper.Info($"NWC exported: {expectedPath}");
                    return true;
                }

                // Search with wildcard in case of naming variation
                var files = Directory.GetFiles(outputFolder, fileNameWithoutExtension + "*.nwc");
                if (files.Length > 0)
                {
                    LogHelper.Info($"NWC exported: {files[0]}");
                    return true;
                }

                // Log all NWC files in output folder for diagnostics
                var allNwc = Directory.GetFiles(outputFolder, "*.nwc");
                LogHelper.Error($"NWC export completed but file not found. NWC files in folder: {allNwc.Length}");
                foreach (var f in allNwc)
                    LogHelper.Info($"  NWC file: {Path.GetFileName(f)}");

                return false;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"NWC export failed: {ex.Message}");
                return false;
            }
        }

        public static bool ExportDwg(Document doc, string outputFolder, string fileNameWithoutExtension)
        {
            try
            {
                // Find the default 3D view to export
                var view3d = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);

                if (view3d == null)
                {
                    LogHelper.Error("No 3D view found for DWG export.");
                    return false;
                }

                using (var tx = new Transaction(doc, "Export DWG"))
                {
                    tx.Start();

                    var options = new DWGExportOptions();
                    var viewIds = new System.Collections.Generic.List<ElementId> { view3d.Id };

                    doc.Export(outputFolder, fileNameWithoutExtension, viewIds, options);

                    tx.Commit();
                }

                // DWG export may produce a file with the view name appended
                var files = Directory.GetFiles(outputFolder, fileNameWithoutExtension + "*.dwg");
                if (files.Length > 0)
                {
                    LogHelper.Info($"DWG exported: {files[0]}");
                    return true;
                }

                LogHelper.Error("DWG export completed but file not found.");
                return false;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"DWG export failed: {ex.Message}");
                return false;
            }
        }
    }
}
