using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Newtonsoft.Json.Linq;

namespace GeolocationAddin.Helpers
{
    public static class IfcConfigLoader
    {
        // Metadata properties that should not be passed through AddOption
        private static readonly HashSet<string> SkipProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Name", "VisibleElementsOfCurrentView"
        };

        public static void ApplyConfigFile(IFCExportOptions options, string configPath)
        {
            if (string.IsNullOrWhiteSpace(configPath))
                return;

            if (!File.Exists(configPath))
            {
                LogHelper.Info($"IFC export config file not found, using defaults: {configPath}");
                return;
            }

            try
            {
                var json = File.ReadAllText(configPath);
                var config = JObject.Parse(json);

                LogHelper.Info($"Applying IFC export config from: {configPath}");
                ApplyConfig(options, config);
            }
            catch (Exception ex)
            {
                LogHelper.Error($"Failed to apply IFC export config, using defaults: {ex.Message}");
            }
        }

        private static void ApplyConfig(IFCExportOptions options, JObject config)
        {
            foreach (var property in config.Properties())
            {
                var name = property.Name;

                if (SkipProperties.Contains(name))
                    continue;

                try
                {
                    if (string.Equals(name, "IFCVersion", StringComparison.OrdinalIgnoreCase))
                    {
                        var version = ParseIfcVersion(property.Value.ToString());
                        if (version.HasValue)
                            options.FileVersion = version.Value;
                        else
                            LogHelper.Info($"Unknown IFC version '{property.Value}', skipping.");
                    }
                    else if (string.Equals(name, "SpaceBoundaries", StringComparison.OrdinalIgnoreCase))
                    {
                        if (int.TryParse(property.Value.ToString(), out var level))
                            options.SpaceBoundaryLevel = level;
                    }
                    else if (string.Equals(name, "ExportBaseQuantities", StringComparison.OrdinalIgnoreCase))
                    {
                        options.ExportBaseQuantities = GetBool(property.Value);
                    }
                    else if (string.Equals(name, "SplitWallsAndColumns", StringComparison.OrdinalIgnoreCase))
                    {
                        options.WallAndColumnSplitting = GetBool(property.Value);
                    }
                    else
                    {
                        // Pass through as string via AddOption
                        var value = FormatValue(property.Value);
                        options.AddOption(name, value);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.Info($"Skipping IFC config property '{name}': {ex.Message}");
                }
            }
        }

        private static IFCVersion? ParseIfcVersion(string value)
        {
            // Try enum parse first (e.g. "IFC4")
            if (Enum.TryParse<IFCVersion>(value, true, out var version))
                return version;

            // Common string mappings from Revit's IFC exporter config
            switch (value.ToUpperInvariant())
            {
                case "IFC2X3":
                case "IFC 2X3":
                    return IFCVersion.IFC2x3;
                case "IFC2X3CV2.0":
                case "IFC2X3 COORDINATION VIEW 2.0":
                    return IFCVersion.IFC2x3CV2;
                case "IFC4":
                case "IFC 4":
                    return IFCVersion.IFC4;
                case "IFC4RV":
                case "IFC4 REFERENCE VIEW":
                    return IFCVersion.IFC4RV;
                case "IFC4DTV":
                case "IFC4 DESIGN TRANSFER VIEW":
                    return IFCVersion.IFC4DTV;
                default:
                    return null;
            }
        }

        private static bool GetBool(JToken token)
        {
            if (token.Type == JTokenType.Boolean)
                return token.Value<bool>();

            return string.Equals(token.ToString(), "true", StringComparison.OrdinalIgnoreCase);
        }

        private static string FormatValue(JToken token)
        {
            if (token.Type == JTokenType.Boolean)
                return token.Value<bool>() ? "true" : "false";

            return token.ToString();
        }
    }
}
