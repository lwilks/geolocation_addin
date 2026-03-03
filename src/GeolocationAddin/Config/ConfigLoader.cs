using System;
using System.IO;
using Newtonsoft.Json;

namespace GeolocationAddin.Config
{
    public static class ConfigLoader
    {
        private const string ConfigDir = @"C:\ProgramData\GeolocationAddin";
        private const string ConfigFileName = "config.json";

        public static string ConfigPath => Path.Combine(ConfigDir, ConfigFileName);

        public static AddinConfig Load()
        {
            var path = ConfigPath;

            if (!File.Exists(path))
                throw new ConfigurationException(
                    $"Config file not found at:\n{path}\n\n" +
                    "Run the deploy script or create the file manually.");

            var json = File.ReadAllText(path);
            var config = JsonConvert.DeserializeObject<AddinConfig>(json);

            if (config == null)
                throw new ConfigurationException("Config file is empty or invalid JSON.");

            Validate(config);
            return config;
        }

        private static void Validate(AddinConfig config)
        {
            if (string.IsNullOrWhiteSpace(config.CsvMappingPath))
                throw new ConfigurationException("csvMappingPath is required in config.");

            if (string.IsNullOrWhiteSpace(config.OutputFolder))
                throw new ConfigurationException("outputFolder is required in config.");

            EnsureDirectory(config.OutputFolder, "outputFolder");

            if (config.ExportSettings.ExportIfc)
            {
                if (string.IsNullOrWhiteSpace(config.IfcOutputFolder))
                    throw new ConfigurationException("ifcOutputFolder is required when IFC export is enabled.");
                EnsureDirectory(config.IfcOutputFolder, "ifcOutputFolder");
            }

            if (config.ExportSettings.ExportNwc)
            {
                if (string.IsNullOrWhiteSpace(config.NwcOutputFolder))
                    throw new ConfigurationException("nwcOutputFolder is required when NWC export is enabled.");
                EnsureDirectory(config.NwcOutputFolder, "nwcOutputFolder");
            }

            if (config.ExportSettings.ExportDwg)
            {
                if (string.IsNullOrWhiteSpace(config.DwgOutputFolder))
                    throw new ConfigurationException("dwgOutputFolder is required when DWG export is enabled.");
                EnsureDirectory(config.DwgOutputFolder, "dwgOutputFolder");
            }
        }

        public static void Save(AddinConfig config)
        {
            Validate(config);

            var dir = Path.GetDirectoryName(ConfigPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var json = JsonConvert.SerializeObject(config, Formatting.Indented);
            File.WriteAllText(ConfigPath, json);
        }

        private static void EnsureDirectory(string path, string fieldName)
        {
            try
            {
                Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                throw new ConfigurationException($"Cannot create directory for {fieldName} ({path}): {ex.Message}");
            }
        }
    }

    public class ConfigurationException : Exception
    {
        public ConfigurationException(string message) : base(message) { }
    }
}
