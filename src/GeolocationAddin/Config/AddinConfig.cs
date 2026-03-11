using Newtonsoft.Json;

namespace GeolocationAddin.Config
{
    public class AddinConfig
    {
        [JsonProperty("siteModelPath")]
        public string SiteModelPath { get; set; }

        [JsonProperty("csvMappingPath")]
        public string CsvMappingPath { get; set; }

        [JsonProperty("linkSourceFolder")]
        public string LinkSourceFolder { get; set; }

        [JsonProperty("outputFolder")]
        public string OutputFolder { get; set; }

        [JsonProperty("ifcOutputFolder")]
        public string IfcOutputFolder { get; set; }

        [JsonProperty("nwcOutputFolder")]
        public string NwcOutputFolder { get; set; }

        [JsonProperty("dwgOutputFolder")]
        public string DwgOutputFolder { get; set; }

        [JsonProperty("exportSettings")]
        public ExportSettings ExportSettings { get; set; } = new ExportSettings();

        [JsonProperty("fuzzyMatchSettings")]
        public FuzzyMatchSettings FuzzyMatchSettings { get; set; } = new FuzzyMatchSettings();
    }

    public class ExportSettings
    {
        [JsonProperty("exportIfc")]
        public bool ExportIfc { get; set; } = true;

        [JsonProperty("exportNwc")]
        public bool ExportNwc { get; set; } = true;

        [JsonProperty("exportDwg")]
        public bool ExportDwg { get; set; } = true;

        [JsonProperty("ifcExportConfigPath")]
        public string IfcExportConfigPath { get; set; }
    }

    public class FuzzyMatchSettings
    {
        [JsonProperty("enabled")]
        public bool Enabled { get; set; } = true;

        [JsonProperty("tokenOverlapThreshold")]
        public double TokenOverlapThreshold { get; set; } = 0.5;

        [JsonProperty("levenshteinThreshold")]
        public double LevenshteinThreshold { get; set; } = 0.4;
    }
}
