using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GeolocationAddin.Config
{
    public class CsvMapping
    {
        private readonly Dictionary<string, List<string>> _mapping;

        private CsvMapping(Dictionary<string, List<string>> mapping)
        {
            _mapping = mapping;
        }

        public int EntryCount => _mapping.Values.Sum(v => v.Count);

        public bool TryGetTargetName(string linkInstanceName, out string targetFileName)
        {
            if (_mapping.TryGetValue(linkInstanceName, out var targets) && targets.Count > 0)
            {
                targetFileName = targets[0];
                return true;
            }
            targetFileName = null;
            return false;
        }

        public string ConsumeTargetName(string linkInstanceName)
        {
            if (!_mapping.TryGetValue(linkInstanceName, out var targets) || targets.Count == 0)
                return null;

            var name = targets[0];
            targets.RemoveAt(0);
            return name;
        }

        public static CsvMapping Load(string csvPath)
        {
            if (!File.Exists(csvPath))
                throw new ConfigurationException($"CSV mapping file not found: {csvPath}");

            var mapping = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            var lines = File.ReadAllLines(csvPath);

            if (lines.Length == 0)
                throw new ConfigurationException("CSV mapping file is empty.");

            // Skip header row
            for (int i = 1; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line))
                    continue;

                var parts = line.Split(',');
                if (parts.Length < 2)
                    throw new ConfigurationException($"CSV line {i + 1} is malformed (expected 2 columns): {line}");

                var linkName = parts[0].Trim();
                var targetName = parts[1].Trim();

                if (string.IsNullOrEmpty(linkName) || string.IsNullOrEmpty(targetName))
                    throw new ConfigurationException($"CSV line {i + 1} has empty values: {line}");

                if (!mapping.ContainsKey(linkName))
                    mapping[linkName] = new List<string>();

                mapping[linkName].Add(targetName);
            }

            if (mapping.Count == 0)
                throw new ConfigurationException("CSV mapping file contains no data rows.");

            return new CsvMapping(mapping);
        }
    }
}
