using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace GeolocationAddin.Helpers
{
    public static class MappingSerializer
    {
        public static List<(string linkName, string targetFileName, string label, string exportViewName)> Import(string path)
        {
            var ext = Path.GetExtension(path);
            if (string.Equals(ext, ".xml", StringComparison.OrdinalIgnoreCase))
                return ImportXml(path);

            return ImportCsv(path);
        }

        public static void Export(string path, IEnumerable<(string linkName, string targetFileName, string label, string exportViewName)> mappings)
        {
            var ext = Path.GetExtension(path);
            if (string.Equals(ext, ".xml", StringComparison.OrdinalIgnoreCase))
                ExportXml(path, mappings);
            else
                ExportCsv(path, mappings);
        }

        public static List<(string linkName, string targetFileName, string label, string exportViewName)> ImportCsv(string path)
        {
            var results = new List<(string, string, string, string)>();
            var lines = File.ReadAllLines(path);

            // Skip header row
            for (int i = 1; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line))
                    continue;

                var parts = line.Split(',');
                if (parts.Length < 2)
                    continue;

                var linkName = parts[0].Trim();
                var targetName = parts[1].Trim();
                var label = parts.Length >= 3 ? parts[2].Trim() : "";
                var exportViewName = parts.Length >= 4 ? parts[3].Trim() : "";

                if (string.IsNullOrEmpty(linkName) || string.IsNullOrEmpty(targetName))
                    continue;

                results.Add((linkName, targetName, label, exportViewName));
            }

            return results;
        }

        public static List<(string linkName, string targetFileName, string label, string exportViewName)> ImportXml(string path)
        {
            var results = new List<(string, string, string, string)>();
            var doc = XDocument.Load(path);

            var root = doc.Root;
            if (root == null)
                return results;

            foreach (var el in root.Elements("Mapping"))
            {
                var linkName = (string)el.Attribute("LinkName");
                var targetName = (string)el.Attribute("TargetFileName");
                var label = (string)el.Attribute("Label") ?? "";
                var exportViewName = (string)el.Attribute("ExportViewName") ?? "";

                if (string.IsNullOrEmpty(linkName) || string.IsNullOrEmpty(targetName))
                    continue;

                results.Add((linkName, targetName, label, exportViewName));
            }

            return results;
        }

        public static void ExportCsv(string path, IEnumerable<(string linkName, string targetFileName, string label, string exportViewName)> mappings)
        {
            var sb = new StringBuilder();
            sb.AppendLine("LinkName,TargetFileName,Label,ExportViewName");

            foreach (var (linkName, targetFileName, label, exportViewName) in mappings)
                sb.AppendLine($"{linkName},{targetFileName},{label},{exportViewName}");

            File.WriteAllText(path, sb.ToString());
        }

        public static void ExportXml(string path, IEnumerable<(string linkName, string targetFileName, string label, string exportViewName)> mappings)
        {
            var doc = new XDocument(
                new XElement("LinkMappings",
                    mappings.Select(m =>
                        new XElement("Mapping",
                            new XAttribute("LinkName", m.linkName),
                            new XAttribute("TargetFileName", m.targetFileName),
                            new XAttribute("Label", m.label),
                            new XAttribute("ExportViewName", m.exportViewName)))));

            doc.Save(path);
        }
    }
}
