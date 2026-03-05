using System.IO;

namespace GeolocationAddin.Helpers
{
    public static class PathHelper
    {
        public static string GetFileNameWithoutExtension(string filePath)
        {
            return Path.GetFileNameWithoutExtension(filePath);
        }

        public static string EnsureRvtExtension(string fileName)
        {
            if (!fileName.EndsWith(".rvt", System.StringComparison.OrdinalIgnoreCase))
                return fileName + ".rvt";
            return fileName;
        }

        public static string CombineAndNormalize(string folder, string fileName)
        {
            return Path.GetFullPath(Path.Combine(folder, fileName));
        }

        public static string ResolveLabel(string path, string label)
        {
            if (string.IsNullOrEmpty(path) || !path.Contains("<LABEL>"))
                return path;

            var resolved = path.Replace("<LABEL>", label ?? "");
            resolved = resolved.Replace("\\\\", "\\");
            return resolved;
        }
    }
}
