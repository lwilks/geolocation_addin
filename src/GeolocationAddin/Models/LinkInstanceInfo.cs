using Autodesk.Revit.DB;

namespace GeolocationAddin.Models
{
    public class LinkInstanceInfo
    {
        public RevitLinkInstance Instance { get; set; }
        public RevitLinkType LinkType { get; set; }
        public ElementId InstanceId { get; set; }
        public ElementId TypeId { get; set; }
        public string InstanceName { get; set; }
        public string SourceFilePath { get; set; }
        public string TargetFileName { get; set; }
        public string TargetFilePath { get; set; }
        public Transform TotalTransform { get; set; }
    }
}
