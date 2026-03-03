namespace GeolocationAddin.Models
{
    public class ProcessingResult
    {
        public string LinkName { get; set; }
        public string TargetFileName { get; set; }
        public bool CopySucceeded { get; set; }
        public bool CoordinatesPublished { get; set; }
        public bool IfcExported { get; set; }
        public bool NwcExported { get; set; }
        public bool DwgExported { get; set; }
        public string ErrorMessage { get; set; }

        public bool FullySucceeded =>
            CopySucceeded && CoordinatesPublished &&
            (!IfcExported || IfcExported) &&
            (!NwcExported || NwcExported) &&
            (!DwgExported || DwgExported) &&
            string.IsNullOrEmpty(ErrorMessage);

        public string Summary
        {
            get
            {
                if (!CopySucceeded)
                    return $"FAILED (copy): {ErrorMessage}";
                if (!CoordinatesPublished)
                    return $"FAILED (coordinates): {ErrorMessage}";

                var exports = "";
                if (IfcExported) exports += " IFC";
                if (NwcExported) exports += " NWC";
                if (DwgExported) exports += " DWG";

                var exportText = string.IsNullOrEmpty(exports) ? "no exports" : $"exported:{exports}";
                return $"OK — {exportText}";
            }
        }
    }
}
