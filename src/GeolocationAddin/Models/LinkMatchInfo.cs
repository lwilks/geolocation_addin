using System.ComponentModel;
using Autodesk.Revit.DB;

namespace GeolocationAddin.Models
{
    public enum MatchType
    {
        Exact,
        Fuzzy,
        None
    }

    public class LinkMatchInfo : INotifyPropertyChanged
    {
        private bool _isSelected;

        public string InstanceName { get; set; }
        public string MatchedCsvKey { get; set; }
        public string TargetFileName { get; set; }
        public MatchType MatchType { get; set; }
        public double TokenScore { get; set; }
        public double LevenshteinScore { get; set; }
        public RevitLinkInstance Instance { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }
        }

        public string MatchDescription
        {
            get
            {
                switch (MatchType)
                {
                    case MatchType.Exact:
                        return "Exact match";
                    case MatchType.Fuzzy:
                        return $"Fuzzy (token: {TokenScore:P0}, lev: {LevenshteinScore:P0})";
                    default:
                        return "No match";
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
