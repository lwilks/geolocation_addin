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
        private string _targetFileName;
        private string _validationError;

        public string InstanceName { get; set; }
        public string MatchedImportKey { get; set; }
        public MatchType MatchType { get; set; }
        public RevitLinkInstance Instance { get; set; }

        public string TargetFileName
        {
            get => _targetFileName;
            set
            {
                if (_targetFileName != value)
                {
                    _targetFileName = value;
                    OnPropertyChanged(nameof(TargetFileName));
                    OnPropertyChanged(nameof(HasTargetFileName));

                    // Deselect if target name was cleared
                    if (!HasTargetFileName && _isSelected)
                    {
                        _isSelected = false;
                        OnPropertyChanged(nameof(IsSelected));
                    }
                }
            }
        }

        public bool HasTargetFileName => !string.IsNullOrWhiteSpace(_targetFileName);

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                // Refuse selection when there's no target file name
                bool newValue = value && HasTargetFileName;
                if (_isSelected != newValue)
                {
                    _isSelected = newValue;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public string ValidationError
        {
            get => _validationError;
            set
            {
                if (_validationError != value)
                {
                    _validationError = value;
                    OnPropertyChanged(nameof(ValidationError));
                    OnPropertyChanged(nameof(HasValidationError));
                }
            }
        }

        public bool HasValidationError => !string.IsNullOrEmpty(_validationError);

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
