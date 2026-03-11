using System.Windows;
using System.Windows.Forms;
using GeolocationAddin.Config;
using Newtonsoft.Json;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace GeolocationAddin.UI
{
    public partial class SettingsWindow : Window
    {
        private readonly AddinConfig _config;

        public SettingsWindow(AddinConfig config)
        {
            // Deep copy via serialize/deserialize so we only commit on Save
            var json = JsonConvert.SerializeObject(config);
            _config = JsonConvert.DeserializeObject<AddinConfig>(json);

            InitializeComponent();
            LoadFromConfig();
        }

        public AddinConfig Config => _config;

        private void LoadFromConfig()
        {
            CsvMappingPathBox.Text = _config.CsvMappingPath ?? "";
            LinkSourceFolderBox.Text = _config.LinkSourceFolder ?? "";
            OutputFolderBox.Text = _config.OutputFolder ?? "";
            IfcOutputFolderBox.Text = _config.IfcOutputFolder ?? "";
            NwcOutputFolderBox.Text = _config.NwcOutputFolder ?? "";
            DwgOutputFolderBox.Text = _config.DwgOutputFolder ?? "";

            ExportIfcCheck.IsChecked = _config.ExportSettings.ExportIfc;
            ExportNwcCheck.IsChecked = _config.ExportSettings.ExportNwc;
            ExportDwgCheck.IsChecked = _config.ExportSettings.ExportDwg;
            IfcExportConfigPathBox.Text = _config.ExportSettings.IfcExportConfigPath ?? "";

            FuzzyEnabledCheck.IsChecked = _config.FuzzyMatchSettings.Enabled;
            TokenThresholdSlider.Value = _config.FuzzyMatchSettings.TokenOverlapThreshold;
            LevenshteinThresholdSlider.Value = _config.FuzzyMatchSettings.LevenshteinThreshold;

            UpdateExportFieldStates();
        }

        private void SaveToConfig()
        {
            _config.CsvMappingPath = CsvMappingPathBox.Text.Trim();
            _config.LinkSourceFolder = LinkSourceFolderBox.Text.Trim();
            _config.OutputFolder = OutputFolderBox.Text.Trim();
            _config.IfcOutputFolder = IfcOutputFolderBox.Text.Trim();
            _config.NwcOutputFolder = NwcOutputFolderBox.Text.Trim();
            _config.DwgOutputFolder = DwgOutputFolderBox.Text.Trim();

            _config.ExportSettings.ExportIfc = ExportIfcCheck.IsChecked == true;
            _config.ExportSettings.ExportNwc = ExportNwcCheck.IsChecked == true;
            _config.ExportSettings.ExportDwg = ExportDwgCheck.IsChecked == true;
            _config.ExportSettings.IfcExportConfigPath = string.IsNullOrWhiteSpace(IfcExportConfigPathBox.Text) ? null : IfcExportConfigPathBox.Text.Trim();

            _config.FuzzyMatchSettings.Enabled = FuzzyEnabledCheck.IsChecked == true;
            _config.FuzzyMatchSettings.TokenOverlapThreshold = TokenThresholdSlider.Value;
            _config.FuzzyMatchSettings.LevenshteinThreshold = LevenshteinThresholdSlider.Value;
        }

        private void UpdateExportFieldStates()
        {
            IfcOutputFolderBox.IsEnabled = ExportIfcCheck.IsChecked == true;
            IfcExportConfigPathBox.IsEnabled = ExportIfcCheck.IsChecked == true;
            NwcOutputFolderBox.IsEnabled = ExportNwcCheck.IsChecked == true;
            DwgOutputFolderBox.IsEnabled = ExportDwgCheck.IsChecked == true;
        }

        private void ExportCheck_Changed(object sender, RoutedEventArgs e)
        {
            UpdateExportFieldStates();
        }

        private void BrowseCsvPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                Title = "Select CSV Mapping File"
            };

            if (dialog.ShowDialog() == true)
                CsvMappingPathBox.Text = dialog.FileName;
        }

        private void BrowseIfcConfigPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Select IFC Export Configuration File"
            };

            if (dialog.ShowDialog() == true)
                IfcExportConfigPathBox.Text = dialog.FileName;
        }

        private void BrowseFolder_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as System.Windows.Controls.Button;
            var targetBoxName = button?.Tag as string;
            if (targetBoxName == null) return;

            var targetBox = FindName(targetBoxName) as System.Windows.Controls.TextBox;
            if (targetBox == null) return;

            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select folder";
                if (!string.IsNullOrEmpty(targetBox.Text))
                    dialog.SelectedPath = targetBox.Text;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    targetBox.Text = dialog.SelectedPath;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveToConfig();
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
