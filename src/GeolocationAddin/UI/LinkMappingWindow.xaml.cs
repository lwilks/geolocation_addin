using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using GeolocationAddin.Config;
using GeolocationAddin.Helpers;
using GeolocationAddin.Models;
using Newtonsoft.Json;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;

namespace GeolocationAddin.UI
{
    public partial class LinkMappingWindow : Window
    {
        private readonly List<LinkMatchInfo> _links;
        private readonly AddinConfig _config;

        public bool Confirmed { get; private set; }

        public LinkMappingWindow(List<LinkMatchInfo> links, AddinConfig config)
        {
            _links = links;

            // Deep copy config so edits are isolated until explicitly saved
            var json = JsonConvert.SerializeObject(config);
            _config = JsonConvert.DeserializeObject<AddinConfig>(json);

            InitializeComponent();

            HeaderText.Text = $"{_links.Count} link(s) found — edit target file names, then process:";
            LinksGrid.ItemsSource = _links;

            // Subscribe to TargetFileName changes for duplicate validation
            foreach (var link in _links)
                link.PropertyChanged += Link_PropertyChanged;

            LoadSettingsFromConfig();
            ValidateDuplicates();
        }

        public AddinConfig Config => _config;

        private void Link_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(LinkMatchInfo.TargetFileName))
                ValidateDuplicates();
        }

        private void ValidateDuplicates()
        {
            // Group non-empty targets case-insensitively
            var groups = _links
                .Where(l => l.HasTargetFileName)
                .GroupBy(l => l.TargetFileName.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();

            var duplicateSet = new HashSet<LinkMatchInfo>(groups.SelectMany(g => g));

            foreach (var link in _links)
            {
                link.ValidationError = duplicateSet.Contains(link)
                    ? "Duplicate target file name"
                    : null;
            }

            bool hasDuplicates = duplicateSet.Count > 0;
            ValidationWarning.Visibility = hasDuplicates ? Visibility.Visible : Visibility.Collapsed;
            ProcessButton.IsEnabled = !hasDuplicates;
        }

        #region Import / Export

        public void MergeImportedMappings(List<(string linkName, string targetFileName, string label, string exportViewName)> imported)
        {
            var fuzzySettings = _config.FuzzyMatchSettings;

            foreach (var (linkName, targetFileName, label, exportViewName) in imported)
            {
                // Try exact match on InstanceName
                var exact = _links.FirstOrDefault(l =>
                    string.Equals(l.InstanceName, linkName, StringComparison.OrdinalIgnoreCase));

                if (exact != null)
                {
                    exact.TargetFileName = targetFileName;
                    exact.Label = label;
                    exact.ExportViewName = exportViewName;
                    exact.MatchedImportKey = linkName;
                    exact.MatchType = MatchType.Exact;
                    exact.IsSelected = true;
                    continue;
                }

                // Fuzzy fallback
                if (fuzzySettings.Enabled)
                {
                    var candidates = _links
                        .Where(l => !l.HasTargetFileName)
                        .Select(l => l.InstanceName)
                        .ToList();

                    var fuzzyResult = FuzzyMatcher.FindBestMatch(
                        linkName, candidates,
                        fuzzySettings.TokenOverlapThreshold,
                        fuzzySettings.LevenshteinThreshold);

                    if (fuzzyResult != null)
                    {
                        var match = _links.First(l =>
                            string.Equals(l.InstanceName, fuzzyResult.MatchedKey, StringComparison.OrdinalIgnoreCase));

                        match.TargetFileName = targetFileName;
                        match.Label = label;
                        match.ExportViewName = exportViewName;
                        match.MatchedImportKey = linkName;
                        match.MatchType = MatchType.Fuzzy;
                        // Fuzzy matches are not auto-selected — user reviews
                    }
                }
            }

            ValidateDuplicates();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Mapping files (*.csv;*.xml)|*.csv;*.xml|CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml|All files (*.*)|*.*",
                Title = "Import Link Mappings"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                var imported = MappingSerializer.Import(dialog.FileName);
                MergeImportedMappings(imported);
                LogHelper.Info($"Imported {imported.Count} mapping(s) from: {dialog.FileName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to import mappings:\n\n{ex.Message}",
                    "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|XML files (*.xml)|*.xml",
                Title = "Export Link Mappings",
                DefaultExt = ".csv"
            };

            if (dialog.ShowDialog() != true)
                return;

            try
            {
                var mappings = _links
                    .Where(l => l.HasTargetFileName)
                    .Select(l => (l.InstanceName, l.TargetFileName, l.Label ?? "", l.ExportViewName ?? ""))
                    .ToList();

                MappingSerializer.Export(dialog.FileName, mappings);
                LogHelper.Info($"Exported {mappings.Count} mapping(s) to: {dialog.FileName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to export mappings:\n\n{ex.Message}",
                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Selection buttons

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in _links.Where(l => l.HasTargetFileName))
                link.IsSelected = true;
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in _links)
                link.IsSelected = false;
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in _links)
            {
                link.TargetFileName = null;
                link.Label = null;
                link.ExportViewName = null;
                link.MatchedImportKey = null;
                link.MatchType = MatchType.None;
            }
        }

        #endregion

        #region Dialog buttons

        private void ProcessSelected_Click(object sender, RoutedEventArgs e)
        {
            SaveSettingsToConfig();
            Confirmed = true;
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Confirmed = false;
            DialogResult = false;
            Close();
        }

        #endregion

        #region Settings tab

        private void LoadSettingsFromConfig()
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

        private void SaveSettingsToConfig()
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

        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            SaveSettingsToConfig();

            try
            {
                ConfigLoader.Save(_config);
                MessageBox.Show("Settings saved.", "Settings",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings:\n\n{ex.Message}",
                    "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        #endregion
    }
}
