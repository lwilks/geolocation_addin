using System.Collections.Generic;
using System.Linq;
using System.Windows;
using GeolocationAddin.Models;

namespace GeolocationAddin.UI
{
    public partial class LinkSelectionWindow : Window
    {
        private readonly List<LinkMatchInfo> _links;

        public bool Confirmed { get; private set; }

        public LinkSelectionWindow(List<LinkMatchInfo> links)
        {
            _links = links;
            InitializeComponent();

            HeaderText.Text = $"{_links.Count} link(s) found. Select which to process:";
            LinksGrid.ItemsSource = _links;
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in _links.Where(l => l.MatchType != MatchType.None))
                link.IsSelected = true;
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in _links)
                link.IsSelected = false;
        }

        private void SelectExactOnly_Click(object sender, RoutedEventArgs e)
        {
            foreach (var link in _links)
                link.IsSelected = link.MatchType == MatchType.Exact;
        }

        private void ProcessSelected_Click(object sender, RoutedEventArgs e)
        {
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
    }
}
