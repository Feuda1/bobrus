using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Bobrus.App.Services;

namespace Bobrus.App
{
    public partial class PluginSelectionWindow : Window
    {
        private readonly PluginRepository _repo = new();
        private List<PluginInfo> _allPlugins = new();
        private readonly List<PluginVersion> _addedPlugins = new();
        
        public IReadOnlyList<PluginVersion> AddedPlugins => _addedPlugins;

        public PluginSelectionWindow()
        {
            InitializeComponent();
            Loaded += OnLoaded;
        }

        private void OnTitleBarMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
                DragMove();
        }

        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                StatusText.Text = "Загрузка списка...";
                _allPlugins = await _repo.GetPluginsAsync();
                PluginsList.ItemsSource = _allPlugins;
                StatusText.Text = $"Найдено: {_allPlugins.Count}";
            }
            catch (Exception ex)
            {
                StatusText.Text = "Ошибка: " + ex.Message;
            }
        }

        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            var query = SearchBox.Text.Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(query))
            {
                PluginsList.ItemsSource = _allPlugins;
            }
            else
            {
                PluginsList.ItemsSource = _allPlugins
                    .Where(p => p.DisplayName.ToLowerInvariant().Contains(query))
                    .ToList();
            }
        }

        private async void OnPluginSelected(object sender, SelectionChangedEventArgs e)
        {
            VersionsList.ItemsSource = null;
            AddButton.IsEnabled = false;

            if (PluginsList.SelectedItem is PluginInfo plugin)
            {
                try
                {
                    StatusText.Text = "Загрузка версий...";
                    var versions = await _repo.GetVersionsAsync(plugin.Url);
                    VersionsList.ItemsSource = versions;
                    StatusText.Text = $"Версий: {versions.Count}";
                }
                catch (Exception ex)
                {
                    StatusText.Text = "Ошибка версий: " + ex.Message;
                }
            }
        }

        private void OnVersionSelected(object sender, SelectionChangedEventArgs e)
        {
            AddButton.IsEnabled = VersionsList.SelectedItem is PluginVersion;
        }

        private void OnAddClicked(object sender, RoutedEventArgs e)
        {
            if (VersionsList.SelectedItem is PluginVersion version)
            {
                if (!_addedPlugins.Any(p => p.Url == version.Url))
                {
                    _addedPlugins.Add(version);
                    StatusText.Text = $"✔ {version.Name} добавлен! (всего: {_addedPlugins.Count})";
                }
                else
                {
                    StatusText.Text = $"Плагин уже добавлен";
                }
                VersionsList.SelectedItem = null;
                AddButton.IsEnabled = false;
            }
        }

        private void OnCloseClicked(object sender, RoutedEventArgs e)
        {
            DialogResult = _addedPlugins.Count > 0;
            Close();
        }
    }
}
