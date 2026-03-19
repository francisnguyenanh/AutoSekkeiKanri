using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using AutoFiller.Core;

namespace AutoFiller.App
{
    // ─────────────────────────────────────────────
    // Value converters (referenced from XAML)
    // ─────────────────────────────────────────────

    /// <summary>Maps <see cref="NodeMatchStatus"/> → <see cref="SolidColorBrush"/>.</summary>
    public class StatusToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is NodeMatchStatus status)
            {
                return status switch
                {
                    NodeMatchStatus.Matched       => Brushes.Green,
                    NodeMatchStatus.LowConfidence => Brushes.DarkOrange,
                    NodeMatchStatus.Unmatched     => Brushes.Red,
                    _                             => Brushes.Gray
                };
            }
            return Brushes.Gray;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>Maps IsUsed (bool) → row background brush.</summary>
    public class UsedToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && !b ? Brushes.MistyRose : Brushes.Transparent;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>bool → Visibility (true → Visible, false → Collapsed).</summary>
    public class BoolToVisibilityConverter : IValueConverter
    {
        public static readonly BoolToVisibilityConverter Instance = new BoolToVisibilityConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && b ? Visibility.Visible : Visibility.Collapsed;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>bool → Visibility (true → Collapsed, false → Visible).</summary>
    public class InverseBoolToVisibilityConverter : IValueConverter
    {
        public static readonly InverseBoolToVisibilityConverter Instance
            = new InverseBoolToVisibilityConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && b ? Visibility.Collapsed : Visibility.Visible;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    /// <summary>bool → FontWeight (true → Bold, false → Normal).</summary>
    public class BoolToFontWeightConverter : IValueConverter
    {
        public static readonly BoolToFontWeightConverter Instance
            = new BoolToFontWeightConverter();

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && b ? FontWeights.SemiBold : FontWeights.Normal;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }

    // ─────────────────────────────────────────────
    // Excel cell picker dialog
    // ─────────────────────────────────────────────

    /// <summary>
    /// Minimal modal dialog that shows all Excel cells in a DataGrid so the
    /// user can click one row and confirm the assignment.
    /// </summary>
    public class ExcelCellPickerDialog : Window
    {
        public ExcelCellValue ChosenCell { get; private set; }

        private readonly DataGrid _grid;

        public ExcelCellPickerDialog(IReadOnlyList<ExcelCellValue> cells)
        {
            Title = "Assign Excel Cell";
            Width = 480;
            Height = 420;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanResize;

            var root = new DockPanel();

            // Search box
            var search = new TextBox
            {
                Margin = new Thickness(6, 6, 6, 4),
                FontSize = 13
            };
            var hint = new TextBlock
            {
                Text = "Search (address or value)…",
                Foreground = Brushes.Gray,
                IsHitTestVisible = false,
                Margin = new Thickness(10, 8, 6, 4),
                FontSize = 13
            };

            var searchLayer = new Grid();
            searchLayer.Children.Add(search);
            searchLayer.Children.Add(hint);
            DockPanel.SetDock(searchLayer, Dock.Top);
            root.Children.Add(searchLayer);

            // Grid
            _grid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserReorderColumns = false,
                CanUserResizeRows = false,
                SelectionMode = DataGridSelectionMode.Single,
                GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
                BorderThickness = new Thickness(0),
                RowHeaderWidth = 0,
                FontSize = 12
            };
            _grid.Columns.Add(new DataGridTextColumn { Header = "Sheet", Binding = new Binding("SheetName"), Width = 70 });
            _grid.Columns.Add(new DataGridTextColumn { Header = "Cell",  Binding = new Binding("CellAddress"), Width = 60 });
            _grid.Columns.Add(new DataGridTextColumn { Header = "Value", Binding = new Binding("RawValue"),   Width = new DataGridLength(1, DataGridLengthUnitType.Star) });

            var all = cells.OrderBy(c => c.Row).ThenBy(c => c.Col).ToList();
            _grid.ItemsSource = all;
            DockPanel.SetDock(_grid, Dock.Top);
            root.Children.Add(_grid);

            // OK / Cancel
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(6)
            };
            DockPanel.SetDock(btnPanel, Dock.Bottom);

            var ok = new Button
            {
                Content = "Assign",
                Width = 80,
                Height = 28,
                Margin = new Thickness(4, 0, 4, 0),
                IsDefault = true
            };
            ok.Click += (_, _) =>
            {
                if (_grid.SelectedItem is ExcelCellValue cv)
                {
                    ChosenCell = cv;
                    DialogResult = true;
                }
                else
                {
                    MessageBox.Show("Select a cell first.", "No selection",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            };

            var cancel = new Button
            {
                Content = "Cancel",
                Width = 80,
                Height = 28,
                IsCancel = true
            };

            btnPanel.Children.Add(ok);
            btnPanel.Children.Add(cancel);
            root.Children.Add(btnPanel);

            Content = root;

            // Live search filter
            search.TextChanged += (_, _) =>
            {
                hint.Visibility = string.IsNullOrEmpty(search.Text)
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                string q = search.Text.Trim().ToLowerInvariant();
                _grid.ItemsSource = string.IsNullOrEmpty(q)
                    ? all
                    : all.Where(c =>
                        (c.CellAddress ?? "").ToLowerInvariant().Contains(q) ||
                        (c.RawValue ?? "").ToLowerInvariant().Contains(q)).ToList();
            };
        }
    }

    // ─────────────────────────────────────────────
    // Code-behind
    // ─────────────────────────────────────────────

    public partial class MappingReviewWindow : Window
    {
        private MappingReviewViewModel ViewModel => DataContext as MappingReviewViewModel;

        public MappingReviewWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Convenience factory — creates the window, loads the data, and shows it.
        /// </summary>
        public static void ShowReview(
            MappingConfig config,
            IReadOnlyList<ExcelCellValue> excelValues)
        {
            var win = new MappingReviewWindow();
            win.ViewModel?.Load(config, excelValues);
            win.Show();
        }

        // ── event handlers ────────────────────────────────────────────────

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (ViewModel == null) return;

            if (e.OldValue is ControlNode prev) prev.IsSelected = false;
            if (e.NewValue is ControlNode next)
            {
                next.IsSelected = true;
                ViewModel.SelectedNode = next;
            }
        }
    }
}
