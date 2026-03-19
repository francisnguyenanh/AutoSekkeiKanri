using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using AutoFiller.Core;

namespace AutoFiller.App
{
    // ─────────────────────────────────────────────
    // Tree node model
    // ─────────────────────────────────────────────

    public enum NodeMatchStatus { Matched, LowConfidence, Unmatched }

    /// <summary>
    /// One node in the left-panel TreeView. Represents either a Tab,
    /// a Grid container, a grid Row summary, or an individual Control.
    /// </summary>
    public class ControlNode : INotifyPropertyChanged
    {
        private NodeMatchStatus _status;
        private string _matchedCell;
        private double _confidence;
        private bool _isAccepted;
        private bool _isSelected;

        public string DisplayName { get; set; }
        public string CurrentValue { get; set; }
        public string TabContext { get; set; }
        public double ScreenX { get; set; }
        public double ScreenY { get; set; }
        public bool IsHeader { get; set; }   // Tab or Grid header — not a leaf field
        public bool IsLeaf => !IsHeader;

        // Backing MatchResult (null for headers and unmatched controls).
        public MatchResult Match { get; set; }

        public NodeMatchStatus Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusGlyph)); }
        }

        public string MatchedCell
        {
            get => _matchedCell;
            set { _matchedCell = value; OnPropertyChanged(); }
        }

        public double Confidence
        {
            get => _confidence;
            set { _confidence = value; OnPropertyChanged(); }
        }

        public bool IsAccepted
        {
            get => _isAccepted;
            set { _isAccepted = value; OnPropertyChanged(); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public string StatusGlyph => Status switch
        {
            NodeMatchStatus.Matched => "✓",
            NodeMatchStatus.LowConfidence => "?",
            NodeMatchStatus.Unmatched => "✗",
            _ => ""
        };

        public ObservableCollection<ControlNode> Children { get; }
            = new ObservableCollection<ControlNode>();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ─────────────────────────────────────────────
    // Excel cell row (right-panel DataGrid)
    // ─────────────────────────────────────────────

    public class ExcelCellRow : INotifyPropertyChanged
    {
        private bool _isUsed;

        public string SheetName { get; set; }
        public string CellAddress { get; set; }
        public string Value { get; set; }
        public int Row { get; set; }
        public int Col { get; set; }

        public bool IsUsed
        {
            get => _isUsed;
            set { _isUsed = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusText)); }
        }

        public string StatusText => IsUsed ? "✓ used" : "? unmatched";

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ─────────────────────────────────────────────
    // RelayCommand
    // ─────────────────────────────────────────────

    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object parameter) => _execute(parameter);
        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }

    // ─────────────────────────────────────────────
    // ViewModel
    // ─────────────────────────────────────────────

    public class MappingReviewViewModel : INotifyPropertyChanged
    {
        // ── backing fields ────────────────────────────────────────────────
        private ControlNode _selectedNode;
        private ExcelCellRow _selectedExcelCell;
        private bool _showUnmatchedOnly;
        private MappingConfig _mappingConfig;
        private IReadOnlyList<ExcelCellValue> _allExcelValues;

        // ── observable collections ────────────────────────────────────────
        public ObservableCollection<ControlNode> ControlTree { get; }
            = new ObservableCollection<ControlNode>();

        public ObservableCollection<ExcelCellRow> ExcelValues { get; }
            = new ObservableCollection<ExcelCellRow>();

        // ── selected-node detail panel ────────────────────────────────────
        public ControlNode SelectedNode
        {
            get => _selectedNode;
            set
            {
                _selectedNode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(DetailControlName));
                OnPropertyChanged(nameof(DetailTabContext));
                OnPropertyChanged(nameof(DetailPosition));
                OnPropertyChanged(nameof(DetailCurrentValue));
                OnPropertyChanged(nameof(DetailExcelCell));
                OnPropertyChanged(nameof(DetailExcelValue));
                OnPropertyChanged(nameof(DetailConfidence));
                OnPropertyChanged(nameof(HasDetail));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public ExcelCellRow SelectedExcelCell
        {
            get => _selectedExcelCell;
            set { _selectedExcelCell = value; OnPropertyChanged(); }
        }

        public bool HasDetail => _selectedNode?.IsLeaf == true;

        public string DetailControlName => _selectedNode?.DisplayName ?? string.Empty;
        public string DetailTabContext => _selectedNode?.TabContext ?? string.Empty;
        public string DetailPosition
            => _selectedNode != null
               ? $"X={_selectedNode.ScreenX:F0},  Y={_selectedNode.ScreenY:F0}"
               : string.Empty;
        public string DetailCurrentValue => _selectedNode?.CurrentValue ?? string.Empty;
        public string DetailExcelCell => _selectedNode?.MatchedCell ?? "(none)";
        public string DetailExcelValue
        {
            get
            {
                if (_selectedNode?.MatchedCell == null) return string.Empty;
                return ExcelValues
                    .FirstOrDefault(e => e.CellAddress == _selectedNode.MatchedCell)?.Value
                    ?? string.Empty;
            }
        }
        public string DetailConfidence
            => _selectedNode?.Confidence > 0
               ? $"{_selectedNode.Confidence:P0}"
               : "—";

        // ── filter ────────────────────────────────────────────────────────
        public bool ShowUnmatchedOnly
        {
            get => _showUnmatchedOnly;
            set
            {
                _showUnmatchedOnly = value;
                OnPropertyChanged();
                RefreshExcelView();
            }
        }

        // ── statistics ────────────────────────────────────────────────────
        private string _statsText = string.Empty;
        public string StatsText
        {
            get => _statsText;
            private set { _statsText = value; OnPropertyChanged(); }
        }

        // ── commands ──────────────────────────────────────────────────────
        public ICommand AcceptMatchCommand { get; }
        public ICommand RejectMatchCommand { get; }
        public ICommand AssignManuallyCommand { get; }
        public ICommand AutoMatchRemainingCommand { get; }
        public ICommand ExportMappingCommand { get; }
        public ICommand ExportReportCommand { get; }

        // ── constructor ───────────────────────────────────────────────────
        public MappingReviewViewModel()
        {
            AcceptMatchCommand = new RelayCommand(
                _ => AcceptMatch(),
                _ => _selectedNode?.IsLeaf == true
                     && _selectedNode.Status != NodeMatchStatus.Unmatched);

            RejectMatchCommand = new RelayCommand(
                _ => RejectMatch(),
                _ => _selectedNode?.IsLeaf == true
                     && _selectedNode.Status != NodeMatchStatus.Unmatched);

            AssignManuallyCommand = new RelayCommand(
                _ => AssignManually(),
                _ => _selectedNode?.IsLeaf == true);

            AutoMatchRemainingCommand = new RelayCommand(
                _ => AutoMatchRemaining());

            ExportMappingCommand = new RelayCommand(
                _ => ExportMapping(),
                _ => _mappingConfig != null);

            ExportReportCommand = new RelayCommand(
                _ => ExportReport(),
                _ => _mappingConfig != null);
        }

        // ── public initialiser ────────────────────────────────────────────

        /// <summary>
        /// Binds the view-model to a <see cref="MappingConfig"/> and the full
        /// Excel cell list. Call this after the auto-match run completes.
        /// </summary>
        public void Load(MappingConfig config, IReadOnlyList<ExcelCellValue> excelValues)
        {
            _mappingConfig = config ?? throw new ArgumentNullException(nameof(config));
            _allExcelValues = excelValues ?? throw new ArgumentNullException(nameof(excelValues));

            BuildControlTree();
            BuildExcelGrid();
            UpdateStats();
        }

        // ── command implementations ───────────────────────────────────────

        private void AcceptMatch()
        {
            if (_selectedNode == null) return;
            _selectedNode.IsAccepted = true;
            _selectedNode.Status = NodeMatchStatus.Matched;
            MarkExcelCellUsed(_selectedNode.MatchedCell, used: true);
            UpdateStats();
        }

        private void RejectMatch()
        {
            if (_selectedNode == null) return;
            MarkExcelCellUsed(_selectedNode.MatchedCell, used: false);
            _selectedNode.Match = null;
            _selectedNode.MatchedCell = null;
            _selectedNode.Confidence = 0;
            _selectedNode.IsAccepted = false;
            _selectedNode.Status = NodeMatchStatus.Unmatched;

            // Remove from MappingConfig HeaderFields.
            _mappingConfig?.HeaderFields.RemoveAll(
                f => f.ControlName == _selectedNode.DisplayName
                     && f.TabContext == _selectedNode.TabContext);

            OnPropertyChanged(nameof(DetailExcelCell));
            OnPropertyChanged(nameof(DetailExcelValue));
            OnPropertyChanged(nameof(DetailConfidence));
            UpdateStats();
        }

        private void AssignManually()
        {
            // Open cell-picker dialog.
            var picker = new ExcelCellPickerDialog(_allExcelValues)
            {
                Owner = System.Windows.Application.Current?.MainWindow
            };

            if (picker.ShowDialog() != true) return;

            var chosen = picker.ChosenCell;
            if (chosen == null) return;

            // Remove previous usage.
            MarkExcelCellUsed(_selectedNode.MatchedCell, used: false);

            _selectedNode.MatchedCell = chosen.CellAddress;
            _selectedNode.Confidence = 1.0;
            _selectedNode.IsAccepted = true;
            _selectedNode.Status = NodeMatchStatus.Matched;
            MarkExcelCellUsed(chosen.CellAddress, used: true);

            // Upsert HeaderFieldMapping.
            if (_mappingConfig != null)
            {
                var existing = _mappingConfig.HeaderFields
                    .FirstOrDefault(f => f.ControlName == _selectedNode.DisplayName
                                         && f.TabContext == _selectedNode.TabContext);
                if (existing != null)
                {
                    existing.ExcelCellAddress = chosen.CellAddress;
                    existing.ExcelRow = chosen.Row;
                    existing.ExcelCol = chosen.Col;
                }
                else
                {
                    _mappingConfig.HeaderFields.Add(new HeaderFieldMapping
                    {
                        ControlName = _selectedNode.DisplayName,
                        TabContext = _selectedNode.TabContext,
                        ClickX = _selectedNode.ScreenX,
                        ClickY = _selectedNode.ScreenY,
                        ExcelCellAddress = chosen.CellAddress,
                        ExcelRow = chosen.Row,
                        ExcelCol = chosen.Col,
                        InputMethod = "SendKeys"
                    });
                }
            }

            OnPropertyChanged(nameof(DetailExcelCell));
            OnPropertyChanged(nameof(DetailExcelValue));
            OnPropertyChanged(nameof(DetailConfidence));
            UpdateStats();
        }

        private void AutoMatchRemaining()
        {
            // Collect controls that are still unmatched.
            var unmatchedNodes = AllLeafNodes()
                .Where(n => n.Status == NodeMatchStatus.Unmatched)
                .ToList();

            var usedAddresses = new HashSet<string>(
                AllLeafNodes()
                    .Where(n => n.MatchedCell != null)
                    .Select(n => n.MatchedCell),
                StringComparer.Ordinal);

            // Available Excel values (unused, non-trivial).
            var available = _allExcelValues
                .Where(e => !usedAddresses.Contains(e.CellAddress))
                .ToList();

            var matcher = new ValueMatcher();

            foreach (var node in unmatchedNodes)
            {
                if (string.IsNullOrWhiteSpace(node.CurrentValue)) continue;

                string norm = new ExcelValueExtractor().Normalize(node.CurrentValue);
                var candidate = available.FirstOrDefault(
                    e => string.Equals(e.NormalizedValue, norm, StringComparison.Ordinal));

                if (candidate == null) continue;

                node.MatchedCell = candidate.CellAddress;
                node.Confidence = 0.9;
                node.Status = NodeMatchStatus.Matched;
                MarkExcelCellUsed(candidate.CellAddress, used: true);
                available.Remove(candidate);
            }

            UpdateStats();
        }

        private void ExportMapping()
        {
            if (_mappingConfig == null) return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"mapping_{DateTime.Now:yyyyMMdd_HHmmss}.json",
                DefaultExt = ".json",
                Filter = "JSON files (*.json)|*.json"
            };
            if (dlg.ShowDialog() != true) return;

            // Sync accepted flags back to config before saving.
            SyncAcceptedToConfig();
            new ValueMatcher().SaveMappingConfig(_mappingConfig, dlg.FileName);
        }

        private void ExportReport()
        {
            if (_mappingConfig == null) return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"mapping_report_{DateTime.Now:yyyyMMdd_HHmmss}.md",
                DefaultExt = ".md",
                Filter = "Markdown (*.md)|*.md|HTML (*.html)|*.html"
            };
            if (dlg.ShowDialog() != true) return;

            SyncAcceptedToConfig();
            new ValueMatcher().SaveMappingReport(_mappingConfig, dlg.FileName);
        }

        // ── tree building ─────────────────────────────────────────────────

        private void BuildControlTree()
        {
            ControlTree.Clear();
            if (_mappingConfig == null) return;

            // Build a lookup: controlName+tab → MatchResult via HeaderFields.
            var fieldLookup = _mappingConfig.HeaderFields.ToDictionary(
                f => MakeKey(f.ControlName, f.TabContext),
                f => f);

            // Group header fields by tab.
            var byTab = _mappingConfig.HeaderFields
                .GroupBy(f => f.TabContext ?? "(unknown)")
                .ToDictionary(g => g.Key, g => g.ToList());

            // Add unmatched controls too.
            var unmatchedByTab = (_mappingConfig.UnmatchedControls ?? new List<UnmatchedControl>())
                .GroupBy(u => u.TabContext ?? "(unknown)")
                .ToDictionary(g => g.Key, g => g.ToList());

            var allTabs = byTab.Keys.Union(unmatchedByTab.Keys).OrderBy(t => t);

            foreach (string tab in allTabs)
            {
                var tabNode = new ControlNode
                {
                    DisplayName = tab,
                    IsHeader = true,
                    TabContext = tab
                };

                // Matched fields.
                if (byTab.TryGetValue(tab, out var fields))
                {
                    foreach (var f in fields.OrderBy(f => f.ExcelRow))
                    {
                        double conf = 1.0; // default for accepted
                        var child = new ControlNode
                        {
                            DisplayName = f.ControlName,
                            TabContext = f.TabContext,
                            ScreenX = f.ClickX,
                            ScreenY = f.ClickY,
                            MatchedCell = f.ExcelCellAddress,
                            Confidence = conf,
                            Status = conf >= 0.8
                                ? NodeMatchStatus.Matched
                                : NodeMatchStatus.LowConfidence
                        };
                        tabNode.Children.Add(child);
                    }
                }

                // Unmatched controls.
                if (unmatchedByTab.TryGetValue(tab, out var unmatched))
                {
                    foreach (var u in unmatched)
                    {
                        tabNode.Children.Add(new ControlNode
                        {
                            DisplayName = u.ControlName,
                            TabContext = u.TabContext,
                            CurrentValue = u.CurrentValue,
                            ScreenX = u.X,
                            ScreenY = u.Y,
                            Status = NodeMatchStatus.Unmatched
                        });
                    }
                }

                // Grid node.
                if (_mappingConfig.Grid?.TabContext == tab
                    && _mappingConfig.Grid.Columns?.Count > 0)
                {
                    var gridNode = new ControlNode
                    {
                        DisplayName = $"Grid ({_mappingConfig.Grid.Columns.Count} columns)",
                        IsHeader = true,
                        TabContext = tab
                    };
                    foreach (var kv in _mappingConfig.Grid.Columns.OrderBy(c => c.Value.ExcelColIndex))
                    {
                        gridNode.Children.Add(new ControlNode
                        {
                            DisplayName = kv.Key,
                            TabContext = tab,
                            MatchedCell = $"Col {kv.Value.ExcelColLetter}",
                            Confidence = 0.9,
                            Status = NodeMatchStatus.Matched
                        });
                    }
                    tabNode.Children.Add(gridNode);
                }

                ControlTree.Add(tabNode);
            }
        }

        private void BuildExcelGrid()
        {
            ExcelValues.Clear();
            if (_allExcelValues == null) return;

            var usedAddresses = new HashSet<string>(
                _mappingConfig?.HeaderFields.Select(f => f.ExcelCellAddress) ?? Enumerable.Empty<string>(),
                StringComparer.Ordinal);

            foreach (var cell in _allExcelValues.OrderBy(c => c.Row).ThenBy(c => c.Col))
            {
                ExcelValues.Add(new ExcelCellRow
                {
                    SheetName = cell.SheetName,
                    CellAddress = cell.CellAddress,
                    Value = cell.RawValue,
                    Row = cell.Row,
                    Col = cell.Col,
                    IsUsed = usedAddresses.Contains(cell.CellAddress)
                });
            }
        }

        // ── helpers ───────────────────────────────────────────────────────

        private void RefreshExcelView()
        {
            // Re-apply filter: use a CollectionView so the ObservableCollection
            // is not rebuilt from scratch.
            var view = System.Windows.Data.CollectionViewSource.GetDefaultView(ExcelValues);
            if (view == null) return;

            if (_showUnmatchedOnly)
                view.Filter = o => o is ExcelCellRow r && !r.IsUsed;
            else
                view.Filter = null;
        }

        private void MarkExcelCellUsed(string cellAddress, bool used)
        {
            if (string.IsNullOrEmpty(cellAddress)) return;
            var row = ExcelValues.FirstOrDefault(e => e.CellAddress == cellAddress);
            if (row != null) row.IsUsed = used;
        }

        private void UpdateStats()
        {
            var leaves = AllLeafNodes().ToList();
            int total = leaves.Count;
            int matched = leaves.Count(n => n.Status != NodeMatchStatus.Unmatched);
            double avgConf = matched > 0
                ? leaves.Where(n => n.Status != NodeMatchStatus.Unmatched)
                        .Average(n => n.Confidence)
                : 0.0;

            StatsText = $"Matched: {matched}/{total} controls  |  Confidence avg: {avgConf:F2}";
        }

        private void SyncAcceptedToConfig()
        {
            // Remove non-accepted nodes from HeaderFields before export.
            if (_mappingConfig == null) return;
            var acceptedNames = AllLeafNodes()
                .Where(n => n.IsAccepted && n.MatchedCell != null)
                .Select(n => MakeKey(n.DisplayName, n.TabContext))
                .ToHashSet();

            _mappingConfig.HeaderFields.RemoveAll(
                f => !acceptedNames.Contains(MakeKey(f.ControlName, f.TabContext)));
        }

        private IEnumerable<ControlNode> AllLeafNodes()
            => ControlTree.SelectMany(AllLeafNodes);

        private static IEnumerable<ControlNode> AllLeafNodes(ControlNode node)
        {
            if (node.IsLeaf) yield return node;
            foreach (var child in node.Children)
            foreach (var leaf in AllLeafNodes(child))
                yield return leaf;
        }

        private static string MakeKey(string name, string tab)
            => $"{tab ?? ""}|{name ?? ""}";

        // ── INotifyPropertyChanged ────────────────────────────────────────
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
