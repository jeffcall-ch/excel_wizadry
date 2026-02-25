using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace WeekDateCalc;

public partial class MainWindow : Window
{
    private const int MinSequenceInputRows = 10;

    private readonly HistoryStore _historyStore = new();
    private readonly ObservableCollection<HistoryEntry> _allHistoryEntries = [];

    public ObservableCollection<SequenceRow> SequenceRows { get; } = [];

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;

        SequenceRows.CollectionChanged += SequenceRows_CollectionChanged;
        EnsureTrailingEditableRows();

        Loaded += async (_, _) =>
        {
            var loaded = await _historyStore.LoadAsync();
            foreach (var entry in loaded.OrderByDescending(x => x.Timestamp))
            {
                _allHistoryEntries.Add(entry);
            }

            ApplyHistoryFilter();
        };
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers != ModifierKeys.None)
        {
            return;
        }

        if (Keyboard.FocusedElement is Button)
        {
            return;
        }

        if (MainTabControl.SelectedIndex == 0)
        {
            CalculateDifference();
            e.Handled = true;
            return;
        }

        if (MainTabControl.SelectedIndex == 1)
        {
            CalculateShift();
            e.Handled = true;
            return;
        }

        if (MainTabControl.SelectedIndex == 2)
        {
            SaveSequenceSnapshot();
            e.Handled = true;
        }
    }

    private void CalculateDifference_Click(object sender, RoutedEventArgs e)
    {
        CalculateDifference();
    }

    private void CalculateDifference()
    {
        DifferenceStatusText.Text = string.Empty;

        if (!DateUtilities.TryParseDate(DateAInput.Text, out var dateA) ||
            !DateUtilities.TryParseDate(DateBInput.Text, out var dateB))
        {
            DifferenceResultText.Text = "-";
            DifferenceStatusText.Text = "Enter valid dates in DD/MM/YYYY or DD.MM.YYYY.";
            return;
        }

        var weeks = DateUtilities.WeeksBetween(dateA, dateB);
        DifferenceResultText.Text = $"{weeks:0.0} weeks";

        AddHistoryEntry(
            mode: "Difference",
            input: $"{dateA:dd/MM/yyyy} <-> {dateB:dd/MM/yyyy}",
            result: $"{weeks:0.0} weeks");
    }

    private void CalculateShift_Click(object sender, RoutedEventArgs e)
    {
        CalculateShift();
    }

    private void CalculateShift()
    {
        ShiftStatusText.Text = string.Empty;

        if (!DateUtilities.TryParseDate(BaseDateInput.Text, out var baseDate))
        {
            ShiftResultText.Text = "-";
            ShiftStatusText.Text = "Enter a valid date in DD/MM/YYYY or DD.MM.YYYY.";
            return;
        }

        if (!int.TryParse(WeeksToAddInput.Text?.Trim(), out var weeks))
        {
            ShiftResultText.Text = "-";
            ShiftStatusText.Text = "Weeks must be a whole number (for example 5 or -5).";
            return;
        }

        var shifted = DateUtilities.AddWeeksWithWorkdayRule(baseDate, weeks);
        ShiftResultText.Text = $"{shifted:dd/MM/yyyy} ({shifted:dddd})";

        AddHistoryEntry(
            mode: "Shift",
            input: $"{baseDate:dd/MM/yyyy}, weeks={weeks}",
            result: $"{shifted:dd/MM/yyyy}");
    }

    private async void ClearHistory_Click(object sender, RoutedEventArgs e)
    {
        _allHistoryEntries.Clear();
        await _historyStore.SaveAsync(_allHistoryEntries);
        ApplyHistoryFilter();
    }

    private void HistoryFilterInput_TextChanged(object sender, TextChangedEventArgs e)
    {
        ApplyHistoryFilter();
    }

    private void SaveSequenceSnapshot_Click(object sender, RoutedEventArgs e)
    {
        SaveSequenceSnapshot();
    }

    private void SaveSequenceSnapshot()
    {
        var usedRows = SequenceRows
            .Where(r => !string.IsNullOrWhiteSpace(r.DateInput))
            .ToList();

        if (usedRows.Count < 2)
        {
            SequenceStatusText.Text = "Add at least two dates before saving.";
            return;
        }

        var input = string.Join(" | ", usedRows.Select(r => r.DateInput.Trim()));
        var result = string.Join(" | ", usedRows.Select(r => string.IsNullOrWhiteSpace(r.WeekDifference) ? "-" : r.WeekDifference));

        AddHistoryEntry("Row List", input, result);
        SequenceStatusText.Text = "Sequence snapshot saved to history.";
    }

    private void HistoryListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (HistoryListView.SelectedItem is not HistoryEntry entry)
        {
            return;
        }

        RestoreHistoryEntry(entry);
    }

    private void RestoreHistoryEntry(HistoryEntry entry)
    {
        if (entry.Mode == "Difference")
        {
            var parts = entry.Input.Split("<->", StringSplitOptions.TrimEntries);
            if (parts.Length == 2)
            {
                DateAInput.Text = parts[0];
                DateBInput.Text = parts[1];
                MainTabControl.SelectedIndex = 0;
                DateAInput.Focus();
                DateAInput.SelectAll();
            }

            return;
        }

        if (entry.Mode == "Shift")
        {
            var parts = entry.Input.Split(',', StringSplitOptions.TrimEntries);
            if (parts.Length >= 2)
            {
                BaseDateInput.Text = parts[0];
                var weeksPart = parts[1].Replace("weeks=", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
                WeeksToAddInput.Text = weeksPart;
                MainTabControl.SelectedIndex = 1;
                BaseDateInput.Focus();
                BaseDateInput.SelectAll();
            }

            return;
        }

        if (entry.Mode == "Row List")
        {
            var dates = entry.Input
                .Split('|', StringSplitOptions.TrimEntries)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            SequenceRows.Clear();
            foreach (var date in dates)
            {
                SequenceRows.Add(new SequenceRow { DateInput = date });
            }

            RecalculateSequenceRows();
            EnsureTrailingEditableRows();

            MainTabControl.SelectedIndex = 2;
            SequenceGrid.Focus();
        }
    }

    private void SequenceRows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<SequenceRow>())
            {
                item.PropertyChanged += SequenceRow_PropertyChanged;
            }
        }

        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<SequenceRow>())
            {
                item.PropertyChanged -= SequenceRow_PropertyChanged;
            }
        }
    }

    private void SequenceRow_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(SequenceRow.DateInput))
        {
            return;
        }

        RecalculateSequenceRows();
        EnsureTrailingEditableRows();
    }

    private void EnsureTrailingEditableRows()
    {
        var lastUsedIndex = -1;
        for (var i = SequenceRows.Count - 1; i >= 0; i--)
        {
            if (!string.IsNullOrWhiteSpace(SequenceRows[i].DateInput))
            {
                lastUsedIndex = i;
                break;
            }
        }

        var targetCount = (lastUsedIndex + 1) + MinSequenceInputRows;
        while (SequenceRows.Count < targetCount)
        {
            SequenceRows.Add(new SequenceRow());
        }

        while (SequenceRows.Count > targetCount && string.IsNullOrWhiteSpace(SequenceRows[^1].DateInput))
        {
            SequenceRows.RemoveAt(SequenceRows.Count - 1);
        }
    }

    private void RecalculateSequenceRows()
    {
        DateTime? previousValidDate = null;

        foreach (var row in SequenceRows)
        {
            if (string.IsNullOrWhiteSpace(row.DateInput))
            {
                row.WeekDifference = string.Empty;
                continue;
            }

            if (!DateUtilities.TryParseDate(row.DateInput, out var currentDate))
            {
                row.WeekDifference = "Invalid date";
                continue;
            }

            if (previousValidDate is null)
            {
                row.WeekDifference = "-";
            }
            else
            {
                var weeks = DateUtilities.WeeksBetween(previousValidDate.Value, currentDate);
                row.WeekDifference = $"{weeks:0.0}";
            }

            previousValidDate = currentDate;
        }
    }

    private async void AddHistoryEntry(string mode, string input, string result)
    {
        _allHistoryEntries.Insert(0, new HistoryEntry
        {
            Timestamp = DateTime.Now,
            Mode = mode,
            Input = input,
            Result = result
        });

        ApplyHistoryFilter();
        await _historyStore.SaveAsync(_allHistoryEntries);
    }

    private void ApplyHistoryFilter()
    {
        var filter = HistoryFilterInput?.Text?.Trim() ?? string.Empty;
        var filtered = string.IsNullOrWhiteSpace(filter)
            ? _allHistoryEntries
            : new ObservableCollection<HistoryEntry>(
                _allHistoryEntries.Where(h =>
                    h.Input.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                    h.Result.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                    h.Mode.Contains(filter, StringComparison.OrdinalIgnoreCase)));

        HistoryListView.ItemsSource = filtered;
    }
}
