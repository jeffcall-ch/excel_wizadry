namespace WeekDateCalc;

public sealed class HistoryEntry
{
    public DateTime Timestamp { get; set; }

    public string Mode { get; set; } = string.Empty;

    public string Input { get; set; } = string.Empty;

    public string Result { get; set; } = string.Empty;
}
