using System.IO;
using System.Text.Json;

namespace WeekDateCalc;

public sealed class HistoryStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly string _historyFilePath;

    public HistoryStore()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var dir = Path.Combine(appData, "WeekDateWizard");
        Directory.CreateDirectory(dir);
        _historyFilePath = Path.Combine(dir, "history.json");
    }

    public async Task<List<HistoryEntry>> LoadAsync()
    {
        if (!File.Exists(_historyFilePath))
        {
            return [];
        }

        try
        {
            await using var stream = File.OpenRead(_historyFilePath);
            var entries = await JsonSerializer.DeserializeAsync<List<HistoryEntry>>(stream, JsonOptions);
            return entries ?? [];
        }
        catch
        {
            return [];
        }
    }

    public async Task SaveAsync(IEnumerable<HistoryEntry> entries)
    {
        await using var stream = File.Create(_historyFilePath);
        await JsonSerializer.SerializeAsync(stream, entries, JsonOptions);
    }
}
