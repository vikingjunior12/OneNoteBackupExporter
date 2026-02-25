using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace OneNoteExporter.Models;

public class NotebookInfo
{
    public string Id           { get; set; } = "";
    public string Name         { get; set; } = "";
    public string Path         { get; set; } = "";
    public string LastModified { get; set; } = "";
    public bool IsCurrentlyViewed { get; set; }
}

public class ExportResult
{
    public bool   Success      { get; set; }
    public string Message      { get; set; } = "";
    public string ExportedPath { get; set; } = "";
}

public class VersionInfo
{
    public string Version           { get; set; } = "2.0.0";
    public bool   OneNoteInstalled  { get; set; }
    public string OneNoteVersion    { get; set; } = "";
}

public class BackupAvailability
{
    public bool   Available      { get; set; }
    public string Reason         { get; set; } = "";
    public string Message        { get; set; } = "";
    public string Path           { get; set; } = "";
    public int    NotebookCount  { get; set; }
}

/// <summary>
/// Wraps NotebookInfo for the WPF notebook list with selection state.
/// </summary>
public class NotebookViewModel : INotifyPropertyChanged
{
    private bool _isSelected;
    private bool _isEnabled = true;

    public NotebookInfo Info { get; }

    public string Id   => Info.Id;
    public string Name => Info.Name;

    public string DisplayName =>
        Info.IsCurrentlyViewed
            ? $"📓 {Info.Name}  (currently open)"
            : $"📓 {Info.Name}";

    public string LastModifiedFormatted
    {
        get
        {
            if (DateTime.TryParse(Info.LastModified, null,
                    System.Globalization.DateTimeStyles.RoundtripKind, out var dt))
                return $"Last modified: {dt.ToLocalTime():dd.MM.yyyy HH:mm}";
            return "Last modified: Unknown";
        }
    }

    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    public bool IsEnabled
    {
        get => _isEnabled;
        set { _isEnabled = value; OnPropertyChanged(); }
    }

    public NotebookViewModel(NotebookInfo info) => Info = info;

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
