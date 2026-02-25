using System.ComponentModel;

namespace WeekDateCalc;

public sealed class SequenceRow : INotifyPropertyChanged
{
    private string _dateInput = string.Empty;
    private string _weekDifference = string.Empty;

    public string DateInput
    {
        get => _dateInput;
        set
        {
            if (_dateInput == value)
            {
                return;
            }

            _dateInput = value;
            OnPropertyChanged(nameof(DateInput));
        }
    }

    public string WeekDifference
    {
        get => _weekDifference;
        set
        {
            if (_weekDifference == value)
            {
                return;
            }

            _weekDifference = value;
            OnPropertyChanged(nameof(WeekDifference));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
