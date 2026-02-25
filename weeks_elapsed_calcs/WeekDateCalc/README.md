# Week Date Wizard

Small WPF app for:
- week difference between two dates (absolute, exact days/7 rounded to 1 decimal)
- add/deduct weeks with workday adjustment rules
- row-by-row dynamic week differences (each row vs previous row)
- persistent searchable history

## Accepted date formats
- `DD/MM/YYYY`
- `DD.MM.YYYY`

## Weekend adjustment rule for add/deduct
- Positive weeks (`>= 0`): if result lands on weekend, move forward to Monday
- Negative weeks (`< 0`): if result lands on weekend, move back to Friday

## Build
```powershell
dotnet build -c Release
```

## Publish standalone EXE (x64)
```powershell
./publish_win_x64.ps1
```

Output:
- `bin/Release/net8.0-windows/win-x64/publish/WeekDateCalc.exe`

## History storage
- `%AppData%/WeekDateWizard/history.json`
