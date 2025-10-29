## Icon Configuration Notes

**Important for future reference:**

The application icon (`appicon.ico`) must be embedded as a WPF **Resource**, not just copied to output.

### Correct .csproj Configuration:
```xml
<ItemGroup>
  <Resource Include="appicon.ico" />
</ItemGroup>
```

### ❌ Incorrect (causes "Cannot locate resource" crash):
```xml
<ItemGroup>
  <None Update="appicon.ico">
    <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
  </None>
</ItemGroup>
```

### XAML Reference:
- For MainWindow: Remove `Icon="appicon.ico"` if it causes issues during development
- The icon in `.csproj` `<ApplicationIcon>appicon.ico</ApplicationIcon>` sets the EXE icon (taskbar/file explorer)
- The XAML `Icon` property requires the resource to be embedded

### Why This Matters:
- WPF looks for icons in the compiled assembly resources
- `<Resource Include>` embeds the file into the DLL/EXE
- `<None Update>` only copies to output directory (file system)
- Single-file publish bundles resources but not loose files