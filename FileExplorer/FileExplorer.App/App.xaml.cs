// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Core.ViewModels;
using FileExplorer.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Serilog;
using System.Text;

namespace FileExplorer.App;

/// <summary>
/// Application entry point and DI composition root.
/// Configures all services, ViewModels, and global exception handling.
/// </summary>
public partial class App : Application
{
    private static IServiceProvider? _services;
    private static readonly string LogDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "FileExplorer",
        "logs");
    private static readonly string CrashDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "FileExplorer",
        "crash-reports");
    private static readonly string WerDumpDirectory = Path.Combine(CrashDirectory, "wer-dumps");

    internal static string? StartupPath { get; set; }

    /// <summary>Gets the DI service provider.</summary>
    public static IServiceProvider Services => _services ?? throw new InvalidOperationException("Services not initialized");

    /// <summary>Gets the main window instance.</summary>
    public static Window? MainWindow { get; private set; }

    /// <summary>Initializes a new instance of the <see cref="App"/> class.</summary>
    public App()
    {
        InitializeComponent();
        UnhandledException += App_UnhandledException;
        RegisterGlobalExceptionHandlers();
        ConfigureWerLocalDumps();

        // Configure DI
        var services = new ServiceCollection();
        ConfigureServices(services);
        _services = services.BuildServiceProvider();
    }

    /// <inheritdoc/>
    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        MainWindow = new MainWindow();
        MainWindow.Activate();
    }

    private static void ConfigureServices(IServiceCollection services)
    {
        // Logging
        Directory.CreateDirectory(LogDirectory);
        Directory.CreateDirectory(CrashDirectory);
        var logPath = Path.Combine(LogDirectory, "log-.txt");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(logPath,
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 14,
                outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {SourceContext} {Message:lj}{NewLine}{Exception}")
            .CreateLogger();

        services.AddLogging(builder =>
        {
            builder.AddSerilog(dispose: true);
        });

        // Infrastructure services
        services.AddInfrastructureServices();

        // ViewModels
        services.AddSingleton<MainViewModel>();
        services.AddSingleton<TreeViewModel>();

        // Tab factory: creates TabContentViewModel instances with DI
        services.AddTransient<TabContentViewModel>();
        services.AddSingleton<Func<TabState?, TabContentViewModel>>(sp => tabState =>
        {
            var vm = new TabContentViewModel(
                sp.GetRequiredService<IFileSystemService>(),
                sp.GetRequiredService<IShellIntegrationService>(),
                sp.GetRequiredService<ICloudStatusService>(),
                sp.GetRequiredService<IPreviewService>(),
                sp.GetRequiredService<IIdentityExtractionService>(),
                sp.GetRequiredService<IIdentityCacheService>(),
                sp.GetRequiredService<ITabPathStateService>(),
                sp.GetRequiredService<ISettingsService>(),
                sp.GetRequiredService<IDirectorySnapshotCache>(),
                sp.GetRequiredService<ILogger<TabContentViewModel>>(),
                tabState);
            return vm;
        });
    }

    private static void RegisterGlobalExceptionHandlers()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            ReportCrash(
                source: "AppDomain.CurrentDomain.UnhandledException",
                exception: args.ExceptionObject as Exception,
                isTerminating: args.IsTerminating,
                details: args.ExceptionObject is not Exception exObj ? args.ExceptionObject?.ToString() : null);
        };

        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            ReportCrash(
                source: "TaskScheduler.UnobservedTaskException",
                exception: args.Exception,
                isTerminating: false,
                details: "Unobserved task exception.");

            // Keep process alive after persisting diagnostics.
            args.SetObserved();
        };
    }

    private static void ConfigureWerLocalDumps()
    {
        try
        {
            Directory.CreateDirectory(WerDumpDirectory);

            using var key = Registry.CurrentUser.CreateSubKey(
                @"Software\Microsoft\Windows\Windows Error Reporting\LocalDumps\FileExplorer.exe",
                true);

            if (key is null)
                return;

            key.SetValue("DumpFolder", WerDumpDirectory, RegistryValueKind.ExpandString);
            key.SetValue("DumpType", 1, RegistryValueKind.DWord);
            key.SetValue("DumpCount", 10, RegistryValueKind.DWord);
        }
        catch
        {
            // Best-effort only. Managed crash report path still works without WER settings.
        }
    }

    private static void ReportCrash(string source, Exception? exception, bool isTerminating, string? details = null)
    {
        try
        {
            WriteCrashReport(source, exception, isTerminating, details);
        }
        catch
        {
            // Never throw while trying to capture crash diagnostics.
        }

        try
        {
            if (exception is not null)
            {
                Log.Fatal(exception, "Unhandled exception from {Source}. IsTerminating={IsTerminating}", source, isTerminating);
            }
            else
            {
                Log.Fatal("Unhandled exception from {Source}. IsTerminating={IsTerminating}. Details={Details}", source, isTerminating, details ?? "n/a");
            }
        }
        catch
        {
            // Ignore logging failures during crash processing.
        }

        try
        {
            Log.CloseAndFlush();
        }
        catch
        {
            // Ignore flush failures during crash processing.
        }
    }

    private static void WriteCrashReport(string source, Exception? exception, bool isTerminating, string? details)
    {
        Directory.CreateDirectory(CrashDirectory);

        var now = DateTimeOffset.Now;
        var fileName = $"crash-{now:yyyyMMdd-HHmmss-fff}-pid{Environment.ProcessId}.txt";
        var filePath = Path.Combine(CrashDirectory, fileName);

        var activeTabPath = TryGetActiveTabPath();
        var sb = new StringBuilder();
        sb.AppendLine("FileExplorer crash report");
        sb.AppendLine($"Timestamp: {now:O}");
        sb.AppendLine($"Source: {source}");
        sb.AppendLine($"IsTerminating: {isTerminating}");
        sb.AppendLine($"ProcessId: {Environment.ProcessId}");
        sb.AppendLine($"ProcessPath: {Environment.ProcessPath}");
        sb.AppendLine($"OS: {Environment.OSVersion}");
        sb.AppendLine($"Runtime: {Environment.Version}");
        sb.AppendLine($"StartupPathArg: {StartupPath ?? "(null)"}");
        sb.AppendLine($"CurrentDirectory: {Environment.CurrentDirectory}");
        sb.AppendLine($"ActiveTabPath: {activeTabPath ?? "(unknown)"}");

        if (!string.IsNullOrWhiteSpace(details))
        {
            sb.AppendLine();
            sb.AppendLine("Details:");
            sb.AppendLine(details);
        }

        if (exception is not null)
        {
            sb.AppendLine();
            sb.AppendLine("Exception:");
            sb.AppendLine(exception.ToString());
        }

        File.WriteAllText(filePath, sb.ToString());
    }

    private static string? TryGetActiveTabPath()
    {
        try
        {
            var mainVm = Services.GetService<MainViewModel>();
            return mainVm?.ActiveTab?.CurrentPath;
        }
        catch
        {
            return null;
        }
    }

    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        ReportCrash(
            source: "Application.UnhandledException",
            exception: e.Exception,
            isTerminating: false,
            details: e.Message);

        // Save session on crash for recovery
        try
        {
            var mainVm = Services.GetService<MainViewModel>();
            mainVm?.SaveSessionAsync().Wait(TimeSpan.FromSeconds(3));
        }
        catch
        {
            // Best-effort session save
        }

        e.Handled = true;
    }
}
