// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Interfaces;
using FileExplorer.Core.Models;
using FileExplorer.Core.ViewModels;
using FileExplorer.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Serilog;

namespace FileExplorer.App;

/// <summary>
/// Application entry point and DI composition root.
/// Configures all services, ViewModels, and global exception handling.
/// </summary>
public partial class App : Application
{
    private static IServiceProvider? _services;

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
        var logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "FileExplorer", "logs", "log-.txt");

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
                sp.GetRequiredService<ILogger<TabContentViewModel>>(),
                tabState);
            return vm;
        });
    }

    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        Log.Fatal(e.Exception, "Unhandled exception");
        Log.CloseAndFlush();

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
