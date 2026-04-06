// Copyright (c) FileExplorer contributors. All rights reserved.

using FileExplorer.Core.Interfaces;
using FileExplorer.Infrastructure.Services;
using Microsoft.Extensions.DependencyInjection;

namespace FileExplorer.Infrastructure;

/// <summary>
/// DI registration extension for the Infrastructure layer.
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>Registers all Infrastructure services.</summary>
    public static IServiceCollection AddInfrastructureServices(this IServiceCollection services)
    {
        services.AddSingleton<IFileSystemService, FileSystemService>();
        services.AddSingleton<IShellIntegrationService, ShellIntegrationService>();
        services.AddSingleton<ICloudStatusService, CloudStatusService>();
        services.AddSingleton<ISettingsService, SettingsService>();
        services.AddSingleton<IPreviewService, PreviewService>();
        services.AddSingleton<IIdentityExtractionService, IdentityExtractionService>();
        services.AddSingleton<IIdentityCacheService, SqliteIdentityCacheService>();
        services.AddSingleton<IDeepIndexingService, DeepIndexingService>();
        services.AddSingleton<IRqUtilityService, RqUtilityService>();
        services.AddSingleton<ISearchService, RqSearchService>();
        return services;
    }
}
