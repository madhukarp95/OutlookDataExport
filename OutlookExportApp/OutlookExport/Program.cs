using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using OutlookExport;
using OutlookExport.Services;
using Serilog;
using Microsoft.Extensions.Logging;

IConfiguration config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .Build();

using IHost host = Host.CreateDefaultBuilder()
    .ConfigureServices(services => services.AddSingleton<ProcessReport>())
    .ConfigureServices(services => services.AddSingleton<InboxService>())
    .ConfigureServices(services => services.AddSingleton<CalendarService>())
    .ConfigureServices(services => services.AddSingleton<SentItemService>())
    .ConfigureServices((context, services) =>
    {
        var configurationRoot = context.Configuration;
        services.Configure<ConfigOptions>(
            configurationRoot.GetSection(nameof(ConfigOptions)));
        services.Configure<FolderCount>(
                    configurationRoot.GetSection(nameof(FolderCount)));
    })
    .ConfigureLogging(builder => builder.ClearProviders())
    .UseSerilog((ctx, cfg) => cfg.ReadFrom.Configuration(ctx.Configuration))
    .Build();

var process = host.Services.GetService<ProcessReport>();
if (process != null)
    process.Download_data();
