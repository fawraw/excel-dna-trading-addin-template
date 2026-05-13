using Microsoft.Extensions.Logging;

namespace TradingAddin.Logging;

public static class AddInLogging
{
    private static readonly ILoggerFactory _factory = LoggerFactory.Create(builder =>
    {
        builder.SetMinimumLevel(LogLevel.Information);
        builder.AddConsole();
    });

    public static ILogger<T> CreateLogger<T>() => _factory.CreateLogger<T>();
}
