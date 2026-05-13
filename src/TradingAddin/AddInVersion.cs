using System.Reflection;

namespace TradingAddin;

public static class AddInVersion
{
    public static string Get() =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0";
}
