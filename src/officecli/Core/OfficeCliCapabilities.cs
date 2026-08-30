using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OfficeCli.Core;

internal sealed record RendererCapability(bool Available, string? Reason = null, string[]? Backends = null);
internal sealed record OfficeCliCapabilityReport(
    int SchemaVersion,
    string OfficeCliVersion,
    string Os,
    string Architecture,
    string RuntimeIdentifier,
    IReadOnlyDictionary<string, RendererCapability> Renderers,
    string[] ValidationProfiles);

internal static class OfficeCliCapabilities
{
    internal static OfficeCliCapabilityReport Detect()
    {
        var version = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0";
        var browsers = HtmlScreenshot.AvailableBackends();
        var windows = OperatingSystem.IsWindows();
        var word = windows && IsComClassRegisteredWindows("000209FF-0000-0000-C000-000000000046");
        var powerpoint = windows && IsComClassRegisteredWindows("91493441-5A91-11CF-8700-00AA0060263B");
        return new(
            1,
            version,
            RuntimeInformation.OSDescription,
            RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
            RuntimeInformation.RuntimeIdentifier,
            new Dictionary<string, RendererCapability>(StringComparer.Ordinal)
            {
                ["nativePowerPoint"] = new(powerpoint, powerpoint ? null : windows ? "Microsoft PowerPoint COM registration was not found." : "Available only on Windows with Microsoft PowerPoint installed."),
                ["nativeWord"] = new(word, word ? null : windows ? "Microsoft Word COM registration was not found." : "Available only on Windows with Microsoft Word installed."),
                ["htmlScreenshot"] = new(browsers.Length > 0, browsers.Length > 0 ? null : "Install Playwright, Chrome/Edge/Chromium, or Firefox.", browsers),
                ["mermaidImage"] = new(Diagram.MermaidImageRenderer.IsAvailable(), Diagram.MermaidImageRenderer.IsAvailable() ? null : "Install mermaid-cli or a Chrome-family browser."),
            },
            ["schema", "strict-opc", "ios-preview"]);
    }

    [SupportedOSPlatform("windows")]
    private static bool IsComClassRegisteredWindows(string clsid)
    {
        try
        {
            using var key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey($@"CLSID\{{{clsid}}}\LocalServer32");
            return key?.GetValue(null) is string value && !string.IsNullOrWhiteSpace(value);
        }
        catch { return false; }
    }
}
