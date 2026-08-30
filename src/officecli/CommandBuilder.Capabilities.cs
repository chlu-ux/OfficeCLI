using System.CommandLine;
using System.Text.Json.Nodes;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildCapabilitiesCommand(Option<bool> jsonOption)
    {
        var command = new Command("capabilities", "Report platform, runtime, renderer, and validation-profile availability.");
        command.Add(jsonOption);
        command.SetAction(result => SafeRun(() =>
        {
            var json = result.GetValue(jsonOption);
            var report = OfficeCliCapabilities.Detect();
            if (json)
            {
                var renderers = new JsonObject();
                foreach (var (name, capability) in report.Renderers)
                    renderers[name] = new JsonObject
                    {
                        ["available"] = capability.Available,
                        ["reason"] = capability.Reason,
                        ["backends"] = capability.Backends == null ? null : new JsonArray(capability.Backends.Select(s => (JsonNode?)JsonValue.Create(s)).ToArray())
                    };
                var body = new JsonObject
                {
                    ["schemaVersion"] = report.SchemaVersion,
                    ["officeCliVersion"] = report.OfficeCliVersion,
                    ["os"] = report.Os,
                    ["architecture"] = report.Architecture,
                    ["runtimeIdentifier"] = report.RuntimeIdentifier,
                    ["renderers"] = renderers,
                    ["validationProfiles"] = new JsonArray(report.ValidationProfiles.Select(s => (JsonNode?)JsonValue.Create(s)).ToArray())
                };
                Console.WriteLine(OutputFormatter.WrapEnvelope(body.ToJsonString()));
            }
            else
            {
                Console.WriteLine($"OfficeCLI {report.OfficeCliVersion}");
                Console.WriteLine($"OS: {report.Os}");
                Console.WriteLine($"Architecture: {report.Architecture}");
                Console.WriteLine($"RID: {report.RuntimeIdentifier}");
                Console.WriteLine("Renderers:");
                foreach (var (name, capability) in report.Renderers)
                    Console.WriteLine($"  {name}: {(capability.Available ? "available" : "unavailable")}" + (capability.Backends is { Length: > 0 } ? $" ({string.Join(", ", capability.Backends)})" : "") + (!capability.Available && capability.Reason != null ? $" — {capability.Reason}" : ""));
                Console.WriteLine($"Validation profiles: {string.Join(", ", report.ValidationProfiles)}");
            }
            return 0;
        }, result.GetValue(jsonOption)));
        return command;
    }
}
