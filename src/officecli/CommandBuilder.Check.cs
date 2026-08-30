// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildValidateCommand(Option<bool> jsonOption)
    {
        var validateFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var validateCommand = new Command("validate", "Validate document against OpenXML schema");
        var profileOption = new Option<string>("--profile") { Description = "Validation profile: schema, strict-opc, or ios-preview", DefaultValueFactory = _ => "schema" };
        validateCommand.Add(validateFileArg);
        validateCommand.Add(jsonOption);
        validateCommand.Add(profileOption);
        validateCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(validateFileArg)!;
            var profile = NormalizeValidationProfile(result.GetValue(profileOption));

            if (TryResident(file.FullName, req =>
            {
                req.Command = "validate";
                req.Json = json;
                req.Args["profile"] = profile;
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var report = RunValidation(handler, profile);
            var errors = report.Errors;
            if (json)
            {
                var validationJson = FormatValidationReport(report, profile);
                // JSON Envelope contract: validate is a *judgment* command —
                // schema errors mean the document failed validation, so the
                // envelope must reflect that on success. exit code already
                // mirrors this at line below.
                Console.WriteLine(OutputFormatter.WrapEnvelope(validationJson, success: errors.Count == 0));
            }
            else
            {
                if (errors.Count == 0)
                {
                    Console.WriteLine($"Validation passed ({profile}): no errors found." + (report.Warnings.Count > 0 ? $" {report.Warnings.Count} warning(s)." : ""));
                }
                else
                {
                    // R7-bt-4: schema validation reports go to stderr —
                    // callers piping `validate` for CI gates need to see
                    // the failure summary on the diagnostic stream rather
                    // than mixed into stdout. Mirrors the resident path.
                    Console.Error.WriteLine($"Found {errors.Count} validation error(s):");
                    foreach (var err in errors)
                    {
                        Console.Error.WriteLine($"  [{err.ErrorType}] {err.Description}");
                        if (err.Path != null) Console.Error.WriteLine($"    Path: {err.Path}");
                        if (err.Part != null) Console.Error.WriteLine($"    Part: {err.Part}");
                    }
                }
            }
            return errors.Count > 0 ? 1 : 0;
        }, json); });

        return validateCommand;
    }

    internal static string NormalizeValidationProfile(string? profile)
    {
        profile = string.IsNullOrWhiteSpace(profile) ? "schema" : profile.Trim().ToLowerInvariant();
        if (profile is not ("schema" or "strict-opc" or "ios-preview"))
            throw new CliException($"Unknown validation profile '{profile}'. Supported: schema, strict-opc, ios-preview.") { Code = "invalid_value" };
        return profile;
    }

    internal static PptxCompatibilityResult RunValidation(IDocumentHandler handler, string profile)
    {
        var schemaErrors = handler.Validate();
        if (profile == "schema") return new(schemaErrors, [], new(0, 0, 0, 0));
        if (handler is not PowerPointHandler ppt)
            throw new CliException($"Validation profile '{profile}' is currently supported only for PPTX files.") { Code = "invalid_value" };
        var snapshot = ppt.CreateValidationSnapshot();
        try
        {
            // An editable resident snapshot is the pre-commit package image.
            // Apply the same temp-file postprocessor Save/Dispose will apply so
            // validation predicts the committed artifact without touching disk.
            if (ppt.ValidationSnapshotNeedsSaveConformance)
                PptxPackageConformance.NormalizeInternalRelationshipTargets(snapshot);
            var compatibility = PptxCompatibilityValidator.Validate(snapshot, profile);
            compatibility.Errors.InsertRange(0, schemaErrors);
            return compatibility;
        }
        finally { try { File.Delete(snapshot); } catch { } }
    }

    internal static string FormatValidationReport(PptxCompatibilityResult report, string profile)
    {
        var errors = System.Text.Json.Nodes.JsonNode.Parse(FormatValidationErrors(report.Errors))!.AsObject();
        var root = new System.Text.Json.Nodes.JsonObject
        {
            ["schemaVersion"] = 1, ["profile"] = profile, ["count"] = report.Errors.Count,
            ["errors"] = errors["errors"]!.DeepClone(),
            ["warnings"] = new System.Text.Json.Nodes.JsonArray(report.Warnings.Select(w => (System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonObject { ["code"] = w.Code, ["message"] = w.Message, ["part"] = w.Part }).ToArray()),
            ["metrics"] = new System.Text.Json.Nodes.JsonObject { ["slides"] = report.Metrics.Slides, ["notesSlides"] = report.Metrics.NotesSlides, ["maxImageWidth"] = report.Metrics.MaxImageWidth, ["maxImageHeight"] = report.Metrics.MaxImageHeight }
        };
        return root.ToJsonString();
    }
}
