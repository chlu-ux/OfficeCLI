using System.IO.Compression;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

internal sealed record PptxCompatibilityWarning(string Code, string Message, string? Part = null);
internal sealed record PptxCompatibilityMetrics(int Slides, int NotesSlides, int MaxImageWidth, int MaxImageHeight);
internal sealed record PptxCompatibilityResult(List<ValidationError> Errors, List<PptxCompatibilityWarning> Warnings, PptxCompatibilityMetrics Metrics);

internal static class PptxCompatibilityValidator
{
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private const string SlideRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    private const string NotesSlideRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    private const string NotesMasterRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

    internal static PptxCompatibilityResult Validate(string path, string profile)
    {
        var errors = new List<ValidationError>(); var warnings = new List<PptxCompatibilityWarning>();
        using var zip = ZipFile.OpenRead(path);
        var entries = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var slides = entries.Count(IsSlide); var notes = entries.Count(IsNotesSlide);
        ValidateManifest(zip, entries, errors);
        ValidateRelationships(zip, entries, errors);
        if (profile == "ios-preview")
        {
            ValidateNotes(zip, entries, errors);
            ValidateEmptyRuns(zip, warnings);
            ValidateAppProperties(zip, slides, notes, warnings);
            ValidateEntryOrder(zip, warnings);
        }
        var (maxWidth, maxHeight) = ReadImageDimensions(zip);
        if (profile == "ios-preview" && (maxWidth > 1920 || maxHeight > 1920))
            warnings.Add(new("image_budget_exceeded", $"Largest raster image is {maxWidth}x{maxHeight}; mobile preview budget is 1920px."));
        return new(errors, warnings, new(slides, notes, maxWidth, maxHeight));
    }

    private static void ValidateManifest(ZipArchive zip, HashSet<string> entries, List<ValidationError> errors)
    {
        var manifestEntry = zip.GetEntry("[Content_Types].xml");
        if (manifestEntry == null) { Add(errors, "manifest_missing", "[Content_Types].xml is missing.", "/[Content_Types].xml"); return; }
        var doc = Load(manifestEntry); var root = doc.Root!;
        var defaults = root.Elements(Ct + "Default").Where(e => e.Attribute("Extension") != null).ToDictionary(e => ((string)e.Attribute("Extension")!).ToLowerInvariant(), e => (string?)e.Attribute("ContentType") ?? "");
        var overrides = root.Elements(Ct + "Override").Where(e => e.Attribute("PartName") != null).ToList();
        foreach (var o in overrides)
        {
            var name = ((string)o.Attribute("PartName")!).TrimStart('/');
            if (!entries.Contains(name)) Add(errors, "dangling_content_type_override", $"Override points to missing part '/{name}'.", "/[Content_Types].xml");
        }
        foreach (var part in entries.Where(e => e != "[Content_Types].xml" && !string.IsNullOrEmpty(Path.GetFileName(e))))
        {
            var hasOverride = overrides.Any(o => string.Equals(((string?)o.Attribute("PartName"))?.TrimStart('/'), part, StringComparison.OrdinalIgnoreCase));
            var ext = Path.GetExtension(part).TrimStart('.').ToLowerInvariant();
            if (!hasOverride && !defaults.ContainsKey(ext)) Add(errors, "content_type_missing", $"Part '/{part}' has no matching Default or Override.", "/[Content_Types].xml");
        }
    }

    private static void ValidateRelationships(ZipArchive zip, HashSet<string> entries, List<ValidationError> errors)
    {
        foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            var owner = Owner(entry.FullName);
            foreach (var rel in Load(entry).Root?.Elements(Rel + "Relationship") ?? [])
            {
                if (string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                var target = (string?)rel.Attribute("Target");
                if (string.IsNullOrEmpty(target)) { Add(errors, "relationship_target_missing", $"Relationship in {entry.FullName} has no Target.", "/" + entry.FullName); continue; }
                if (target.StartsWith('/')) Add(errors, "absolute_relationship_target", $"Internal Target '{target}' must be relative.", "/" + entry.FullName);
                var resolved = Resolve(owner, target);
                if (!entries.Contains(resolved)) Add(errors, "dangling_relationship", $"Relationship Target '{target}' resolves to missing part '/{resolved}'.", "/" + entry.FullName);
            }
        }
    }

    private static void ValidateNotes(ZipArchive zip, HashSet<string> entries, List<ValidationError> errors)
    {
        foreach (var notes in entries.Where(IsNotesSlide))
        {
            var relEntry = zip.GetEntry(RelationshipName(notes));
            if (relEntry == null) { Add(errors, "notes_relationships_missing", $"Notes part '/{notes}' has no relationships part.", "/" + notes); continue; }
            var types = Load(relEntry).Root?.Elements(Rel + "Relationship").Select(e => (string?)e.Attribute("Type")).ToList() ?? [];
            if (!types.Contains(SlideRel)) Add(errors, "notes_slide_link_missing", "Notes slide is not linked to its slide.", "/" + notes);
            if (!types.Contains(NotesMasterRel)) Add(errors, "notes_master_link_missing", "Notes slide is not linked to a notes master.", "/" + notes);
        }
        foreach (var slide in entries.Where(IsSlide))
        {
            var relEntry = zip.GetEntry(RelationshipName(slide)); if (relEntry == null) continue;
            foreach (var rel in Load(relEntry).Root?.Elements(Rel + "Relationship").Where(e => (string?)e.Attribute("Type") == NotesSlideRel) ?? [])
                if (!entries.Contains(Resolve("ppt/slides/", (string)rel.Attribute("Target")!))) Add(errors, "slide_notes_link_dangling", "Slide notes relationship is dangling.", "/" + slide);
        }
    }

    private static void ValidateEmptyRuns(ZipArchive zip, List<PptxCompatibilityWarning> warnings)
    {
        foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            if (Load(entry).Descendants(A + "r").Any(r => r.Elements().Any() && r.Elements().All(c => c.Name == A + "t") && r.Value.Length == 0))
                warnings.Add(new("semantically_empty_run", "Part contains an empty run with no formatting or visible content.", "/" + entry.FullName));
    }

    private static void ValidateAppProperties(ZipArchive zip, int slides, int notes, List<PptxCompatibilityWarning> warnings)
    {
        var entry = zip.GetEntry("docProps/app.xml"); if (entry == null) { warnings.Add(new("extended_properties_missing", "docProps/app.xml is missing.")); return; }
        var root = Load(entry).Root!;
        if ((int?)root.Element(Ep + "Slides") != slides || (int?)root.Element(Ep + "Notes") != notes)
            warnings.Add(new("extended_properties_stale", $"app.xml counts do not match package (Slides={slides}, Notes={notes}).", "/docProps/app.xml"));
        var titleCount = root.Element(Ep + "TitlesOfParts")?.Descendants().Count(e => e.Name.LocalName == "lpstr") ?? -1;
        if (titleCount != slides && !(slides == 0 && titleCount == -1)) warnings.Add(new("titles_of_parts_stale", $"TitlesOfParts contains {titleCount} entries; expected {slides}.", "/docProps/app.xml"));
    }

    private static void ValidateEntryOrder(ZipArchive zip, List<PptxCompatibilityWarning> warnings)
    {
        if (zip.Entries.Count == 0 || zip.Entries[0].FullName != "[Content_Types].xml") warnings.Add(new("streaming_entry_order", "[Content_Types].xml is not the first ZIP entry."));
        if (zip.Entries.Count < 2 || zip.Entries[1].FullName != "_rels/.rels") warnings.Add(new("streaming_entry_order", "_rels/.rels is not the second ZIP entry."));
    }

    private static (int Width, int Height) ReadImageDimensions(ZipArchive zip)
    {
        var maxW = 0; var maxH = 0;
        var header = new byte[24];
        foreach (var e in zip.Entries.Where(e => e.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase)))
        {
            Array.Clear(header); using var s = e.Open(); var n = s.Read(header); var h = header.AsSpan();
            int w = 0, ht = 0;
            if (n >= 24 && h[0] == 0x89 && h[1] == 0x50) { w = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(h[16..20]); ht = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(h[20..24]); }
            else if (n >= 10 && h[0] == 'G' && h[1] == 'I' && h[2] == 'F') { w = h[6] | h[7] << 8; ht = h[8] | h[9] << 8; }
            maxW = Math.Max(maxW, w); maxH = Math.Max(maxH, ht);
        }
        return (maxW, maxH);
    }

    private static bool IsSlide(string p) => p.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
    private static bool IsNotesSlide(string p) => p.StartsWith("ppt/notesSlides/notesSlide", StringComparison.OrdinalIgnoreCase) && p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
    private static void Add(List<ValidationError> e, string code, string message, string part) => e.Add(new(code, message, null, part));
    private static XDocument Load(ZipArchiveEntry e) { using var s = e.Open(); return XDocument.Load(s); }
    private static string RelationshipName(string part) { var i = part.LastIndexOf('/'); return part[..(i + 1)] + "_rels/" + part[(i + 1)..] + ".rels"; }
    private static string Owner(string rels) { if (rels == "_rels/.rels") return ""; var i = rels.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase); return i < 0 ? "" : rels[..(i + 1)]; }
    private static string Resolve(string owner, string target) => Uri.UnescapeDataString(new Uri(new Uri("http://package/" + owner), target).AbsolutePath.TrimStart('/'));
}
