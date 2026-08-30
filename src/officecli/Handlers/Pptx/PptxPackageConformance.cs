using System.IO.Compression;
using System.Xml.Linq;

namespace OfficeCli.Handlers;

/// <summary>Idempotent, package-level PPTX compatibility repairs.</summary>
internal static class PptxPackageConformance
{
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Ep = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    private static readonly XNamespace Vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    private const string SlideRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
    private const string NotesSlideRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
    private const string NotesMasterRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";

    internal static void NormalizeInternalRelationshipTargets(string packagePath)
    {
        using (var zip = ZipFile.Open(packagePath, ZipArchiveMode.Update))
        {
            NormalizeRelationships(zip);
            RepairNotesGraph(zip);
            ReconcileContentTypes(zip);
            ReconcileExtendedProperties(zip);
            RemoveSemanticallyEmptyRuns(zip);
        }
        ReorderForStreamingReaders(packagePath);
    }

    private static void NormalizeRelationships(ZipArchive zip)
    {
        foreach (var entry in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            var doc = Load(entry); var changed = false; var owner = OwnerDirectory(entry.FullName);
            foreach (var relationship in doc.Root?.Elements(Rel + "Relationship") ?? [])
            {
                if (string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                var target = relationship.Attribute("Target");
                if (target?.Value.StartsWith('/') != true) continue;
                target.Value = RelativeTarget(owner, target.Value); changed = true;
            }
            if (changed) Save(entry, doc);
        }
    }

    private static void RepairNotesGraph(ZipArchive zip)
    {
        var names = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var master = names.Where(n => n.StartsWith("ppt/notesMasters/notesMaster", StringComparison.OrdinalIgnoreCase) && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).Order().FirstOrDefault();
        if (master != null)
        {
            var rels = ReadOrNew(zip, "ppt/_rels/presentation.xml.rels"); var changed = false;
            var rid = EnsureRelationship(rels, NotesMasterRel, RelativeTarget("ppt/", master), ref changed);
            if (changed) Write(zip, "ppt/_rels/presentation.xml.rels", rels);
            var presentationEntry = zip.GetEntry("ppt/presentation.xml");
            if (presentationEntry != null)
            {
                var presentation = Load(presentationEntry); var root = presentation.Root;
                if (root != null)
                {
                    var list = root.Element(P + "notesMasterIdLst");
                    if (list == null)
                    {
                        list = new XElement(P + "notesMasterIdLst");
                        var before = root.Elements().FirstOrDefault(e => e.Name == P + "handoutMasterIdLst" || e.Name == P + "sldIdLst");
                        if (before == null) root.Add(list); else before.AddBeforeSelf(list);
                    }
                    var ids = list.Elements(P + "notesMasterId").ToList();
                    if (ids.Count != 1 || (string?)ids[0].Attribute(R + "id") != rid)
                    { list.ReplaceNodes(new XElement(P + "notesMasterId", new XAttribute(R + "id", rid))); Save(presentationEntry, presentation); }
                }
            }
        }

        foreach (var notes in names.Where(n => n.StartsWith("ppt/notesSlides/notesSlide", StringComparison.OrdinalIgnoreCase) && n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var relName = RelationshipName(notes); var rels = ReadOrNew(zip, relName); var changed = false;
            if (master != null) EnsureRelationship(rels, NotesMasterRel, RelativeTarget("ppt/notesSlides/", master), ref changed);
            var slide = FindOwningSlide(zip, notes);
            if (slide != null)
            {
                EnsureRelationship(rels, SlideRel, RelativeTarget("ppt/notesSlides/", slide), ref changed);
                var slideRelsName = RelationshipName(slide); var slideRels = ReadOrNew(zip, slideRelsName); var slideChanged = false;
                EnsureRelationship(slideRels, NotesSlideRel, RelativeTarget("ppt/slides/", notes), ref slideChanged);
                if (slideChanged) Write(zip, slideRelsName, slideRels);
            }
            if (changed) Write(zip, relName, rels);
        }
    }

    private static string? FindOwningSlide(ZipArchive zip, string notes)
    {
        foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("ppt/slides/_rels/slide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
            if (Load(entry).Root?.Elements(Rel + "Relationship").Any(e => (string?)e.Attribute("Type") == NotesSlideRel && ResolveTarget("ppt/slides/", (string?)e.Attribute("Target")) == notes) == true)
                return entry.FullName.Replace("ppt/slides/_rels/", "ppt/slides/", StringComparison.OrdinalIgnoreCase)[..^5];
        // PowerPoint and the Open XML SDK allocate slideN/notesSlideN in lockstep.
        // Use that convention only as an unambiguous fallback when both links are gone.
        var number = Number(notes);
        var candidate = $"ppt/slides/slide{number}.xml";
        return number != int.MaxValue && zip.GetEntry(candidate) != null ? candidate : null;
    }

    private static void ReconcileContentTypes(ZipArchive zip)
    {
        var entry = zip.GetEntry("[Content_Types].xml") ?? throw new InvalidDataException("PPTX package has no [Content_Types].xml.");
        var doc = Load(entry); var root = doc.Root ?? throw new InvalidDataException("Invalid [Content_Types].xml.");
        var parts = zip.Entries.Where(e => !string.IsNullOrEmpty(e.Name) && e.FullName != "[Content_Types].xml").Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var e in root.Elements(Ct + "Override").Where(e => !parts.Contains(((string?)e.Attribute("PartName") ?? "").TrimStart('/'))).ToList()) e.Remove();
        EnsureDefault(root, "rels", "application/vnd.openxmlformats-package.relationships+xml");
        EnsureDefault(root, "xml", "application/xml");
        foreach (var part in parts)
        {
            var type = KnownContentType(part);
            if (type != null) EnsureOverride(root, "/" + part, type);
            else
            {
                var ext = Path.GetExtension(part).TrimStart('.');
                if (ext.Length != 0 && !root.Elements(Ct + "Default").Any(e => string.Equals((string?)e.Attribute("Extension"), ext, StringComparison.OrdinalIgnoreCase)))
                    root.Add(new XElement(Ct + "Default", new XAttribute("Extension", ext), new XAttribute("ContentType", Mime(ext))));
            }
        }
        Save(entry, doc);
    }

    private static void ReconcileExtendedProperties(ZipArchive zip)
    {
        var entry = zip.GetEntry("docProps/app.xml"); if (entry == null) return;
        var doc = Load(entry); var root = doc.Root; if (root == null) return;
        var slideEntries = zip.Entries.Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).OrderBy(e => Number(e.FullName)).ToList();
        var notes = zip.Entries.Count(e => e.FullName.StartsWith("ppt/notesSlides/notesSlide", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
        Set(root, Ep + "Slides", slideEntries.Count.ToString()); Set(root, Ep + "Notes", notes.ToString());
        var titles = slideEntries.Select(e =>
        {
            var xml = Load(e); var shape = xml.Descendants(P + "sp").FirstOrDefault(s => s.Descendants(P + "ph").Any(ph => new[] { "title", "ctrTitle", "subTitle" }.Contains((string?)ph.Attribute("type"))));
            return string.Concat((shape ?? xml.Root)?.Descendants(A + "t").Select(t => t.Value) ?? []).Trim();
        }).ToList();
        root.Element(Ep + "HeadingPairs")?.Remove(); root.Element(Ep + "TitlesOfParts")?.Remove();
        root.Add(new XElement(Ep + "HeadingPairs", new XElement(Vt + "vector", new XAttribute("size", 2), new XAttribute("baseType", "variant"), new XElement(Vt + "variant", new XElement(Vt + "lpstr", "Slides")), new XElement(Vt + "variant", new XElement(Vt + "i4", slideEntries.Count)))));
        if (titles.Count > 0)
            root.Add(new XElement(Ep + "TitlesOfParts", new XElement(Vt + "vector", new XAttribute("size", titles.Count), new XAttribute("baseType", "lpstr"), titles.Select(t => new XElement(Vt + "lpstr", t)))));
        Save(entry, doc);
    }

    private static void RemoveSemanticallyEmptyRuns(ZipArchive zip)
    {
        foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase) && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            var doc = Load(entry);
            var empty = doc.Descendants(A + "r").Where(r => r.Elements().Any() && r.Elements().All(c => c.Name == A + "t") && r.Elements(A + "t").All(t => t.Value.Length == 0)).ToList();
            if (empty.Count == 0) continue; empty.Remove(); Save(entry, doc);
        }
    }

    private static string EnsureRelationship(XDocument doc, string type, string target, ref bool changed)
    {
        var root = doc.Root!; var existing = root.Elements(Rel + "Relationship").FirstOrDefault(e => (string?)e.Attribute("Type") == type);
        if (existing != null) { if ((string?)existing.Attribute("Target") != target) { existing.SetAttributeValue("Target", target); changed = true; } return (string)existing.Attribute("Id")!; }
        var used = root.Elements(Rel + "Relationship").Select(e => (string?)e.Attribute("Id")).ToHashSet(); var n = 1; while (used.Contains("rId" + n)) n++;
        var id = "rId" + n; root.Add(new XElement(Rel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target))); changed = true; return id;
    }

    private static string? KnownContentType(string part) => part.ToLowerInvariant() switch
    {
        "ppt/presentation.xml" => "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "docprops/core.xml" => "application/vnd.openxmlformats-package.core-properties+xml",
        "docprops/app.xml" => "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        _ when part.StartsWith("ppt/theme/", StringComparison.OrdinalIgnoreCase) && part.EndsWith(".xml") => "application/vnd.openxmlformats-officedocument.theme+xml",
        _ when part.Contains("/theme/", StringComparison.OrdinalIgnoreCase) && part.EndsWith(".xml") => "application/vnd.openxmlformats-officedocument.theme+xml",
        _ when IsDirectXmlPart(part, "ppt/slides/") => "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        _ when IsDirectXmlPart(part, "ppt/notesSlides/") => "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
        _ when IsDirectXmlPart(part, "ppt/notesMasters/") => "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml",
        _ when IsDirectXmlPart(part, "ppt/slideMasters/") => "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
        _ when IsDirectXmlPart(part, "ppt/slideLayouts/") => "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
        _ => null
    };
    private static bool IsDirectXmlPart(string part, string directory) => part.StartsWith(directory, StringComparison.OrdinalIgnoreCase) && part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && !part[directory.Length..].Contains('/');
    private static string Mime(string ext) => ext.ToLowerInvariant() switch { "png" => "image/png", "jpg" or "jpeg" => "image/jpeg", "gif" => "image/gif", "svg" => "image/svg+xml", _ => "application/octet-stream" };
    private static void EnsureDefault(XElement root, string ext, string type) { if (!root.Elements(Ct + "Default").Any(e => string.Equals((string?)e.Attribute("Extension"), ext, StringComparison.OrdinalIgnoreCase))) root.AddFirst(new XElement(Ct + "Default", new XAttribute("Extension", ext), new XAttribute("ContentType", type))); }
    private static void EnsureOverride(XElement root, string name, string type) { var e = root.Elements(Ct + "Override").FirstOrDefault(x => string.Equals((string?)x.Attribute("PartName"), name, StringComparison.OrdinalIgnoreCase)); if (e == null) root.Add(new XElement(Ct + "Override", new XAttribute("PartName", name), new XAttribute("ContentType", type))); else e.SetAttributeValue("ContentType", type); }
    private static void Set(XElement root, XName name, string value) { var e = root.Element(name); if (e == null) root.Add(new XElement(name, value)); else e.Value = value; }
    private static int Number(string name) { var d = new string(Path.GetFileNameWithoutExtension(name).Where(char.IsDigit).ToArray()); return int.TryParse(d, out var n) ? n : int.MaxValue; }
    private static string RelationshipName(string part) { var i = part.LastIndexOf('/'); return part[..(i + 1)] + "_rels/" + part[(i + 1)..] + ".rels"; }
    private static string RelativeTarget(string owner, string target) => Uri.UnescapeDataString(new Uri("http://package/" + owner).MakeRelativeUri(new Uri("http://package/" + target.TrimStart('/'))).ToString());
    private static string? ResolveTarget(string owner, string? target) => string.IsNullOrEmpty(target) ? null : Uri.UnescapeDataString(new Uri(new Uri("http://package/" + owner), target).AbsolutePath.TrimStart('/'));
    private static XDocument Load(ZipArchiveEntry e) { using var s = e.Open(); return XDocument.Load(s); }
    private static void Save(ZipArchiveEntry e, XDocument d) { using var s = e.Open(); s.SetLength(0); d.Save(s, SaveOptions.DisableFormatting); }
    private static XDocument ReadOrNew(ZipArchive zip, string name) { var e = zip.GetEntry(name); return e == null ? new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), new XElement(Rel + "Relationships")) : Load(e); }
    private static void Write(ZipArchive zip, string name, XDocument doc) => Save(zip.GetEntry(name) ?? zip.CreateEntry(name, CompressionLevel.Optimal), doc);

    private static void ReorderForStreamingReaders(string path)
    {
        var temp = path + $".repack-{Guid.NewGuid():N}";
        try
        {
            using (var source = ZipFile.OpenRead(path)) using (var destination = ZipFile.Open(temp, ZipArchiveMode.Create))
            {
                var ordered = new List<ZipArchiveEntry>(); Add("[Content_Types].xml"); Add("_rels/.rels");
                ordered.AddRange(source.Entries.Where(e => !string.Equals(e.FullName, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase) && !string.Equals(e.FullName, "_rels/.rels", StringComparison.OrdinalIgnoreCase)));
                foreach (var e in ordered) { var d = destination.CreateEntry(e.FullName, CompressionLevel.Optimal); d.LastWriteTime = e.LastWriteTime; d.ExternalAttributes = e.ExternalAttributes; using var input = e.Open(); using var output = d.Open(); input.CopyTo(output); }
                void Add(string name) { var e = source.Entries.FirstOrDefault(x => string.Equals(x.FullName, name, StringComparison.OrdinalIgnoreCase)); if (e != null) ordered.Add(e); }
            }
            File.Move(temp, path, true);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }

    private static string OwnerDirectory(string rels)
    {
        if (string.Equals(rels, "_rels/.rels", StringComparison.OrdinalIgnoreCase)) return "";
        const string marker = "/_rels/"; var i = rels.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase); return i < 0 ? "" : rels[..(i + 1)];
    }
}
