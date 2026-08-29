using System.IO.Compression;
using System.Xml.Linq;

namespace OfficeCli.Handlers;

internal static class PptxPackageConformance
{
    private static readonly XNamespace RelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// Rewrites absolute-path internal OPC relationship targets to paths relative
    /// to the relationship owner. External targets are left untouched.
    /// </summary>
    internal static void NormalizeInternalRelationshipTargets(string packagePath)
    {
        using (var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update))
        {
            foreach (var entry in archive.Entries
                         .Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                         .ToList())
            {
                XDocument document;
                using (var input = entry.Open())
                    document = XDocument.Load(input);

                var changed = false;
                var ownerDirectory = GetOwnerDirectory(entry.FullName);
                foreach (var relationship in document.Root?.Elements(RelationshipsNamespace + "Relationship")
                             ?? Enumerable.Empty<XElement>())
                {
                    if (string.Equals((string?)relationship.Attribute("TargetMode"), "External",
                            StringComparison.OrdinalIgnoreCase))
                        continue;

                    var targetAttribute = relationship.Attribute("Target");
                    var target = targetAttribute?.Value;
                    if (string.IsNullOrEmpty(target) || !target.StartsWith("/", StringComparison.Ordinal))
                        continue;

                    var ownerUri = new Uri($"http://package/{ownerDirectory}", UriKind.Absolute);
                    var targetUri = new Uri($"http://package/{target.TrimStart('/')}", UriKind.Absolute);
                    targetAttribute!.Value = Uri.UnescapeDataString(ownerUri.MakeRelativeUri(targetUri).ToString());
                    changed = true;
                }

                if (!changed) continue;
                using var output = entry.Open();
                output.SetLength(0);
                document.Save(output, SaveOptions.DisableFormatting);
            }
        }

        ReorderForStreamingReaders(packagePath);
    }

    /// <summary>
    /// Places the content-type manifest and root relationships before all parts.
    /// OPC does not require this physical order, but lightweight streaming
    /// previewers (notably some iOS integrations) are more reliable with it.
    /// The package being rewritten is already AtomicPackageWriter's temporary
    /// file, so a repack failure cannot damage the user's original document.
    /// </summary>
    private static void ReorderForStreamingReaders(string packagePath)
    {
        var repackedPath = packagePath + $".repack-{Guid.NewGuid():N}";
        try
        {
            using (var source = ZipFile.OpenRead(packagePath))
            using (var destination = ZipFile.Open(repackedPath, ZipArchiveMode.Create))
            {
                var ordered = new List<ZipArchiveEntry>();
                AddIfPresent("[Content_Types].xml");
                AddIfPresent("_rels/.rels");
                ordered.AddRange(source.Entries.Where(e =>
                    !string.Equals(e.FullName, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(e.FullName, "_rels/.rels", StringComparison.OrdinalIgnoreCase)));

                foreach (var sourceEntry in ordered)
                {
                    var destinationEntry = destination.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
                    destinationEntry.LastWriteTime = sourceEntry.LastWriteTime;
                    destinationEntry.ExternalAttributes = sourceEntry.ExternalAttributes;
                    using var input = sourceEntry.Open();
                    using var output = destinationEntry.Open();
                    input.CopyTo(output);
                }

                void AddIfPresent(string name)
                {
                    var entry = source.Entries.FirstOrDefault(e =>
                        string.Equals(e.FullName, name, StringComparison.OrdinalIgnoreCase));
                    if (entry != null) ordered.Add(entry);
                }
            }

            File.Move(repackedPath, packagePath, overwrite: true);
        }
        finally
        {
            if (File.Exists(repackedPath)) File.Delete(repackedPath);
        }
    }

    private static string GetOwnerDirectory(string relationshipEntry)
    {
        if (string.Equals(relationshipEntry, "_rels/.rels", StringComparison.OrdinalIgnoreCase))
            return string.Empty;

        const string marker = "/_rels/";
        var markerIndex = relationshipEntry.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase);
        return markerIndex < 0 ? string.Empty : relationshipEntry[..(markerIndex + 1)];
    }
}
