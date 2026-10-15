using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>Where a document's identity values were found.</summary>
[Flags]
public enum IdentitySources
{
    /// <summary>The document carries no identity; everything came from the metadata seed.</summary>
    None = 0,
    /// <summary>The template's bound report store (<see cref="TemplateManifest.DataStore"/>).</summary>
    DataStore = 1,
    /// <summary>docProps/custom.xml properties the manifest binds to metadata keys.</summary>
    CustomProperties = 2,
    /// <summary>docProps/core.xml properties the manifest binds to metadata keys.</summary>
    CoreProperties = 4,
}

/// <summary>The identity a document records about itself.</summary>
/// <param name="Values">Metadata keys and values found, in discovery order.</param>
/// <param name="Sources">The parts that contributed at least one value.</param>
public sealed record DocumentIdentity(ReportMetadata Values, IdentitySources Sources)
{
    /// <summary>
    /// The subset of <paramref name="keys"/> (in the order given) for which the document has no
    /// non-blank value of its own, i.e. the values a metadata seed would have to supply.
    /// </summary>
    /// <example><c>identity.Missing(manifest.IdentityKeys)</c> or <c>identity.Missing("title", "projectCode")</c>.</example>
    public IReadOnlyList<string> Missing(params IEnumerable<string> keys) => [.. keys.Where(k => Values.Get(k) is null).Distinct()];
}

/// <summary>
/// Recovers a document's own identity as the template's manifest defines it, most authoritative
/// source first: the bound report store -&gt; docProps/custom.xml -&gt; docProps/core.xml (only
/// properties bound to a metadata key count). The first non-blank value for a key wins.
/// </summary>
/// <remarks>
/// Cover text is deliberately NOT parsed: a legacy document's title exists only as typed
/// characters, and guessing which paragraph is the title is exactly the kind of fragile
/// inference this toolkit is supposed to replace. Those documents are seeded once from
/// report metadata and carry their identity themselves afterwards.
/// </remarks>
public static class IdentityReader
{
    public static DocumentIdentity Read(WordprocessingDocument doc, TemplateManifest manifest)
    {
        var found = new List<KeyValuePair<string, string>>();
        var sources = IdentitySources.None;
        void Take(string key, string? value, IdentitySources source)
        {
            if (string.IsNullOrWhiteSpace(value) || found.Exists(f => f.Key == key)) return;
            found.Add(new(key, value));
            sources |= source;
        }

        foreach (var part in doc.MainDocumentPart?.CustomXmlParts ?? [])
            if (DataStoreWriter.ReportRoot(part, manifest.DataStore) is { } report)
                foreach (var el in report.Elements())
                    Take(el.Name.LocalName, el.Value, IdentitySources.DataStore);

        foreach (var prop in doc.CustomFilePropertiesPart?.Properties?.Elements<CustomDocumentProperty>() ?? [])
        {
            var key = manifest.CustomProperties.FirstOrDefault(c => c.Key is not null && c.Name == prop.Name?.Value)?.Key;
            if (key is not null) Take(key, prop.VTLPWSTR?.Text, IdentitySources.CustomProperties);
        }

        if (doc.CoreFilePropertiesPart is { } core && TryLoad(core) is { } cx)
            foreach (var binding in manifest.CoreProperties.Where(b => b.Key is not null))
                Take(binding.Key!, cx.Descendants(TemplateManifest.CorePropertyElement(binding.Name)).FirstOrDefault()?.Value,
                    IdentitySources.CoreProperties);
        return new DocumentIdentity(ReportMetadata.From(found), sources);
    }

    private static XDocument? TryLoad(OpenXmlPart part)
    {
        try
        {
            return Docx.Load(part);
        }
        catch (XmlException)
        {
            return null;
        }
    }
}
