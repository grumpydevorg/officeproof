using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>
/// Writes report metadata into the package's OWN property parts (docProps/core.xml and
/// docProps/custom.xml), not only the private report store, as the template's
/// <see cref="TemplateManifest.CoreProperties"/> and <see cref="TemplateManifest.CustomProperties"/> say.
/// </summary>
/// <remarks>
/// Until 2026-08-27 <c>generate</c> emitted no docProps/core.xml at all -- no part, no content
/// type, no relationship -- so every document it produced had an empty Title and Author in
/// Word's File &gt; Properties, and was invisible to Spotlight, Drive and any DMS. It also meant
/// a cover control bound to the core store rendered blank, because the store did not exist.
/// <para>
/// Title/Subject/Creator have native homes in core.xml; metadata with no Word built-in goes to
/// docProps/custom.xml where File &gt; Properties &gt; Custom can reach it.
/// </para>
/// </remarks>
public static class DocPropsWriter
{
    // Word's format id for user-defined document properties.
    private const string UserDefinedFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTerms = "http://purl.org/dc/terms/";
    private static readonly XNamespace DcmiType = "http://purl.org/dc/dcmitype/";
    private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";

    /// <summary>Sets the manifest's core and custom properties from <paramref name="metadata"/>.</summary>
    /// <param name="doc">The package to write into; existing property parts are reused.</param>
    /// <param name="metadata">The values to write.</param>
    /// <param name="manifest">Which property takes which metadata key (or fixed text).</param>
    /// <param name="newDocument">
    /// True when officeproof created the package (generate): also stamp cp:lastModifiedBy (the
    /// creator) and cp:revision 1. False for an existing document (retemplate): its edit
    /// history is left exactly as Word recorded it.
    /// </param>
    public static void Write(WordprocessingDocument doc, ReportMetadata metadata, TemplateManifest manifest, bool newDocument = true)
    {
        WriteCore(doc, metadata, manifest.CoreProperties, newDocument);
        WriteCustom(doc, metadata, manifest.CustomProperties);
    }

    private static void WriteCore(WordprocessingDocument doc, ReportMetadata metadata, IReadOnlyList<PropertyBinding> bindings, bool newDocument)
    {
        // doc.PackageProperties setters do NOT persist a core.xml part here (verified
        // 2026-08-27: after setting Title/Subject/Creator/Category the package still had 24
        // parts and no docProps/core.xml). Write the part explicitly instead -- the SDK relates
        // and content-types it for us.
        //
        // REUSE an existing part, do not add a second one. On the retemplate path the input
        // already HAS these parts, and AddNewPart silently keeps the existing one rather than
        // replacing it, so a retemplated report kept its original empty title and account-name
        // dc:creator while reporting success. Deleting first works but costs the canonical
        // part name -- the SDK then writes docProps/core2.xml, which is legal (parts are
        // addressed by relationship, not path) but breaks any naive consumer that looks for
        // docProps/core.xml.
        var existing = doc.CoreFilePropertiesPart;
        var core = existing ?? doc.AddCoreFilePropertiesPart();

        // Keep what the document already records (created/modified dates, keywords) and
        // overwrite only the properties the metadata owns.
        var xml = existing is null ? null : TryLoad(core);
        var root = xml?.Root is { } r && r.Name == Cp + "coreProperties" ? r : NewCoreRoot();
        foreach (var (prefix, ns) in new[] { ("cp", Cp), ("dc", Dc), ("dcterms", DcTerms), ("dcmitype", DcmiType), ("xsi", Xsi) })
            if (root.GetPrefixOfNamespace(ns) is null)
                root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);

        string? creator = null;
        foreach (var binding in bindings)
        {
            var value = binding.Resolve(metadata) ?? "";
            Set(root, TemplateManifest.CorePropertyElement(binding.Name), value);
            if (binding.Name == "creator") creator = value;
        }
        // cp:revision counts Word's saves and cp:lastModifiedBy names the last person to save;
        // both belong to the document's history, not to its identity. Until 2026-10-15
        // retemplate reset a report's revision 406 to 1 and replaced its last author with the
        // metadata author (review finding 6), so only a document officeproof creates gets them.
        if (newDocument)
        {
            if (creator is not null) Set(root, Cp + "lastModifiedBy", creator);
            Set(root, Cp + "revision", "1");
        }
        Docx.Save(core, new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root));
    }

    private static void WriteCustom(WordprocessingDocument doc, ReportMetadata metadata, IReadOnlyList<PropertyBinding> bindings)
    {
        var custom = bindings
            .Select(b => (Property: b.Name, Value: b.Resolve(metadata)))
            .Where(c => c.Value is not null)
            .ToList();
        if (custom.Count == 0) return;

        // Same reuse rule as core.xml. The document's other custom properties survive; ours are
        // replaced. Property ids are renumbered because they must be unique: ids 0 and 1 are
        // reserved, so the first user property is 2.
        var existing = doc.CustomFilePropertiesPart;
        var names = custom.Select(c => c.Property).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var kept = (existing?.Properties?.Elements<CustomDocumentProperty>() ?? [])
            .Where(p => p.Name?.Value is not { } n || !names.Contains(n))
            .Select(p => (CustomDocumentProperty)p.CloneNode(true))
            .ToList();
        var part = existing ?? doc.AddCustomFilePropertiesPart();

        var container = new Properties();
        foreach (var p in kept) container.AppendChild(p);
        foreach (var (name, value) in custom)
            container.AppendChild(new CustomDocumentProperty
            {
                FormatId = UserDefinedFormatId,
                Name = name,
                VTLPWSTR = new VTLPWSTR(value!),
            });
        var pid = 2;
        foreach (var p in container.Elements<CustomDocumentProperty>()) p.PropertyId = pid++;

        // Assign through the typed root, not Properties.Save(part): the SDK re-saves a loaded root
        // when the package closes, which would overwrite a part written behind its back.
        part.Properties = container;
        part.Properties.Save();
    }

    private static XElement NewCoreRoot() => new(Cp + "coreProperties");

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

    private static void Set(XElement root, XName name, string value)
    {
        if (root.Element(name) is { } el) el.Value = value;
        else root.Add(new XElement(name, value));
    }
}
