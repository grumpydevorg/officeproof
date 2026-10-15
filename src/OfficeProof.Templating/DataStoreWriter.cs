using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>
/// Writes the custom-XML data store (<c>&lt;prefix:report&gt;</c> in the template's
/// <see cref="DataStoreSpec"/> namespace) that the cover and header content controls bind to
/// (<c>w:dataBinding/@w:storeItemID</c>). Editing a bound control in Word updates the store
/// and every other control bound to the same element.
/// </summary>
public static class DataStoreWriter
{
    private static readonly XNamespace Ds = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

    /// <summary>
    /// Writes <paramref name="metadata"/> as the document's one report store and returns its
    /// item id ("{GUID}"), the value bound controls carry in <c>w:storeItemID</c>.
    /// </summary>
    /// <remarks>
    /// The legacy CLI added a store on every run, so a retemplated document that already had one
    /// carried two, and its existing controls stayed bound to the stale copy. Now any existing
    /// report store is replaced and its item id is reused, so controls already bound to it see
    /// the new values; a document without one gets a fresh id.
    /// </remarks>
    public static string Write(MainDocumentPart main, ReportMetadata metadata, DataStoreSpec store)
    {
        var existing = main.CustomXmlParts.Where(p => ReportRoot(p, store) is not null).ToList();
        var storeId = existing.Select(ItemId).FirstOrDefault(id => id is not null) ?? PartTemplate.Braced(Guid.NewGuid());
        foreach (var old in existing) main.DeletePart(old);

        XNamespace ns = store.Namespace;
        var part = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        Docx.Save(part, new XDocument(new XElement(store.Root,
            new XAttribute(XNamespace.Xmlns + store.Prefix, store.Namespace),
            metadata.Entries.Select(e => new XElement(ns + e.Key, e.Value)))));

        var props = part.AddNewPart<CustomXmlPropertiesPart>();
        Docx.Save(props, new XDocument(new XElement(Ds + "datastoreItem",
            new XAttribute(Ds + "itemID", storeId),
            new XAttribute(XNamespace.Xmlns + "ds", Ds.NamespaceName),
            new XElement(Ds + "schemaRefs", new XElement(Ds + "schemaRef", new XAttribute(Ds + "uri", store.Namespace))))));
        return storeId;
    }

    /// <summary>The root of the part's XML when it is a report store of <paramref name="store"/>; null otherwise, including unparseable parts.</summary>
    internal static XElement? ReportRoot(CustomXmlPart part, DataStoreSpec store)
    {
        try
        {
            var root = Docx.Load(part).Root;
            return root?.Name == store.Root ? root : null;
        }
        catch (XmlException)
        {
            return null;
        }
    }

    private static string? ItemId(CustomXmlPart part)
    {
        if (part.CustomXmlPropertiesPart is not { } props) return null;
        try
        {
            return (string?)Docx.Load(props).Root?.Attribute(Ds + "itemID");
        }
        catch (XmlException)
        {
            return null;
        }
    }
}
