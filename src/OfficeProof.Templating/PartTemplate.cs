using System.Security;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeProof.Core;

namespace OfficeProof.Templating;

/// <summary>
/// Fills a template part's <c>{{TOKEN}}</c> placeholders. Upper-case structural tokens
/// (<c>{{STOREID}}</c>, relationship ids) are substituted verbatim; every metadata key
/// <c>{{key}}</c> is substituted XML-escaped, since the values land in element text.
/// A token with no value stays in the output, as it always has.
/// </summary>
internal static class PartTemplate
{
    public static string Fill(string xml, ReportMetadata metadata, params (string Token, string Value)[] structural)
    {
        foreach (var (token, value) in structural)
            xml = xml.Replace("{{" + token + "}}", value, StringComparison.Ordinal);
        foreach (var (key, value) in metadata.Entries)
            xml = xml.Replace("{{" + key + "}}", SecurityElement.Escape(value), StringComparison.Ordinal);
        return xml;
    }

    /// <summary>Writes XML text as a part's content, UTF-8 without a byte-order mark.</summary>
    public static void WriteText(OpenXmlPart part, string xml) => Docx.Write(part, Encoding.UTF8.GetBytes(xml));

    /// <summary>"{GUID}" in the upper-case form Word writes for store item ids and font keys.</summary>
    public static string Braced(Guid guid) => "{" + guid.ToString("D").ToUpperInvariant() + "}";
}
