using System.Security.Cryptography;
using System.Xml.Linq;
using OfficeProof.Core;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeProof.Verify;

/// <summary>
/// What the relationship references in story markup point at. A redline can change a hyperlink
/// URL, an image or an embedded object through the part's .rels or the target part alone, with
/// the story XML byte-identical; the canonical story form therefore carries the resolved target
/// (external URI, or a digest of the target part) instead of the relationship id, so a pure
/// renumbering of ids compares equal and any retarget does not.
/// </summary>
internal sealed class Relationships
{
    /// <summary>VML's o:relid (v:imagedata, OLE previews) is a relationship id outside the r: namespace.</summary>
    public static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";

    private readonly Dictionary<(OpenXmlPartContainer, string), string> _resolved = [];
    private readonly Dictionary<OpenXmlPart, string> _digests = [];

    /// <summary>True for an attribute whose value is a relationship id of the containing part.</summary>
    public static bool IsReference(XAttribute a) => a.Name.Namespace == Ns.R || a.Name == O + "relid";

    /// <summary>
    /// The target of relationship <paramref name="id"/> of <paramref name="owner"/>:
    /// <c>external &lt;uri&gt;</c>, <c>internal-link &lt;uri&gt;</c>, or
    /// <c>part &lt;content type&gt; sha256:&lt;digest&gt;</c> for a package part (its bytes plus,
    /// recursively, its own relationship targets). Unknown ids give <c>unresolved &lt;id&gt;</c>.
    /// </summary>
    public string Resolve(OpenXmlPartContainer owner, string id)
    {
        if (_resolved.TryGetValue((owner, id), out var known)) return known;
        var target = Uncached(owner, id);
        _resolved[(owner, id)] = target;
        return target;
    }

    /// <summary>The part behind relationship <paramref name="id"/> of <paramref name="owner"/>, or null.</summary>
    public static OpenXmlPart? Part(OpenXmlPartContainer owner, string id) =>
        owner.Parts.FirstOrDefault(p => p.RelationshipId == id).OpenXmlPart;

    private string Uncached(OpenXmlPartContainer owner, string id)
    {
        if (Part(owner, id) is { } part) return "part " + part.ContentType + " sha256:" + Digest(part);
        if (owner.HyperlinkRelationships.FirstOrDefault(r => r.Id == id) is { } link)
            return (link.IsExternal ? "external " : "internal-link ") + link.Uri.OriginalString;
        if (owner.ExternalRelationships.FirstOrDefault(r => r.Id == id) is { } external)
            return "external " + external.Uri.OriginalString;
        if (owner.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == id) is { } data)
            return "data " + data.DataPart.ContentType + " sha256:" + Hash(data.DataPart.GetStream());
        return "unresolved " + id;
    }

    private static IEnumerable<string> Ids(OpenXmlPartContainer owner) =>
        owner.Parts.Select(p => p.RelationshipId)
            .Concat(owner.HyperlinkRelationships.Select(r => r.Id))
            .Concat(owner.ExternalRelationships.Select(r => r.Id))
            .Concat(owner.DataPartReferenceRelationships.Select(r => r.Id));

    // The part's bytes plus, recursively, what each of its own relationships resolves to.
    private string Digest(OpenXmlPart part)
    {
        if (_digests.TryGetValue(part, out var digest)) return digest;
        _digests[part] = "(cycle)";
        var bytes = Hash(part.GetStream(FileMode.Open, FileAccess.Read));
        var children = string.Join("\n", Ids(part).Order(StringComparer.Ordinal).Select(id => id + "=" + Resolve(part, id)));
        digest = children.Length == 0
            ? bytes
            : Convert.ToHexStringLower(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(bytes + "\n" + children)));
        _digests[part] = digest;
        return digest;
    }

    private static string Hash(Stream stream)
    {
        using (stream) return Convert.ToHexStringLower(SHA256.HashData(stream));
    }

    /// <summary>
    /// Package part names (no leading slash) that the story comparison covers in a package:
    /// every story part and every part reached through a relationship referenced from story
    /// markup, recursively. A change to any other part is untracked and reported separately.
    /// </summary>
    public static IReadOnlySet<string> CoveredParts(byte[] package)
    {
        using var stream = new MemoryStream(package);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var covered = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (doc.MainDocumentPart is not { } main) return covered;
        foreach (var part in Docx.StoryParts(main))
        {
            covered.Add(Name(part));
            var root = Docx.Load(part).Root;
            if (root is null) continue;
            var ids = root.DescendantsAndSelf().Attributes().Where(IsReference).Select(a => a.Value).ToHashSet(StringComparer.Ordinal);
            foreach (var pair in part.Parts.Where(p => ids.Contains(p.RelationshipId)))
                AddReachable(pair.OpenXmlPart, covered);
        }
        return covered;

        static void AddReachable(OpenXmlPart part, HashSet<string> covered)
        {
            if (!covered.Add(Name(part))) return;
            foreach (var child in part.Parts) AddReachable(child.OpenXmlPart, covered);
        }
    }

    private static string Name(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
}
