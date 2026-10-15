using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OfficeProof.Core;

/// <summary>Package-level helpers shared by every Word component.</summary>
public static class Docx
{
    /// <summary>
    /// Copies <paramref name="input"/> to <paramref name="output"/> and opens the copy for editing.
    /// Components never edit their input in place.
    /// </summary>
    public static WordprocessingDocument OpenCopy(string input, string output)
    {
        if (Path.GetFullPath(input) == Path.GetFullPath(output))
            throw new ArgumentException("output must differ from input; officeproof never edits in place", nameof(output));
        File.Copy(input, output, overwrite: true);
        return WordprocessingDocument.Open(output, isEditable: true);
    }

    /// <summary>Opens a document read-only.</summary>
    public static WordprocessingDocument OpenRead(string path) => WordprocessingDocument.Open(path, isEditable: false);

    /// <summary>
    /// The parts that carry document text: main body, headers, footers, footnotes, endnotes, comments.
    /// </summary>
    public static IEnumerable<OpenXmlPart> StoryParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is { } fn) yield return fn;
        if (main.EndnotesPart is { } en) yield return en;
        if (main.WordprocessingCommentsPart is { } c) yield return c;
    }

    /// <summary>Loads a part as an XDocument, preserving whitespace.</summary>
    public static XDocument Load(OpenXmlPart part)
    {
        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(s, LoadOptions.PreserveWhitespace);
    }

    /// <summary>Replaces a part's content with <paramref name="xml"/>.</summary>
    public static void Save(OpenXmlPart part, XDocument xml)
    {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        xml.Save(s, SaveOptions.DisableFormatting);
    }

    /// <summary>Writes raw bytes as a part's content.</summary>
    public static void Write(OpenXmlPart part, byte[] bytes)
    {
        using var s = part.GetStream(FileMode.Create, FileAccess.Write);
        s.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Schema validation with the Open XML SDK validator; every error is a finding with rule "schema".</summary>
    public static IReadOnlyList<Finding> Validate(OpenXmlPackage package, FileFormatVersions version = FileFormatVersions.Office2019) =>
        new OpenXmlValidator(version).Validate(package)
            .Select(e => new Finding("schema", Severity.Error, e.Description, e.Part?.Uri.ToString(), e.Path?.XPath))
            .ToList();

    /// <summary>Validates the document at <paramref name="path"/>.</summary>
    public static IReadOnlyList<Finding> Validate(string path)
    {
        using var doc = OpenRead(path);
        return Validate(doc);
    }

    // Unix regular file, rw-r--r--, in the high word of the zip external attributes.
    private const int RegularFileMode = 0x81A4 << 16;

    /// <summary>
    /// Gives every zip entry of the Office package at <paramref name="path"/> a normal file mode.
    /// System.IO.Packaging writes entries with mode 0 (zipinfo shows "?---------"): Word does not
    /// care, but `unzip` then extracts unreadable files (seen 2026-10-15 unpacking a redline).
    /// Entries are not reopened, so their compressed bytes are copied unchanged.
    /// </summary>
    public static void NormalizeEntryModes(string path)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Update);
        foreach (var entry in zip.Entries)
            if (entry.ExternalAttributes != RegularFileMode) entry.ExternalAttributes = RegularFileMode;
    }
}
