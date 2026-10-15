using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeProof.Templating;

/// <summary>What <see cref="SettingsInjector.Apply"/> added.</summary>
public sealed record SettingsChange(bool EmbedTrueTypeFontsAdded, bool UpdateFieldsAdded);

/// <summary>
/// Turns on the two settings a retemplated document needs, inserted at their schema position.
/// </summary>
/// <remarks>
/// CT_Settings is a strict xsd:sequence of ~90 children, so appending is a schema error. The
/// legacy CLI kept two hand-written partial order tables; this uses the SDK's schema-aware
/// <c>AddChild</c>, which knows the full sequence (including the w14/w15 extension elements
/// Word writes at the end).
/// <list type="bullet">
/// <item><c>w:embedTrueTypeFonts</c>: marks the document as carrying its fonts, so Word keeps
/// embedding them when it saves.</item>
/// <item><c>w:updateFields</c>: Word offers to refresh fields (TOC, PAGE, NUMPAGES) on open, so
/// cached results from the original document are not shipped stale. An existing
/// <c>w:updateFields</c> is left as the document has it.</item>
/// </list>
/// </remarks>
public static class SettingsInjector
{
    public static SettingsChange Apply(MainDocumentPart main)
    {
        var part = main.DocumentSettingsPart;
        Settings settings;
        if (part is null)
        {
            part = main.AddNewPart<DocumentSettingsPart>();
            settings = part.Settings = new Settings();
        }
        else
        {
            settings = part.Settings ??= new Settings();
        }

        var embed = settings.GetFirstChild<EmbedTrueTypeFonts>() is null;
        if (embed) settings.AddChild(new EmbedTrueTypeFonts(), throwOnError: true);
        var update = settings.GetFirstChild<UpdateFieldsOnOpen>() is null;
        if (update) settings.AddChild(new UpdateFieldsOnOpen { Val = true }, throwOnError: true);
        if (embed || update) settings.Save();
        return new SettingsChange(embed, update);
    }
}
