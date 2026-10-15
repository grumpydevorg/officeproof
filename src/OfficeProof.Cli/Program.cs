using System.CommandLine;
using OfficeProof.Cli;

// officeproof is a thin host: every verb parses its options, calls one or more component
// libraries (Templating, Lint, Redline, Verify, Spreadsheet) and prints their results.
// Components never print; exit codes follow the findings (1 when any finding is an error).
var root = new RootCommand("officeproof: composable tools for Word (.docx/.dotx) and Excel (.xlsx) packages")
{
    TemplateCommands.Generate(),
    TemplateCommands.Retemplate(),
    CheckCommands.Validate(),
    CheckCommands.Lint(),
    RedlineCommand.Create(),
    EditCommands.Replace(),
    CheckCommands.Verify(),
    CheckCommands.Outline(),
    XlsxCommands.Create(),
};
return root.Parse(args).Invoke();
