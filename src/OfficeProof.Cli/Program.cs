using System.CommandLine;
using System.Text;
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

// Windows consoles start in the OEM code page (437, 850, ...), which has no "—", "…" or "•":
// .NET drops or replaces them on screen and in redirected output. Write UTF-8, as on macOS and
// Linux, and give the console its code page back on exit.
if (!OperatingSystem.IsWindows()) return root.Parse(args).Invoke();
var consoleEncoding = Console.OutputEncoding;
Console.OutputEncoding = Encoding.UTF8;
try
{
    return root.Parse(args).Invoke();
}
finally
{
    Console.OutputEncoding = consoleEncoding;
}
