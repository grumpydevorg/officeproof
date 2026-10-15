namespace OfficeProof.Edit.Tests;

/// <summary>A per-test scratch directory under the system temp path.</summary>
internal sealed class TempDir : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeproof-edit-" + Guid.NewGuid().ToString("N"));

    public TempDir() => Directory.CreateDirectory(Path);

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose() => Directory.Delete(Path, recursive: true);
}
