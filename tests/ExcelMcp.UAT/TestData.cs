namespace ExcelMcp.UAT;

/// <summary>
/// Shared helpers for UAT tests: locates test-data workbooks by walking up
/// the directory tree from the test output directory.
/// </summary>
internal static class TestData
{
    /// <summary>Returns the absolute path to the given test-data file.</summary>
    /// <param name="fileName">File name, e.g. "ProjectTracking.xlsx"</param>
    public static string GetPath(string fileName)
    {
        var dir = new DirectoryInfo(Directory.GetCurrentDirectory());
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "test-data", fileName);
            if (File.Exists(candidate))
                return candidate;
            dir = dir.Parent;
        }

        throw new FileNotFoundException(
            $"test-data/{fileName} not found. Run 'scripts/create-sample-workbooks.ps1' to generate test data.",
            fileName);
    }

    /// <summary>
    /// Copies a test-data file to a temporary directory and returns the copy's path.
    /// Use this for write-operation tests to avoid modifying the originals.
    /// The caller is responsible for deleting the temp file (or rely on OS cleanup).
    /// </summary>
    public static string GetTempCopy(string fileName)
    {
        var source = GetPath(fileName);
        var tempDir = Path.Combine(Path.GetTempPath(), "ExcelMcp.UAT", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var dest = Path.Combine(tempDir, fileName);
        File.Copy(source, dest, overwrite: true);
        return dest;
    }
}
