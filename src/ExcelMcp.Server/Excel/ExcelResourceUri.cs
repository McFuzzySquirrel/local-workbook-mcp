namespace ExcelMcp.Server.Excel;

internal static class ExcelResourceUri
{
    public static readonly Uri WorkbookUri = new("excel://workbook");

    public static Uri CreateWorksheetUri(string worksheetName)
    {
        return new Uri($"excel://worksheet/{Uri.EscapeDataString(worksheetName)}");
    }

    public static Uri CreateTableUri(string worksheetName, string tableName)
    {
        return new Uri($"excel://worksheet/{Uri.EscapeDataString(worksheetName)}/table/{Uri.EscapeDataString(tableName)}");
    }

    public static bool TryParse(Uri uri, out string? worksheet, out string? table)
    {
        worksheet = null;
        table = null;

        if (!UriEquals(uri.Scheme, "excel"))
        {
            return false;
        }

        var host = uri.Host;
        if (UriEquals(host, "workbook"))
        {
            return true;
        }

        if (!UriEquals(host, "worksheet"))
        {
            return false;
        }

        var segments = uri.AbsolutePath.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0)
        {
            return false;
        }

        worksheet = Uri.UnescapeDataString(segments[0]);
        if (segments.Length >= 3 && UriEquals(segments[1], "table"))
        {
            table = Uri.UnescapeDataString(segments[2]);
        }

        return true;
    }

    private static bool UriEquals(string? left, string right)
    {
        return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }
}
