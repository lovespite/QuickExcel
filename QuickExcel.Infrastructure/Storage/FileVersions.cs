using System.Security.Cryptography;
using QuickExcel.Core;

namespace QuickExcel.Storage;

public static class FileVersions
{
    public static FileVersion Read(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return Read(stream);
        }
        catch (FileNotFoundException) { return new(null); }
        catch (DirectoryNotFoundException) { return new(null); }
    }
    public static FileVersion Read(Stream stream) => new(Convert.ToHexString(SHA256.HashData(stream)));
    public static void Require(string path, FileVersion expected)
    {
        if (Read(path) != expected) throw new WorkbookFileConflictException();
    }
}
