using System;
using System.IO;

namespace DRIT.Spreadsheet.Examples.Shared
{
    internal static class ExampleSupport
    {
        internal static string RootDirectory
        {
            get
            {
                var directory = new DirectoryInfo(Environment.CurrentDirectory);
                while (directory != null)
                {
                    if (File.Exists(Path.Combine(directory.FullName, "DRIT.Spreadsheet.Examples.sln")))
                    {
                        return directory.FullName;
                    }

                    directory = directory.Parent;
                }

                return Environment.CurrentDirectory;
            }
        }

        internal static string InputPath(string fileName)
        {
            return EnsurePath("In", fileName);
        }

        internal static string OutputPath(string fileName)
        {
            return EnsurePath("Out", fileName);
        }

        internal static void Require(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }

        internal static void ReportRoundTrip(string path, Workbook workbook)
        {
            Console.WriteLine($"Saved: {path}");
            Console.WriteLine($"Worksheets after reload: {workbook.Sheets.Count}");
        }

        private static string EnsurePath(string directoryName, string fileName)
        {
            var directory = Path.Combine(RootDirectory, directoryName);
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, fileName);
        }
    }
}