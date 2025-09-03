using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace DocListManager
{
    public class DirCrawler
    {
        public static IEnumerable<string> StartCrawler(string directory, List<string> excludePatterns)
        {
            // Retrieve all files in the specified directory and its subdirectories
            IEnumerable<string> allFiles = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories);

            // Filter files based on exclusion patterns
            IEnumerable<string> filteredPaths = allFiles.Where(filePath => !isExcluded(excludePatterns, filePath));

            return filteredPaths;
        }

        public static IEnumerable<string> StartCrawlerDistributedDirs(string directory, List<string> extensions, List<string>  excludePatterns)
        {
            // Define the directory name patterns to match (e.g., _dist*, _appr*)
            List<string> directoryPatterns = new List<string> { "_dist", "_appr" };

            // Retrieve directories in the current level only
            var directories = Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly);

            bool hasMatchingDirectories = false;

            foreach (var dir in directories)
            {
                string dirName = Path.GetFileName(dir);

                // Stop traversing if the directory starts with _withdraw or _WITHDRAW
                if (dirName.StartsWith("_withdraw", StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Skip this directory and move on to the next one
                }

                if (isExcluded(excludePatterns, dir) &&
                !directoryPatterns.Any(dp => dirName.StartsWith(dp, StringComparison.OrdinalIgnoreCase)))
                {
                    continue; // Excluded and not in allow-list — skip
                }

                Dictionary<DocIdentGlobals.DocIdentFields, string> parentDirIdentifiers = CheckNC.GetIdentifiers(directory, true, false);

                // Check if the directory name matches the patterns
                bool matchesPattern = directoryPatterns.Any(pattern => dirName.StartsWith(pattern, StringComparison.OrdinalIgnoreCase));

                if (matchesPattern)
                {
                    string dirId = CheckNC.GetDirId(directory);
                    hasMatchingDirectories = true;

                    // Retrieve files with the specific extensions in the matching directory (top-level only)
                    // Also retrieving only files which match the base ID of the parent directory so that old I
                    var filesInDir = Directory.GetFiles(dir, "*.*", SearchOption.TopDirectoryOnly)
                        .Where(file =>
                            extensions.Contains(Path.GetExtension(file).ToLower()) &&
                            Path.GetFileNameWithoutExtension(file).StartsWith(dirId, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (filesInDir.Count > 0)
                    {

                        // Group files by base name (e.g., "report1")
                        var groupedFiles = filesInDir
                            .GroupBy(file => Path.GetFileNameWithoutExtension(file), StringComparer.OrdinalIgnoreCase);

                        foreach (var group in groupedFiles)
                        {
                            // For each group, get most recent file
                            var mostRecent = group.OrderByDescending(file => File.GetLastWriteTime(file)).First();
                            var baseFileName = Path.GetFileNameWithoutExtension(mostRecent);

                            // Get all files with the same base name, any extension
                            var relatedFiles = Directory.GetFiles(dir, $"{baseFileName}.*", SearchOption.TopDirectoryOnly)                                                        .ToList();

                            foreach (var file in relatedFiles)
                            {
                                yield return file;
                            }
                        }
                    }
                }
                else
                {
                    // Recurse into non-matching directories
                    foreach (var file in StartCrawlerDistributedDirs(dir, extensions, excludePatterns))
                    {
                        yield return file;
                    }
                }
            }

            // Optionally, you can log that no matching directories were found
            if (!hasMatchingDirectories)
            {
                // Log or handle the case where no matching directories were found at this level
            }
        }


        private static bool isExcluded(List<string> excludePatterns, string filePath)
        {
            // Extract the file name and directory name from the full path
            string fileName = Path.GetFileName(filePath);
            string directoryName = Path.GetDirectoryName(filePath);

            // Check if the file name or any part of the directory name matches the exclusion patterns
            return excludePatterns.Any(pattern =>
                isMatch(fileName, pattern) || directoryName.Split(Path.DirectorySeparatorChar).Any(dir => isMatch(dir, pattern)));
        }

        private static bool isMatch(string name, string pattern)
        {
            // Convert wildcard pattern to regex pattern
            string regexPattern = "^" + Regex.Escape(pattern)
                .Replace(@"\*", ".*")  // '*' matches any sequence of characters
                .Replace(@"\?", ".")   // '?' matches any single character
                + "$";

            // Perform case-insensitive match
            return Regex.IsMatch(name, regexPattern, RegexOptions.IgnoreCase);
        }
    }
}

