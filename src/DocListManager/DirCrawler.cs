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

        public static IEnumerable<string> StartCrawlerDistributedDirs(string directory, List<string> extensions)
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

                // Check if the directory name matches the patterns
                bool matchesPattern = directoryPatterns.Any(pattern => dirName.StartsWith(pattern, StringComparison.OrdinalIgnoreCase));

                if (matchesPattern)
                {
                    hasMatchingDirectories = true;

                    // Retrieve files with the specific extensions in the matching directory (top-level only)
                    var filesInDir = Directory.GetFiles(dir, "*.*", SearchOption.TopDirectoryOnly)
                                              .Where(file => extensions.Contains(Path.GetExtension(file).ToLower()))
                                              .ToList();

                    if (filesInDir.Count > 0)
                    {
                        // Get the base name of the most recent file
                        var mostRecentFile = filesInDir.OrderByDescending(file => new FileInfo(file).LastWriteTime).First();
                        string baseFileName = Path.GetFileNameWithoutExtension(mostRecentFile);

                        // Retrieve all files with the same base name but different extensions
                        var matchingFiles = Directory.GetFiles(dir, $"{baseFileName}.*", SearchOption.TopDirectoryOnly)
                                                     .Where(file => !file.Equals(mostRecentFile, StringComparison.OrdinalIgnoreCase))
                                                     .ToList();

                        // Include the most recent file in the result
                        matchingFiles.Add(mostRecentFile);

                        foreach (var file in matchingFiles)
                        {
                            yield return file;
                        }
                    }
                }
                else
                {
                    // Recurse into non-matching directories
                    foreach (var file in StartCrawlerDistributedDirs(dir, extensions))
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

