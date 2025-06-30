using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security;
using static DocListManager.XMLtoObject;

// Usage Example:
// DocListManager.exe  --xmlConfigFile "D:\rpande\itrust\testDocList\doclist_c#\FileWatcher\DocListManager\bin\Debug\net5.0\_input\DirToMonitor.xml" --docListTemplate "D:\rpande\itrust\testDocList\doclist_c#\FileWatcher\DocListManager\bin\Debug\net5.0\_input\DocList-Template.xlsm" --logdir N:\MG\ISMS\_input

namespace DocListManager
{
class Program
    {       
        static bool CheckDirectoryAndPermissions(string logDir)
    {
        try
        {               
                // Check if directory exists
                if (!System.IO.Directory.Exists(logDir))
            {
                Console.WriteLine($"Error: The directory {logDir} does not exist.");
                return false;
            }

            // Check write permissions by attempting to create a temporary file
            string testFilePath = Path.Combine(logDir, "test.tmp");
            File.WriteAllText(testFilePath, "test");
            File.Delete(testFilePath);
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            Console.WriteLine($"Error: No write permissions for the directory {logDir}.");
            return false;
        }
        catch (SecurityException)
        {
            Console.WriteLine($"Error: Security exception encountered for the directory {logDir}.");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: An unexpected error occurred while checking the directory {logDir}. Details: {ex.Message}");
            return false;
        }
    }

        static (string xmlConfigFile, string docListTemplate, string logFile, bool isVer) ProcessArguments(string[] args)
        {
            string xmlConfigFile = null;
            string docListTemplate = null;
            string logdir = null;
            bool isVersion = false;

            //TODO: If no command line arguments are given, display an error message in the terminal.

            // Iterate through command line arguments
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--version":
                        Console.WriteLine($"Version: {Assembly.GetExecutingAssembly().GetName().Version}");
                        isVersion = true;
                        break;
                    case "--xmlConfigFile":
                        // Ensure there is another argument after the flag
                        if (i + 1 < args.Length)
                        {
                            xmlConfigFile = args[i + 1];
                            i++; // Skip next argument (value)
                        }
                        else
                        {
                            Console.WriteLine("Error: --xmlConfigFile flag requires a value.");
                            return (null, null, null, isVersion); // Return null values to indicate failure
                        }
                        break;

                    case "--docListTemplate":
                        // Ensure there is another argument after the flag
                        if (i + 1 < args.Length)
                        {
                            docListTemplate = args[i + 1];
                            // Check that the file exists and is of type excel.
                            i++; // Skip next argument (value)
                        }
                        else
                        {
                            Console.WriteLine("Error: --docListTemplate flag requires a value.");

                            return (null, null, null, isVersion); // Return null values to indicate failure
                        }
                        break;
                    case "--logdir":
                        // Ensure there is another argument after the flag
                        if (i + 1 < args.Length)
                        {
                            string logdirS = args[i + 1];
                            logdir = logdirS.TrimEnd('"');                            
                     
                            // Check write permissions in this dir
                            // Call the method to check directory and permissions
                            if (!CheckDirectoryAndPermissions(logdir))
                            {
                                Console.WriteLine($"Error: Directory {logdir} does not exist or write permissions are not granted.");
                                return (null, null, null, isVersion);
                            }

                            i++; // Skip next argument (value)
                        }
                        else
                        {
                            Console.WriteLine("Error: --logdir flag requires a value.");
                            return (null, null, null, isVersion); // Return null values to indicate failure
                        }
                        break;

                    default:
                        Console.WriteLine($"Unknown argument: {args[i]}");
                        break;
                }
            }

            return (xmlConfigFile, docListTemplate, logdir, isVersion);
        }  

        static void Main(string[] args)
        {
            // xmlConfigFile = @"_input\\DirToMonitor.xml";
            // docListTemplate = @"_input\\DocList-Template.xlsm"
            (string xmlConfigFile, string docListTemplate, string logdir, bool isVersion) = ProcessArguments(args);

            if (isVersion == true && (xmlConfigFile== null && logdir==null && docListTemplate==null))
                return;

            if (logdir == null)
            {
                Console.WriteLine("Error: Argument: --logdir not provided");
                return;
            }

            Logger.Initialize(logdir);
            if (xmlConfigFile == null ||
                docListTemplate == null)
            {
                Logger.Instance.LogWrite("Error: Arguments not provided --xmlConfigFile, --docListTemplate");
                return;
            }
            
            var dirList = XMLtoObject.DeserializeObject(xmlConfigFile);                        

            foreach(var dir in dirList)
            {
                if (dir.active)
                {
                    Logger.Instance.LogWrite($"Arguments - xmlConfigFile: {xmlConfigFile}, docListTemplate {docListTemplate}");
                    Logger.Instance.LogWrite($"Directory path to monitor: {dir.ISMSWorkDir}");
                    Logger.Instance.LogWrite($"Directory path to save: {dir.docListSavePath}");
                    
                    string DirInput = dir.ISMSWorkDir;
                    DirectoryInfo Dir = new DirectoryInfo(DirInput);
                    string path = Dir.FullName;
                    string docListSavePath = dir.docListSavePath;
                    DirectoryInfo SDir = new DirectoryInfo(docListSavePath);
                    string spath = SDir.FullName;
                    string sDocListName = dir.docListName;
                    string sactivityName = dir.activityName;
                    string sProjectAcr = dir.activityAcronymn;
                    bool bdocListOverwrite = dir.docListOverwrite;
                    string sISMSPublishDir = dir.ISMSPublishDir;
                    string useExistingDocListEntries = dir.useExistingDocListEntries;
                    try
                    {
                        useExistingDocListEntries = Path.GetFullPath(dir.useExistingDocListEntries);
                        throw new FileNotFoundException("The file was not found.", dir.useExistingDocListEntries);
                    } catch (Exception ex) {
                        Logger.Instance.LogWrite($"The file provided by useExistingDocListEntries xml element is not found: {useExistingDocListEntries}");
                    }
                    string templateDoclistSheet = dir.templateDoclistSheet;
                    string templateMiscSheet = dir.templateMiscSheet;
                    string templateMappingSheet = dir.templateMappingSheet;
                    string existingDocListSheet = dir.existingDocListSheet;

                    List<string> excludeDir = new List<string>();

                    if (dir.Exclusions != null && dir.Exclusions.ExcludeNames != null)
                    {
                        foreach (var exclusion in dir.Exclusions.ExcludeNames)
                        {
                            string excludeName = exclusion.Name;
                            if (!string.IsNullOrEmpty(exclusion.PostFixWildCard))
                            {
                                excludeName = excludeName + exclusion.PostFixWildCard;                                
                            }
                            excludeDir.Add(excludeName);
                        }
                    }                   

                    Logger.Instance.LogWrite($"Directory paths excluded: {string.Join(", ", excludeDir)}");
                    Logger.Instance.LogWrite($"Directory paths published: {sISMSPublishDir}");
                    ExcelDB excelDB = new ExcelDB(docListTemplate, path, spath, sDocListName, sactivityName, sProjectAcr, bdocListOverwrite, excludeDir, sISMSPublishDir, useExistingDocListEntries, templateDoclistSheet, templateMiscSheet, templateMappingSheet, existingDocListSheet);
                    bool wtoConsole = true;
                    do
                    {
                        if (wtoConsole)
                        {
                            Console.WriteLine("File is currently locked. Please close the file.");
                            Logger.Instance.LogWrite("File is currently locked. Please close the file.");
                            wtoConsole = false;
                        }

                    } while (excelDB.IsFileLocked());

                    bool xlCreated = excelDB.CreateExcelDB();

                    if (xlCreated) Logger.Instance.LogWrite($"DocList created for directory {path}"); 
                    else Logger.Instance.LogWrite($"DocList could not be created");
                    // Dispose of the ExcelDB instance to free up resources
                    excelDB = null;
                    // Force garbage collection and finalize objects
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }                
            }            
           
        }

    }

    
 }