using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using static DocListManager.XMLtoObject;
using Microsoft.Office.Interop.Excel;
using System.Reflection.Metadata.Ecma335;
using static DocListManager.FileProp;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Office2010.Excel;
using Shell32;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocListManager;

namespace DocListManager
{
    class ExcelDB
    {
        // Variables which stay same across diff instances of Directory entries
        private static string template;
        private static string fullpathtemplate;
        private static DateTime currentdate = DateTime.Today;

        // Variables which Vary  across instances of Directory entries
        private string directory;
        private string activityName;
        private string ofilepath;
        private string filename;
        private string projectAcr;
        private bool docListOverwrite;
        private List<string> excludeDirList;
        private string sISMSPublishDir;
        private string sUseExistingDocListEntries;
        private string templateDoclistSheet;
        private string templateMiscSheet;
        private string templateMappingSheet;
        private string existingDocListSheetName;

        private Dictionary<string, string> tabletype = new Dictionary<string, string>();
        private Dictionary<string, string> tableorg = new Dictionary<string, string>();

        public void populateExistingDoclistEntries(Excel.Application xlApp, string sUseExistingDocListEntries, Dictionary<string, DocListExistingEntries> existingDocList, string existingDocListSheetName)
        {
            if (File.Exists(sUseExistingDocListEntries))
            {
                // Check if the file is locked
                if (IsExistingFileLocked(sUseExistingDocListEntries))
                {
                    Logger.Instance.LogWrite($"The file {sUseExistingDocListEntries} is currently locked and cannot be read: " + sUseExistingDocListEntries);
                    return;
                }
                if(xlApp == null)
                {
                    Logger.Instance.LogWrite($"The Excel application not running ");
                    return;
                }

                Excel.Workbook ExistingDocList = xlApp.Workbooks.Open(sUseExistingDocListEntries);
                Excel.Worksheet existingDocListSheet = GetWorksheetByName(ExistingDocList, existingDocListSheetName);

                if (existingDocListSheet != null)
                {
                    // Assuming the Excel worksheet starts at row 2 (row 1 is header)
                    int startRow = 2;
                    int lastRow = existingDocListSheet.Cells[existingDocListSheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row; // Find the last non-empty row

                    // Loop through each row in the worksheet
                    for (int row = startRow; row <= lastRow; row++)
                    {
                        // Extract the key (assuming it's in the column corresponding to eColId)
                        string key = existingDocListSheet.Cells[row, (int)DocIdentGlobals.DocIdentFieldsExcelColumns.eColId].Value?.ToString();

                        if (!string.IsNullOrEmpty(key))
                        {
                            // Create a dictionary to store the field values for this row
                            var docEntryIdentifiers = new Dictionary<DocIdentGlobals.DocIdentFieldsExcelColumns, string>();

                            // Populate the dictionary with all the column values
                            foreach (DocIdentGlobals.DocIdentFieldsExcelColumns col in Enum.GetValues(typeof(DocIdentGlobals.DocIdentFieldsExcelColumns)))
                            {
                                string value = existingDocListSheet.Cells[row, (int)col].Value?.ToString();
                                docEntryIdentifiers[col] = value ?? string.Empty; // Use empty string if value is null
                            }

                            // Create a new DocListExistingEntries object with the populated dictionary
                            try
                            {
                                var entry = new DocListExistingEntries(docEntryIdentifiers);

                                // Add the entry to the dictionary, handle duplicates if necessary
                                if (!existingDocList.ContainsKey(key))
                                {
                                    existingDocList.Add(key, entry);
                                }
                                else
                                {
                                    // Handle duplicates (e.g., overwrite or log the duplicate entry)
                                    Logger.Instance.LogWrite($"Duplicate key found in row {row}: {key}. Overwriting entry.");
                                    existingDocList[key] = entry; // Overwrite the existing entry
                                }
                            }
                            catch (ArgumentException ex)
                            {
                                // Log any issues with the dictionary size (should be rare if all columns are correctly mapped)
                                Logger.Instance.LogWrite($"Error creating DocListExistingEntries for row {row}: {ex.Message}");
                            }
                        }
                        else
                        {
                            // Handle cases where the key is empty or null
                            Logger.Instance.LogWrite($"Empty key found in row {row}. Skipping.");
                        }
                    }

                    // Logging the population result
                    Logger.Instance.LogWrite($"DocList population completed. Total entries: {existingDocList.Count}");
                }
                else
                {
                    // Handle the case where the worksheet does not exist
                    Logger.Instance.LogWrite($"Worksheet 'DocList' not found in the specified workbook {sUseExistingDocListEntries}."+ sUseExistingDocListEntries);
                }

                ExistingDocList.Close(false);
            }
        }

        public Excel.Worksheet GetWorksheetByName(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet; // Return the worksheet if found
                }
            }
            Logger.Instance.LogWrite($"Worksheet {sheetName} not found in the specified workbook {workbook.Name}");
            return null; // Return null if not found
        }

        private bool IsExistingFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    // If we can open the file, it's not locked
                    return false;
                }
            }
            catch (IOException)
            {
                // If an IOException occurs, the file is likely locked
                return true;
            }
        }
           
        public ExcelDB(string szTemplate, string path, string spath, string sdocListName, string sactivityName, string sProjectAcr, bool bdocListOverwrite, List<string> lExcludeDirList, string ISMSPublishDir, string useExistingDocListEntries, string stemplateDoclistSheet, string stemplateMiscSheet, string stemplateMappingSheet, string sexistingDocListSheet)
        {
            template = szTemplate;
            fullpathtemplate = Path.GetFullPath(template);
            activityName = sactivityName;
            projectAcr = sProjectAcr;
            docListOverwrite = bdocListOverwrite;
            excludeDirList = lExcludeDirList;
            sISMSPublishDir = ISMSPublishDir;
            sUseExistingDocListEntries = useExistingDocListEntries;
            templateDoclistSheet = stemplateDoclistSheet;
            templateMiscSheet = stemplateMiscSheet;
            templateMappingSheet = stemplateMappingSheet;
            existingDocListSheetName = sexistingDocListSheet;

            directory = path;

            
            if (sdocListName != null)
            {
                // Get Old version Number from file Name 
                string currentMostRecentDocListName = GetMostRecentDocList(spath, sdocListName);

                if (!docListOverwrite)
                {
                    if (!string.IsNullOrEmpty(currentMostRecentDocListName))
                    {
                        string currentVersion = ExtractVersion(currentMostRecentDocListName);
                        string currentPrefix = ExtractPrefix(currentMostRecentDocListName);
                        string updatedVersion = IncrementThirdLevelVersion(currentVersion);
                        // Get extension of the template
                        string extension = Path.GetExtension(template);
                        string updatedFileName = currentPrefix + updatedVersion + extension;
                        ofilepath = spath + "\\" + updatedFileName;
                    }
                } else {
                    if (!string.IsNullOrEmpty(currentMostRecentDocListName))
                    { 
                        // docListOverwrite doclist if it exists 
                        ofilepath = spath + "\\" + currentMostRecentDocListName;
                    }

                }
            }

     
            if (string.IsNullOrEmpty(ofilepath))
            {
                // Default doclist name
                if (string.IsNullOrEmpty(sdocListName)) { 
                    // in case or error and docListName is not set
                    filename = "0D_PMD_" + projectAcr + "-DocList_v0.1.xlsm";
                    ofilepath = spath + "\\" + filename;
                }
                else 
                {
                    // Get from docListName set in Template file
                    filename = sdocListName + "_v0.1.xlsm";
                    ofilepath = spath + "\\" + filename;
                }
            }
        }

        public static string ExtractPrefix(string filename)
        {
            // Find the index of the first occurrence of "_v"
            int index = filename.IndexOf("_v");

            // If "_v" is found, extract the substring before it
            if (index != -1)
            {
                return filename.Substring(0, index);
            }
            // If "_v" is not found, return the original filename
            else
            {
                return filename;
            }
        }

        public static string IncrementThirdLevelVersion(string version)
        {
            // Split the version string by dots
            string[] versionParts = version.Split('.');

            // Check if the version has at least three levels
            if (versionParts.Length >= 3)
            {
                // Increment the third level version number
                versionParts[2] = (int.Parse(versionParts[2]) + 1).ToString();

                // Remove any additional version levels
                Array.Resize(ref versionParts, 3);
            }
            else
            {
                // If the version has less than three levels, add '.1' as the third level
                Array.Resize(ref versionParts, 3);
                versionParts[2] = "1";
            }

            // Join the version parts back together
            string newVersion = string.Join(".", versionParts);

            return newVersion;
        }

        public static string ExtractVersion(string fileName)
        {
            // Regular expression to match the version part, e.g., "_v0.2"
            Regex versionRegex = new Regex(@"_v\d+(\.\d+)+", RegexOptions.IgnoreCase);
            Match match = versionRegex.Match(fileName);
            if (match.Success)
            {
                return match.Value;
            }
            return null;
        }

        public string GetMostRecentDocList(string spath, string sdocListName)
        {
            string sMostRecentDocList = "";
            // Get files matching the prefix and extension criteria
            DirectoryInfo directoryInfo = new DirectoryInfo(spath);
            FileInfo[] files = directoryInfo.GetFiles($"{sdocListName}_v*.*")
                                             .Where(file => file.Extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                                                            file.Extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
                                                            file.Extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                                                            )
                                             .ToArray();

            // Get the most recent file
            FileInfo mostRecentFile = files.OrderByDescending(file => file.LastWriteTime).FirstOrDefault();

            if (mostRecentFile != null)
            {
                sMostRecentDocList = mostRecentFile.Name;
            }

            return sMostRecentDocList;
        }

       //If file does not exist exception is also catched. Has to be taken care of.
        public bool IsFileLocked()
        {
            bool isOpened = false;
            if (File.Exists(ofilepath))
            {
                try
                {
                    using (FileStream fs = File.Open(ofilepath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        // File is not locked, so we can close it
                        fs.Close();
                    }
                }
                catch (IOException ex) when ((ex.HResult & 0x0000FFFF) == 32)
                {
                    isOpened = true; // File is locked
                }
            }

            return isOpened;
        }

        private static void ReleaseObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Logger.Instance.LogWrite($"releasing COM object:");
                    Marshal.FinalReleaseComObject(obj);
                }
                catch (Exception ex)
                {
                    Logger.Instance.LogWrite($"Error releasing COM object: {ex.Message}");
                }
                finally
                {
                    obj = null;
                }
            }
        }

        public static void AddDocEntryIfValid(List<DocIdent> list, Dictionary<DocIdentGlobals.DocIdentFields, string> docEntryIdentifiers)
        {
            // Create a new DocIdent instance from the provided dictionary
            var newEntry = new DocIdent(docEntryIdentifiers);

            // Check if an entry with the same ID already exists in the list
            var existingEntry = list.FirstOrDefault(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eId) == newEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId));

            if (existingEntry != null)
            {
                // If an existing entry is found, handle comments
                var newComment = newEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors);

                if (!string.IsNullOrWhiteSpace(newComment))
                {
                    string newCommentTrimmed = newComment.Trim();

                    // Check if the new comment is not already a substring of the existing comment
                    if (!existingEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors).Contains(newCommentTrimmed, StringComparison.OrdinalIgnoreCase))
                    {
                        // Append the new comment to the existing comment
                        string updatedComments = existingEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors);
                        updatedComments += string.IsNullOrWhiteSpace(updatedComments) ? newCommentTrimmed : " " + newCommentTrimmed;

                        // Set the updated comments
                        existingEntry.SetFieldValue(DocIdentGlobals.DocIdentFields.eErrors, updatedComments);
                    }

                    // Optionally log the action (comment out if not needed)
                    // Console.WriteLine($"Updated entry with ID {newEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)} by appending comments.");
                }
                else
                {
                    // Optionally log if the new entry has no comments
                    // Console.WriteLine($"Entry with ID {newEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)} has an empty comment. No changes made.");
                }
            }
            else
            {
                // If no existing entry is found, add the new entry to the list
                list.Add(newEntry);

                // Optionally log the action (comment out if not needed)
                // Console.WriteLine($"Added new entry with ID {newEntry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)} to the list.");
            }
        }

        public void LoadTablesFromMappingSheet(Excel.Worksheet sheet)
        {
            LoadTableType(sheet);
            LoadTableOrg(sheet);
        }
        private void LoadTableType(Excel.Worksheet sheet)
        {
            try
            {
                // Attempt to load the table
                LoadTable(sheet, "TableType", tabletype);
            }
            catch (Exception ex)
            {
                // Log the error if the LoadTable call fails
                // You can use any logging mechanism, such as Console.WriteLine, or use a logging framework like NLog or log4net.
                string errorMessage = $"Error loading table: {ex.Message}";

                // If you are using a logging class like Logger, you can log the error message
                Logger.Instance.LogWrite(errorMessage);

            }
        }

        private void LoadTableOrg(Excel.Worksheet sheet)
        {
             try
            {
                // Attempt to load the table
                LoadTable(sheet, "TableOrg", tableorg);
            }
            catch (Exception ex)
            {
                // Log the error if the LoadTable call fails
                // You can use any logging mechanism, such as Console.WriteLine, or use a logging framework like NLog or log4net.
                string errorMessage = $"Error loading table: {ex.Message}";

                // If you are using a logging class like Logger, you can log the error message
                Logger.Instance.LogWrite(errorMessage);

            }
        }

        private void LoadTable(Excel.Worksheet sheet, string tableName, Dictionary<string, string> table)
        {
            // Assuming the first row contains headers
            Excel.Range tableRange = sheet.ListObjects[tableName].Range;
            if (tableRange == null)
            {
                throw new ArgumentException($"The table '{tableName}' does not exist in the mapping sheet.");
            }

            for (int row = 2; row <= tableRange.Rows.Count; row++) // Start from row 2 to skip headers
            {
                string key = (tableRange.Cells[row, 1] as Excel.Range).Text;
                string value = (tableRange.Cells[row, 2] as Excel.Range).Text;

                if (!string.IsNullOrWhiteSpace(key))
                {
                    table[key] = value;
                }
            }
        }

        public void updateErrorMsgIfUnique(DocIdent entry, string newComment)
        {
            string existingComment = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors);

            // Update the comment if needed
            if (!string.IsNullOrEmpty(existingComment))
            {
                // Append new comment if it does not already exist
                if (!existingComment.Contains(newComment))
                {
                    string updatedComment = existingComment + "\n" + newComment;
                    entry.SetFieldValue(DocIdentGlobals.DocIdentFields.eErrors, updatedComment);
                }
            }
            else
            {
                // If no existing comment, set the new one
                entry.SetFieldValue(DocIdentGlobals.DocIdentFields.eErrors, newComment);
            }
        }

        public void updateDocEntryErrorIfUnique(Dictionary<DocIdentGlobals.DocIdentFields, string> docEntryIdentifiers, string errorString)
        {
            string existingComment = docEntryIdentifiers[DocIdentGlobals.DocIdentFields.eErrors];

            // Update the comment if needed
            if (!string.IsNullOrEmpty(existingComment))
            {
                // Append new comment if it does not already exist
                if (!existingComment.Contains(errorString))
                {
                    string updatedComment = existingComment + "\n" + errorString;
                    docEntryIdentifiers[DocIdentGlobals.DocIdentFields.eErrors] = updatedComment;

                }
            }
            else
            {
                // If no existing comment, set the new one
                docEntryIdentifiers[DocIdentGlobals.DocIdentFields.eErrors] = errorString;
            }

        }
        public bool CreateExcelDB()
        {
            Excel.Application xlApp = null;
            Excel.Workbook DocList = null;
            Excel.Worksheet hist = null;
            Excel.Worksheet doclist = null;
            Excel.Worksheet misc = null;
            Excel.Worksheet mapping = null;

            bool success = true;

            try
            {
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;

                if (xlApp == null)
                {
                    success = false;
                    Logger.Instance.LogWrite("Excel is not properly installed!!");
                    return success;
                }

                if (!File.Exists(fullpathtemplate))
                {
                    return false;
                }
                // Check the file extension
                string extension = Path.GetExtension(fullpathtemplate).ToLower();
                if (extension != ".xls" && extension != ".xlsx" && extension != ".xlsm")
                {
                    return false;
                }

                DocList = xlApp.Workbooks.Open(fullpathtemplate);

                hist = DocList.Worksheets[1];
                doclist = GetWorksheetByName(DocList, templateDoclistSheet);
                misc = GetWorksheetByName(DocList, templateMiscSheet);
                mapping = GetWorksheetByName(DocList, templateMappingSheet);

                
                if (hist == null || doclist == null || misc == null || mapping == null)
                {
                    return false;
                }

                // Check following tables in Mapping sheet..
                if (mapping != null)
                {
                    LoadTablesFromMappingSheet(mapping);
                }

                hist.Cells[9, 3] = currentdate;
                hist.Cells[15, 2] = currentdate;
                hist.Cells[1, 1] = activityName;
                hist.Cells[2, 1] = $"{projectAcr} - DocList";


                List<DocIdent> docListConfEntries = new List<DocIdent>();
                List<DocIdent> docListNonConfEntries = new List<DocIdent>();

                Dictionary<string, DocListExistingEntries> existingDocList = new Dictionary<string, DocListExistingEntries>();
                if (!string.IsNullOrEmpty(this.sUseExistingDocListEntries))
                {
                    populateExistingDoclistEntries(xlApp, this.sUseExistingDocListEntries, existingDocList, existingDocListSheetName);
                }

                //In This run _dist*, _appro* entries will be ignored 
                using (IEnumerator<string> filepaths = DirCrawler.StartCrawler(directory, excludeDirList).GetEnumerator())
                {
                    while (filepaths.MoveNext())
                    {
                        string file = filepaths.Current;

                        // Get file attributes and check if it's hidden
                        FileInfo fileInfo = new FileInfo(file);
                        if ((fileInfo.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden || fileInfo.Length==0)
                        {
                            // Skip hidden files
                            continue;
                        }

                        bool conformity = CheckNC.CheckConformity(file);

                        //string[] identifiers = CheckNC.GetIdentifiers(file, conformity);
                        Dictionary<DocIdentGlobals.DocIdentFields, string> docEntryIdentifiers = CheckNC.GetIdentifiers(file, conformity, false/*not in distr*/);


                        if (conformity)
                        {
                            FileProp fileprop = new FileProp(file);
                            if (fileprop.IsShortcut())
                            {
                                if (!fileprop.ResolveShortcutTarget())
                                    updateDocEntryErrorIfUnique(docEntryIdentifiers, "Invalid target of shortcut.");
                            }
                            docListConfEntries.Add(new DocIdent(docEntryIdentifiers));
                            Logger.Instance.LogWrite($"Read FileName: {file}, {string.Join(" -> ", docEntryIdentifiers.Select(kv => $"{kv.Key}: {kv.Value}"))}");
                        }
                        else
                        {
                            bool smallconformity = CheckNC.CheckSmallConformity(file);
                            FileProp fileprop = new FileProp(file);
                            if (fileprop.IsShortcut())
                            {
                                if (!fileprop.ResolveShortcutTarget())
                                    updateDocEntryErrorIfUnique(docEntryIdentifiers, "Invalid target of shortcut.");
                            }
                            if (smallconformity)
                            {
                                updateDocEntryErrorIfUnique(docEntryIdentifiers, "NC not correct.");// set 11 as comments index and make it a const.                        
                            }
                            docListNonConfEntries.Add(new DocIdent(docEntryIdentifiers));
                        }
                    }
                }

                List<string> extensions = new List<string> {".pdf", ".docx", ".xlsx", ".xlsm"};

                // Explicity crawl _dist* , _approv* directories only
                using (IEnumerator<string> filepaths = DirCrawler.StartCrawlerDistributedDirs(directory, extensions).GetEnumerator())
                {
                    while (filepaths.MoveNext())
                    {
                        string file = filepaths.Current;
                        Logger.Instance.LogWrite($"Processing DIR file: {file}");
                        FileProp fileProp = new FileProp(file);
                        string classification = fileProp.getDocumentProperty(file, fileProp.getFileExtension(), DocListManager.FileProp.eDocumentProperty.Classification);
                        Logger.Instance.LogWrite($"File {file}: Classification = {classification} ");

                        Dictionary<DocIdentGlobals.DocIdentFields, string> docEntryIdentifiers = CheckNC.GetIdentifiers(file, true, true /* in dist folder*/);
                        Logger.Instance.LogWrite($"File {file}: identifiers: {string.Join(" -> ", docEntryIdentifiers.Select(kv => $"{kv.Key}: {kv.Value}"))}");

                        // Ascertain the classification of doc in distributed or approved folder
                        if (fileProp.checkParentDirIsDistributed(file))
                        {
                            // Sanitize classification
                            if (!string.IsNullOrEmpty(classification) &&
                                !classification.ToLower().StartsWith("interne") &&
                                !classification.ToLower().StartsWith("internal") && 
                                !classification.ToLower().StartsWith("public") )
                            {

                                // Error if document is not public or internal
                                updateDocEntryErrorIfUnique(docEntryIdentifiers, $"Classification of latest distributed document in _distr* folder should be Internal or public");
//                                docEntryIdentifiers[DocIdentGlobals.DocIdentFields.eErrors] = $"Classification of latest distributed document in _distr* folder should be Internal or public";
                            }

                            // Sanitize that the document with same time stamp exists in publication folder as well
                            // check only the .pdf versions of file in publication folder
                            // For other files it will create a very complex logic
                            if (!string.IsNullOrEmpty(sISMSPublishDir) && fileProp.getFileExtension() == ".pdf")
                            {
                                string expectedCheckSum = CheckNC.ComputeFileChecksum(file);
                                string comments = CheckNC.FindAndVerifyFile(file, sISMSPublishDir, expectedCheckSum);
                                if (!string.IsNullOrEmpty (comments))
                                {
                                    updateDocEntryErrorIfUnique(docEntryIdentifiers, comments);
                                   // docEntryIdentifiers[DocIdentGlobals.DocIdentFields.eErrors] = comments;                                    
                                }
                            }
                            // Only add if either comments exist and the entry doesnt already exist.
                            AddDocEntryIfValid(docListConfEntries, docEntryIdentifiers);                            
                        } 
                        else if (fileProp.checkParentDirIsApproved(file))
                        {
                            if (!string.IsNullOrEmpty(classification) && !classification.ToLower().StartsWith("restr"))
                            {
                                // Error if document is not restricted
                                updateDocEntryErrorIfUnique(docEntryIdentifiers, "Classification of latest approved document in _appr* folder should be Restricted");
                                //docEntryIdentifiers[DocIdentGlobals.DocIdentFields.eErrors] = "Classification of latest approved document in _appr* folder should be Restricted";
                            }
                            AddDocEntryIfValid(docListConfEntries, docEntryIdentifiers);
                        }                        
                    }
                }
                    
                    // Sanity - 1: Entries with same id, title, version, extension - Duplicate entry

                    // Group entries by Id, TitleAcr, Version, and Ext, then filter groups with more than one item
                    var duplicateUniqueEntries = docListConfEntries
                        .GroupBy(entry => new
                        {
                            Id = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId),
                            Title = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle),
                            Version = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion),
                            Extension = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension)
                        })
                        .Where(group => group.Count() > 1)
                        .SelectMany(group => group);

                    // Entries duplicate
                    foreach (var entry in duplicateUniqueEntries)
                    {
                        var currentComment = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors);
                        // Define the new comment to append
                        var newComment = "Duplicate values for this entry (id, title, version, extension is same).";
                        updateErrorMsgIfUnique(entry, newComment);

                        Logger.Instance.LogWrite($"After Duplicate entry: Id = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)}, " +
                                                 $"TitleAcr = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle)}, " +
                                                 $"Version = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion)}, " +
                                                 $"Ext = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension)}, " +
                                                 $"Comment = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors)}");
                    }


                // Sanity - 2: Entries with same id, domain, title, type, extension, folder - Dulicate entryEntries need archiving
                // Group entries by Id, Domain, TitleAcr, Type, Extension, and Folder, then filter groups with more than one item
                var duplicateUnarchivedEntries = docListConfEntries
                    .GroupBy(entry => new
                    {
                        Id = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId),
                        Domain = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eDomain),
                        Title = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle),
                        Type = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eType),
                        Extension = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension),
                        Folder = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eFolder)
                    })
                    .Where(group => group.Count() > 1)
                    .SelectMany(group => group);

                // Update the comment for entries that need archiving, if the comment is not already set
                foreach (var entry in duplicateUnarchivedEntries)
                {
                    // Define the new comment to append
                    var newComment = "Please archive the old entries in this folder (title, version, ext is same)";
                    updateErrorMsgIfUnique(entry, newComment);

                    Logger.Instance.LogWrite($"After Archived entry: Id = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)}, " +
                                            $"TitleAcr = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle)}, " +
                                            $"Version = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion)}, " +
                                            $"Ext = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension)}, " +
                                            $"Comment = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors)}");
                    
                }


                // Sanity 3 - Same id, domain, type , folder but different Title.
                // Group entries by Id, Domain, Type, Extension, and Folder, then filter groups with more than one item
                var duplicateUniqueEntriesDiffTitle = docListConfEntries
                    .GroupBy(entry => new
                    {
                        Id = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId),
                        Domain = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eDomain),
                        Type = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eType),
                        Extension = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension),
                        Folder = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eFolder)
                    })
                    .Where(group => group.Count() > 1)
                    .SelectMany(group => group);

                // Entries duplicate
                foreach (var entry in duplicateUniqueEntriesDiffTitle)
                {
                    // Define the new comment to append
                    string newComment = "Duplicate values for this entry (id, domain, type, extension is same but title = " +
                    entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle) + " is different).";
                    updateErrorMsgIfUnique(entry, newComment);
                    // Log the details of the duplicate entry
                    Logger.Instance.LogWrite($"After Duplicate entry: Id = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)}, " +
                                                $"TitleAcr = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle)}, " +
                                                $"Version = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion)}, " +
                                                $"Ext = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension)}, " +
                                                $"Comment = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors)}");
                    
                }

                // Sanity 3.1 - ascertain that parent folder ID is a subset of id 
                var invalidIDWrtParent = docListConfEntries
                    .GroupBy(entry => new
                    {
                        Id = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId),
                        Domain = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eDomain),
                        Type = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eType),
                        Extension = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension),
                        Folder = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eFolder)
                    })
                    //.Where(group => group.Count() > 1)
                    .SelectMany(group => group
                        .Where(entry => !entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)
                                        .Contains(entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eParentFolderID))))
                    .ToList();

                foreach (var entry in invalidIDWrtParent)
                {
  
                    // Construct the new comment with details about the ID mismatch
                    StringBuilder newCommentBuilder = new StringBuilder();
                    newCommentBuilder.Append("ID of the parent directory (")
                                     .Append(entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eParentFolderID))
                                     .Append(") is not a subset of ID of the file = (")
                                     .Append(entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId))
                                     .Append(").");

                    string newComment = newCommentBuilder.ToString();

                    // Log the details of the mismatched entry
                    Logger.Instance.LogWrite($"ID of the parent directory: Id = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eParentFolderID)}, " +
                                             $"is not a subset of child id = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId)}, {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eFilename)}," +
                                             $" TitleAcr = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle)}, " +
                                             $"Version = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion)}, " +
                                             $"Ext = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension)}, " +
                                             $"Comment = {entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors)}");

                    // Update the comment if needed
                    updateErrorMsgIfUnique(entry, newComment);
                }


                // Sanity 4 - If the ISMSPublishDir is provided as a config
                if (!string.IsNullOrWhiteSpace(sISMSPublishDir))
                    {
                        foreach (var entry in duplicateUniqueEntriesDiffTitle)
                        {
                            // Get most recent entry in _distr* folder if it exists
                            // Check this entry is present in sISMSPublishDir folder
                            // Check the Classification of document should be internal/public (Comments field)
                            // All entries in Distr folder must have a level 1/2 (final version)

                            // If _approve* folder exists
                            // the most recent file in _approve folder if it exists should 
                            // All entries in _approved folder must have a level 1 / 2 (final version)

                        }
                    }
                // Sort the entries 
                // Sort docListConfEntries by multiple fields
                var sortedDocs = docListConfEntries
                    .OrderBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eId))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eType))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eChangeDate))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors))
                    .ToList();
                docListConfEntries = sortedDocs;

                // Sort docListNonConfEntries by multiple fields
                var sortedNCDocs = docListNonConfEntries
                    .OrderBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eFolder))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eFilename))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eChangeDate))
                    .ThenBy(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors))
                    .ToList();
                docListNonConfEntries = sortedNCDocs;

                var nonDuplicateEntriesRemovingPDFExtensionEntries = docListConfEntries
                    .GroupBy(d => new
                    {
                        eId = d.GetFieldValue(DocIdentGlobals.DocIdentFields.eId),
                        eType = d.GetFieldValue(DocIdentGlobals.DocIdentFields.eType),
                        eAcronymn = d.GetFieldValue(DocIdentGlobals.DocIdentFields.eAcronymn),
                        eVersion = d.GetFieldValue(DocIdentGlobals.DocIdentFields.eVersion),
                        eFolder = d.GetFieldValue(DocIdentGlobals.DocIdentFields.eFolder)
                    })
                    .SelectMany(group =>
                    {
                        // Check if the group contains both .pdf and non-.pdf files based on eExtension
                        var pdfDocs = group.Where(d => d.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension).Equals(".pdf", StringComparison.OrdinalIgnoreCase)).ToList();
                        var nonPdfDocs = group.Where(d => !d.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension).Equals(".pdf", StringComparison.OrdinalIgnoreCase)).ToList();

                        // If there are non-pdf documents in the group, remove the .pdf documents
                        if (nonPdfDocs.Any())
                        {
                            // Return all non-pdf docs and exclude all .pdf docs
                            return group.Where(d => !d.GetFieldValue(DocIdentGlobals.DocIdentFields.eExtension).Equals(".pdf", StringComparison.OrdinalIgnoreCase));
                        }
                        else
                        {
                            // If there are no non-pdf documents, return the group as is
                            return group;
                        }
                    })
                    .ToList();
                docListConfEntries = nonDuplicateEntriesRemovingPDFExtensionEntries;

                Dictionary<string, DocIdent> mapDocList = new Dictionary<string, DocIdent>();
                // Write in excel only unique entries and avoid duplicates where details are in comments field.
                if (docListConfEntries.Count > 0)
                {
                    string prevEntryKey = null;
                    string prevEntryComment = null;
                    int rowCount = 0;
                    Logger.Instance.LogWrite($"Logger processing {docListConfEntries.Count} entries");
                    int rowMax = docListConfEntries.Count + existingDocList.Count + 1;
                    int colMax = Enum.GetValues(typeof(DocIdentGlobals.DocIdentFieldsExcelColumns)).Length;
                    object[,] finalEntryArray = new object[rowMax, colMax];
                    List<int> newEntryRowIndices = new List<int>(); // List to store rows that are additionally added in ISMS repo that should be colored blue
                    List<int> removedEntryRowIndices = new List<int>(); // List to store rows that correspond to entries in exisiting DocList not found in ISMS repo that should be colored red
                    for (int j = 0; j < docListConfEntries.Count; j++)
                    {
                        DocIdent entry = docListConfEntries[j];
                        string entryId = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eId);
                        string entryTitle = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eTitle);
                        string entryComments = entry.GetFieldValue(DocIdentGlobals.DocIdentFields.eErrors);
                        string currentEntryKey = entryId + entryTitle;
                        if (!mapDocList.ContainsKey(entryId)) { 
                            mapDocList.Add(entryId, entry);
                        }

                        if (j == 0 || !(prevEntryKey == currentEntryKey &&
                                         !string.IsNullOrEmpty(entryComments) &&
                                         prevEntryComment == entryComments))
                        {
                            if (!string.IsNullOrEmpty(this.sUseExistingDocListEntries) &&
                                existingDocList.TryGetValue(entryId, out DocListExistingEntries existingDocListEntry))
                            {
                                AddToFinalEntryArray(finalEntryArray, rowCount, entry, existingDocListEntry);
                            }
                            else
                            {
                                AddToFinalEntryArray(finalEntryArray, rowCount, entry, null);
                                newEntryRowIndices.Add(rowCount);
                            }
                            rowCount++;
                            existingDocListEntry = null;
                        }

                        prevEntryKey = currentEntryKey;
                        prevEntryComment = entryComments;
                        entry = null;
                        // Logger.Instance.LogWrite($"Logger processed: {j}");
                    }

                    // Entries found only in exisiting doc list must be marked with RED rows
                    if (!string.IsNullOrEmpty(this.sUseExistingDocListEntries))
                    {
                        foreach (var entry in existingDocList)
                        {
                            // Not the dummy entry
                            if (!mapDocList.ContainsKey(entry.Key) && !entry.Key.Equals("AA"))
                            {
                                AddExistingEntryToFinalEntryArray(finalEntryArray, rowCount, entry.Value);
                                removedEntryRowIndices.Add(rowCount);
                                rowCount++;
                            }
                        }
                    }                    
                    mapDocList = null;
                    existingDocList = null;
                    docListConfEntries = null;
                    WriteToExcelRows(doclist, finalEntryArray, newEntryRowIndices, removedEntryRowIndices);
                    newEntryRowIndices = null;
                    removedEntryRowIndices = null;
                    finalEntryArray = null;
                }

                // write finalEntryArray to doclist sheet
                Logger.Instance.LogWrite($"Logger processed All entries");
                if (docListNonConfEntries.Count > 0) { 
                    int rowMax = docListNonConfEntries.Count + 1;
                    int colMax = Enum.GetValues(typeof(DocIdentGlobals.DocIdentFieldsExcelColumns)).Length;
                    object[,] finalNonConfEntryArray = new object[rowMax, colMax];

                    // Write non-conf entries to Excel
                    for (int k = 0; k < docListNonConfEntries.Count; k++)
                    {
                        DocIdent entry = docListNonConfEntries[k];
                        AddToFinalEntryArray(finalNonConfEntryArray, k, entry, null);
                    }
                    docListNonConfEntries = null;
                    WriteToExcelRows(misc, finalNonConfEntryArray, new List<int>(), new List<int>());
                    finalNonConfEntryArray = null;
                }
                
                
                SaveExcel(DocList, ofilepath);
            }
            catch (Exception ex)
            {
                success = false;
                Logger.Instance.LogWrite("Error: " + ex.Message);
            }
            finally
            {
                ReleaseObject(hist);
                ReleaseObject(doclist);
                ReleaseObject(misc);
                ReleaseObject(DocList);

                if (xlApp != null)
                {
                    xlApp.Quit();
                    ReleaseObject(xlApp);
                    xlApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return success;
        }

        public void SaveExcel(Excel.Workbook wb, string file)
        {
        string filepath;
            bool isDirectory = file.IndexOfAny(Path.GetInvalidPathChars()) == -1;
            if (isDirectory)
            {
                DirectoryInfo Dir = new DirectoryInfo(@file);
                filepath = Dir.FullName;
            }
            else
            {
                filepath = directory + "\\" + file;
            }
            Console.WriteLine(filepath);
            Logger.Instance.LogWrite($"Saved the excel {filepath}");

            try
            {
                
                wb.SaveAs2(filepath);
                wb.Close();
            }
            catch (Exception ex) when (ex.Message.Contains("0x800A03EC"))
            {
                //Console.WriteLine("Do you want file to be renamed ? [y/n]");
                //string rename = Console.ReadLine();
                //switch (rename)
                //{
                //    case "y":
                //        Console.WriteLine("Please insert new filename or path");
                //        string newName = Console.ReadLine();
                //        SaveExcel(wb, newName);
                //        break;
                //    case "n":
                //        break;
                //    default:
                //        Console.WriteLine("Only [y/n] valid input");
                //        SaveExcel(wb, file);
                //        break;
                //}
            }
        }

        public void WriteToExcelRows(Excel.Worksheet sheet, object[,] finalEntryArray, List<int> newEntryRowIndices, List<int> removedEntryRowIndices)
        {
            int rowMax = finalEntryArray.GetLength(0);
            int colMax = finalEntryArray.GetLength(1);

            // Define the starting row for insertion (after the second row)
            int startInsertRow = 3; // Starting from row 3 (after the second row)



            // Insert rows after the second row to accommodate the new data
            sheet.Rows[startInsertRow + ":" + (startInsertRow + rowMax - 1)].Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            // Define the range to write to (starting from the new third row)
            var startCell = sheet.Cells[startInsertRow, 1]; // Starting from A3
            var endCell = sheet.Cells[startInsertRow + rowMax - 1, colMax]; // End cell based on the size of the array

            // Create a range that matches the size of the finalEntryArray
            Excel.Range range = sheet.Range[startCell, endCell];

            // Set the values for the entire range in one go
            object[,] values = new object[rowMax, colMax];

            // Populate the values array from finalEntryArray
            for (int row = 0; row < rowMax; row++)
            {
                for (int col = 0; col < colMax; col++)
                {
                    // Ensure that numeric values requiring leading zeros are treated as text
                    // You can adjust this based on which columns you expect to contain leading zeros
                    if (finalEntryArray[row, col] is string)
                    {
                        string strValue = finalEntryArray[row, col].ToString();
                        if (int.TryParse(strValue, out _) || double.TryParse(strValue, out _))
                        {
                            // If it can be converted to a number, prepend a single quote to force Excel to treat it as text
                            values[row, col] = "'" + strValue;
                        }
                        else
                        {
                            // If it's not a numeric string, just use it as it is
                            values[row, col] = strValue;
                        }
                    }
                    else if (finalEntryArray[row, col] is int || finalEntryArray[row, col] is double)
                    {
                        // For numeric values, ensure leading zeros are preserved as text (e.g., "001")
                        values[row, col] = finalEntryArray[row, col].ToString(); // Convert numbers to string
                    }
                    else
                    {
                        values[row, col] = finalEntryArray[row, col];  // Regular values, no special treatment
                    }
                }
            }

            // Write the entire range at once
            range.Value2 = values;
            // id column should be a string always
            sheet.Columns[1].NumberFormat = "@";
            // Apply date formats to appropriate columns
            ApplyDateFormat(sheet, 20, rowMax);   // Published
            ApplyDateFormat(sheet, 25, rowMax);  // Confirm date
            ApplyDateFormat(sheet, 26, rowMax);   // Next review plan
            ApplyDateFormat(sheet, 28, rowMax);  // Changed On

            // Apply red font color to the rows in newEntryRowIndices not present in existing doclist
            // Only if existing doclist entries provided
            if (!string.IsNullOrEmpty(sUseExistingDocListEntries))
            {
                foreach (int rowIndex in newEntryRowIndices)
                {
                    Excel.Range rowRange = sheet.Range[sheet.Cells[rowIndex + 3, 1], sheet.Cells[rowIndex + 3, colMax]]; // +3 to account for row 1 and row 2 being headers
                    rowRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue); // Set font color to Blue
                }
                foreach (int rowIndex in removedEntryRowIndices)
                {
                    Excel.Range rowRange = sheet.Range[sheet.Cells[rowIndex + 3, 1], sheet.Cells[rowIndex + 3, colMax]]; // +3 to account for row 1 and row 2 being headers
                    rowRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); // Set font color to Red
                }
            }

            // Color red the last error column (before isFilled)
            Excel.Range lastColumnRange = sheet.Range[sheet.Cells[2, colMax], sheet.Cells[rowMax + 1, colMax-1]];
            lastColumnRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); // Set font color to Red

            // Remove empty rows if necessary
            RemoveEmptyRows(sheet, rowMax, colMax);

            // Apply formulas from row 2 to all other rows - gives out of memory error ths commented
            //ApplyFormulas(sheet, startInsertRow, rowMax, colMax);
        }

        private void ApplyFormulas(Excel.Worksheet sheet, int startInsertRow, int rowMax, int colMax)
        {
            // Get the range for row 2 (from column 1 to colMax)
            Excel.Range sourceRow = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, colMax]];

            // Get the target range from row 3 to the last inserted row
            Excel.Range targetRange = sheet.Range[sheet.Cells[startInsertRow, 1], sheet.Cells[startInsertRow + rowMax - 1, colMax]];

            // Copy the formulas from row 2 to the target range
            targetRange.Formula = sourceRow.Formula;

            // Optional: Release COM objects to avoid memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceRow);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(targetRange);
        }

        private void ApplyDateFormat(Excel.Worksheet sheet, int columnIndex, int rowMax)
        {
            // Apply date format to the specified column (columnIndex is 1-based, so 2 means Column B)
            Excel.Range columnRange = sheet.Range[sheet.Cells[2, columnIndex], sheet.Cells[rowMax + 1, columnIndex]]; // From row 2 to the last row
            columnRange.NumberFormat = "dd/MM/yyyy"; // You can customize this date format as needed

        }
        
        private void RemoveEmptyRows(Excel.Worksheet sheet, int rowMax, int colMax)
        {
            // Loop from the last row upwards to avoid skipping rows when deleting
            for (int row = rowMax; row >= 2; row--)
            {
                bool isEmptyRow = true;

                // Check if the row is empty (you can customize this logic as needed)
                for (int col = 1; col <= colMax; col++)
                {
                    var cellValue = sheet.Cells[row, col].Value2;
                    if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString()))
                    {
                        isEmptyRow = false;
                        break; // No need to check further, this row has data
                    }
                }

                // If the row is empty, try to delete it
                if (isEmptyRow)
                {
                    try
                    {
                        Excel.Range rowRange = sheet.Rows[row];
                        rowRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp); // Shift remaining rows up
                    }
                    catch (Exception ex)
                    {
                        // Log the error and continue if row deletion fails
                        Logger.Instance.LogWrite($"Error deleting row {row}: {ex.Message}");
                    }
                }
            }
        }
        public void AddExistingEntryToFinalEntryArray(object[,] finalEntryArray, int arrayIndex, DocListExistingEntries existingDocListEntry)
        {
            // Assuming existing entries are all aligned with the template
            // Excel Columns to enum mapping:
            int startColumnIndex = DocIdentGlobals.DocIdentConsts.DocExcelIdentFieldsConstOffsetBegin;

            // Ensure the finalEntryArray has enough rows
            if (arrayIndex >= finalEntryArray.GetLength(0))
            {
                int arglen = finalEntryArray.GetLength(0);
                throw new ArgumentOutOfRangeException(nameof(arrayIndex), "Array index is out of bounds.");
            }

            foreach (DocIdentGlobals.DocIdentFieldsExcelColumns field in Enum.GetValues(typeof(DocIdentGlobals.DocIdentFieldsExcelColumns)))
            {
                int index = (int)field;

                // Index should be within the valid range for Excel columns (assuming same template was used to generate exisiting doclist)
                string value = existingDocListEntry.GetFieldValue(field);

                // Conversion for TYPE
                if (field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColType)
                {
                    if (tabletype.TryGetValue(value, out string newValue))
                    {
                        value = newValue; // Update value if found in the dictionary
                    }
                }

                if (field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColOrgAcr)
                {
                    if (tableorg.TryGetValue(value, out string newValue))
                    {
                        // entry for eOrganization is prev column of eOrgAcr
                        finalEntryArray[arrayIndex, index - startColumnIndex - 1] = newValue;
                    }
                }

                // Write the value to the DS if it's not empty
                if (!string.IsNullOrEmpty(value))
                {
                    if (field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColPublished || field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColChangedOn)
                    {
                        DateTime parsedDate;
                        bool isDate = DateTime.TryParse(value, out parsedDate);
                        if (isDate)
                        {
                            // If it's a valid date, store it as a DateTime value (strip the time if needed)
                            finalEntryArray[arrayIndex, index - startColumnIndex] = parsedDate.Date;  // Store only the date part, no time
                        }
                        else
                        {
                            // If it's not a date, store the value as a string
                            finalEntryArray[arrayIndex, index - startColumnIndex] = value; // Write existing value to the array
                        }
                    }
                    else
                    {
                        finalEntryArray[arrayIndex, index - startColumnIndex] = value; // Use arrayIndex for row and adjusted column index
                    }
                }                   
            }
        }
        
        public void AddToFinalEntryArray(object[,] finalEntryArray, int arrayIndex, DocIdent docEntry, DocListExistingEntries existingDocListEntry)
        {
            // Excel Columns to enum mapping:
            int startColumnIndex = DocIdentGlobals.DocIdentConsts.DocExcelIdentFieldsConstOffsetBegin;

            // Ensure the finalEntryArray has enough rows
            if (arrayIndex >= finalEntryArray.GetLength(0))
            {
                int arglen = finalEntryArray.GetLength(0);
                throw new ArgumentOutOfRangeException(nameof(arrayIndex), "Array index is out of bounds.");
            }

            // Iterate through all values in the DocIdentGlobals.DocIdentFields enum
            foreach (DocIdentGlobals.DocIdentFieldsExcelColumns field in Enum.GetValues(typeof(DocIdentGlobals.DocIdentFieldsExcelColumns)))
            {
                int index = (int)field;

                // Ensure the index is within the valid range for Excel columns
                if (index >= startColumnIndex && index <= DocIdentGlobals.DocIdentConsts.DocExcelFieldsConstOffsetEnd)
                {
                    DocIdentGlobals.DocIdentFields docIdField = DocIdentGlobals.GetDocIdentFieldForExcelColumn(field);

                    // Get the value from the docEntry for the current field
                    string value = string.Empty;

                    if (docIdField != 0)
                        value = docEntry.GetFieldValue(docIdField);
                    // Conversion for TYPE
                    if (field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColType)
                    {
                        if (tabletype.TryGetValue(value, out string newValue))
                        {
                            value = newValue; // Update value if found in the dictionary
                        }
                    }

                    if (field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColOrgAcr) 
                    {
                        if (tableorg.TryGetValue(value, out string newValue))
                        {                               
                            // entry for eOrganization is prev column of eOrgAcr
                            finalEntryArray[arrayIndex, index - startColumnIndex-1] = newValue;
                        }
                    }


                    // Write the value to the DS if it's not empty
                    if (!string.IsNullOrEmpty(value))
                    {
                        if (field == DocIdentGlobals.DocIdentFieldsExcelColumns.eColPublished || field ==DocIdentGlobals.DocIdentFieldsExcelColumns.eColChangedOn)
                        {                            
                            DateTime parsedDate;
                            bool isDate = DateTime.TryParse(value, out parsedDate);
                            if (isDate)
                            {
                                // If it's a valid date, store it as a DateTime value (strip the time if needed)
                                finalEntryArray[arrayIndex, index - startColumnIndex] = parsedDate.Date;  // Store only the date part, no time
                            }
                            else
                            {
                                // If it's not a date, store the value as a string
                                finalEntryArray[arrayIndex, index - startColumnIndex] = value; // Write existing value to the array
                            }
                        } else { 
                            finalEntryArray[arrayIndex, index - startColumnIndex] = value; // Use arrayIndex for row and adjusted column index
                        }
                    }
                    else if (!string.IsNullOrEmpty(this.sUseExistingDocListEntries) && existingDocListEntry != null)
                    {
                        // If no value from docEntry, check existingDocListEntry
                        string existingValue = existingDocListEntry.GetFieldValue(field);
                        if (!string.IsNullOrEmpty(existingValue))
                        {
                            // Try to parse the existing value as a date
                            DateTime parsedDate;
                            bool isDate = DateTime.TryParse(existingValue, out parsedDate);

                            if (isDate)
                            {
                                // If it's a valid date, store it as a DateTime value (strip the time if needed)
                                finalEntryArray[arrayIndex, index - startColumnIndex] = parsedDate.Date;  // Store only the date part, no time
                            }
                            else
                            {
                                // If it's not a date, store the value as a string
                                finalEntryArray[arrayIndex, index - startColumnIndex] = existingValue; // Write existing value to the array
                            }
                        }
                    }
                }
            }
        }


    }
}
