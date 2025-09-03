using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Reflection.Metadata.Ecma335;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Security.Cryptography;
using Shell32;
using System.Reflection.Metadata;
using System.Runtime.ConstrainedExecution;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Drawing;
using System.Net.NetworkInformation;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace DocListManager
{
    public static class VersionClassifier
    {
        public enum DocumentStatus
        {
            NotRequired = 0,        // "0.0 - Not required"
            ToDo = 1,               // "0.1 - To do"
            InitialDraft = 2,       // "0.1 - InitialDraft"
            Draft = 3,              // "0.2 - Draft"
            StructureAndContentValidated = 4,  // "0.3 - Structure and content validated"
            FinalDraft = 5,         // "0.4 - Final draft"
            ToApprove = 6,          // "1.0 - To approve"
            ForRevision = 7,        // "1.x - For revision"
            Current = 8,            // "1.x.x - Current"
            Unknown = 9,            // "? - Unknown"
            Revised = 10,           // "1.x - Revised"
            Distributed = 11,       // "1.0 - Distributed"
            Withdrawn = 12,         // "9 - Withdrawn"
            IncorrectDist = 13      // "? - IncorrectDist"
        }
        // The dictionary to map DocumentStatus to its description
        private static readonly Dictionary<DocumentStatus, string> statusDescriptions = new()
        {
            { DocumentStatus.NotRequired, "0.0 - Not required" }, // Not tool gen
            { DocumentStatus.ToDo, "0.1 - To do" }, // Not tool gen
            { DocumentStatus.InitialDraft, "0.1 - InitialDraft" }, // In ISMS working dir
            { DocumentStatus.Draft, "0.2 - Draft" }, // In ISMS working dir
            { DocumentStatus.StructureAndContentValidated, "0.3 - Structure and content validated" }, // In ISMS working dir
            { DocumentStatus.FinalDraft, "0.4 - Final draft" }, // In ISMS working dir
            { DocumentStatus.ToApprove, "1.0 - To approve" }, // In ISMS working dir
            { DocumentStatus.ForRevision, "1.x - For revision" }, // In ISMS working dir
            { DocumentStatus.Current, "1.x.x - Current" }, // In ISMS working dir
            { DocumentStatus.Unknown, "? - Unknown" }, // In ISMS working dir
            { DocumentStatus.Revised, "1.x - Revised" }, // In distr
            { DocumentStatus.Distributed, "1.0 - Distributed" }, // In distr
            { DocumentStatus.Withdrawn, "9 - Withdrawn" }, // Not tool gen
            { DocumentStatus.IncorrectDist, "? - IncorrectDist" } // In distr
        };
        /// The dictionary to map DocumentStatus to its description
        public static string GetDescription(this DocumentStatus status)
        {
            return statusDescriptions.TryGetValue(status, out var description) ? description : string.Empty;
        }

        /// @Description: Determines the document status based on the version string and whether it is in the distribution folder.         
        ///  The version string is expected to be in the format "vX.Y.Z" where X, Y, and Z are integers.
        ///  The method returns a string description of the document status based on the version and folder location.
        /// @ARG1 : "version": The version string of the document.
        /// @ARG2: "inDistributionFolder"Indicates if the document is in the distribution folder.
        /// Returns: A string description of the document status.
               
        public static string GetDocumentStatus(string version, bool inDistributionFolder)
        {
            // Remove the leading 'v' if present
            if (version.StartsWith("v"))
            {
                version = version.Substring(1);
            }

            // Split the version string by dots
            var parts = version.Split('.');

            if (!inDistributionFolder)
            {
                // Parse the version numbers
                if (parts.Length >= 2 && int.TryParse(parts[0], out int major) && int.TryParse(parts[1], out int minor))
                {
                    if (major == 0)
                    {
                        if (minor == 1)
                        {
                            return GetDescription(DocumentStatus.InitialDraft); // All subversions of 0.1 are InitialDraft
                        }
                        else if (minor == 2)
                        {
                            return GetDescription(DocumentStatus.Draft); // All subversions of 0.2 and higher are Draft
                        }
                        else if (minor == 3)
                        {
                            return GetDescription(DocumentStatus.StructureAndContentValidated); // All subversions of 0.3 and higher are StructureAndContentValidated
                        }
                        else if (minor == 4)
                        {
                            return GetDescription(DocumentStatus.FinalDraft); // All subversions of 0.4 and higher are Final Draft
                        }
                        else
                        {
                            return GetDescription(DocumentStatus.FinalDraft); // All subversions of 0.5,0.6, 0.7.. and higher are Final Draft
                        }
                    }
                    else if (major == 1)
                    {
                        if (minor == 0)
                        {
                            if (parts.Length == 2)
                            {
                                return GetDescription(DocumentStatus.ToApprove); // Version 1.1, 1.0 etc is To approve
                            }
                            else
                            {
                                return GetDescription(DocumentStatus.Current); // Version 1.0.x (with any patch version) is Current
                            }
                        }
                        else if (minor == 1)
                        {
                            if (parts.Length == 2)
                            {
                                return GetDescription(DocumentStatus.ForRevision); // Version 1.1 is For revision
                            }
                            else if (parts.Length > 2)
                            {
                                return GetDescription(DocumentStatus.Current); // Any 1.1.x version is Current
                            }
                        }
                        else if (parts.Length > 2)
                        {
                            return GetDescription(DocumentStatus.Current); // Any other 1.x.x version is Current
                        }
                    }
                    else if (major > 1)
                    {
                        if (parts.Length == 2)
                        {
                            return GetDescription(DocumentStatus.ForRevision); // 2.0, 2.1, 2.2, etc. are For Revision
                        }
                        else if (parts.Length > 2)
                        {
                            return GetDescription(DocumentStatus.Current); // Any 2.x.x version is Current
                        }
                    }
                }


                // Return Unknown if no conditions match
                return GetDescription(DocumentStatus.Unknown);
            }
            else
            {
                // Handle the distributed folder case
                if (parts.Length >= 2 && int.TryParse(parts[0], out int major) && int.TryParse(parts[1], out int minor))
                {
                    if (major == 1)
                    {
                        if (minor == 0)
                        {
                            return GetDescription(DocumentStatus.Distributed); // Version 1.0 in distributed folder
                        }
                        else if (parts.Length == 2)
                        {
                            return GetDescription(DocumentStatus.Revised); // 1.1, 1.2, 1.3, etc. are Revised
                        }
                        else
                        {
                            return GetDescription(DocumentStatus.IncorrectDist); // Anything else is an error in distributed
                        }
                    }
                    else if (parts.Length == 2)
                    {
                        return GetDescription(DocumentStatus.Revised); // Other versions (e.g., 2.x) are Revised in distributed folder
                    }
                    else
                    {
                        return GetDescription(DocumentStatus.IncorrectDist); // Anything else is error in distributed folder
                    }
                }
                else
                {
                    return GetDescription(DocumentStatus.IncorrectDist); // Anything else is an error in distributed folder
                }
            }
        }
    }

    /// Class to handle file properties and metadata extraction
        class FileProp
    {
        // Properties
        private string fileFullPathName;
        private FileInfo fileInfo;

        // Constructor accepting a full file path
        public FileProp(string szFullPathName)
        {
            fileFullPathName = szFullPathName;
            fileInfo = new FileInfo(szFullPathName);
        }
        // 
        public FileProp() { }

        // Method to check if the file is a shortcut based on its extension
        public bool IsShortcut()
        {
            if (!fileInfo.Exists) { return false; }

            // Check if the file has a .lnk extension
            return fileInfo.Extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase);
        }

        // Method to resolve the shortcut target using IWshRuntimeLibrary
        public bool ResolveShortcutTarget()
        {
            try
            {
                var shell = new IWshRuntimeLibrary.WshShell();
                var shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(fileFullPathName);

                string targetPath = shortcut.TargetPath;

                // Check if the target path exists
                if (File.Exists(targetPath) || Directory.Exists(targetPath))
                {
                    return true;
                }
                else
                {
                    // Target does not exist, handle accordingly (return null or an error)
                    return false; // Or you can throw an exception or log an error as needed
                }
            }
            catch (Exception)
            {
                return false; // If there's an error resolving, return null
            }
        }
        // Method to get the file name from the full path
        public string getFileName()
        {
            string szFileName = "";
            try
            {
                // Get Last modification time of the file
                // Create a FileInfo object
                 // Check if the file exists before deleting
                if (fileInfo.Exists)
                {
                    szFileName = System.IO.Path.GetFileName(fileFullPathName);
                }
                else
                {
                    Logger.Instance.LogWrite($"File does not exist. {fileFullPathName}");
                }
            }
            catch (Exception ex) { Logger.Instance.LogWrite($"Error: {ex.Message}");  Logger.Instance.Log($"Error: {ex.Message}", Console.Out); }
            return szFileName;
        }
        // Method to get the file extension from the full path
        public string getFileExtension()
        {
            string szFileExtension = "";
            try
            {                
                FileInfo fileInfo = new FileInfo(fileFullPathName);
                if (fileInfo.Exists)
                {
                    szFileExtension = System.IO.Path.GetExtension(fileFullPathName).ToLower();
                    
                }
                else
                {
                    Logger.Instance.LogWrite($"File does not have a valid extension. {fileFullPathName}");
                }
            }
            catch (Exception ex) { Logger.Instance.LogWrite($"Error: {ex.Message}"); Logger.Instance.Log($"Error: {ex.Message}", Console.Out); }
            return szFileExtension;
        }
        // Method to get the file author from the document properties
        public string getFileAuthor()
        {
            string szFileAuthor = "";
            try
            {
                // Check if the file is a Word document (.docx or .doc)
                if (fileFullPathName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || fileFullPathName.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                {
                    // Open the Word document in read-only mode
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileFullPathName, false))
                    {
                        // Get the document's Creator (Author) from the WordprocessingDocument properties
                        szFileAuthor = wordDoc.PackageProperties.Creator;
                    }
                }
                // Check if the file is an Excel document (.xlsx or .xlsm)
                else if (fileFullPathName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || fileFullPathName.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                {
                    // Open the Excel document in read-only mode
                    using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(fileFullPathName, false))
                    {
                        // Get the document's LastAuthor from the CoreFilePropertiesPart
                        var coreProperties = excelDoc.PackageProperties.Creator;
                        //var authorElement = coreProperties.Elements<DocumentFormat.OpenXml.CoreProperties.LastAuthor>().FirstOrDefault();

                        if (coreProperties != null)
                        {
                            szFileAuthor = coreProperties;
                        }
                        
                    }
                }                
            }
            catch (Exception ex)
            {
                Logger.Instance.LogWrite($"Error: {ex.Message}");
                Logger.Instance.Log($"Error: {ex.Message}", Console.Out);
                szFileAuthor = "Error retrieving author.";
            }

            return szFileAuthor;
        }


    //  Method to get the directory of the file     
    // This method returns the directory of the file as a string.
    // It checks if the file exists before attempting to get the directory.
    // If the file does not exist, it logs an error message and returns an empty string.     
    public string getISMSWorkDir()
        {
            string szFileDir = "";
            try
            {
                // Get Last modification time of the file
                // Create a FileInfo object
                FileInfo fileInfo = new FileInfo(fileFullPathName);
                // Check if the file exists before deleting
                if (fileInfo.Exists)
                {
                    szFileDir = System.IO.Path.GetDirectoryName(fileFullPathName);
                }
                else
                {
                    Logger.Instance.LogWrite($"File does not exist {fileFullPathName}");
                }
            }
            catch (Exception ex) { Logger.Instance.LogWrite($"Error: {ex.Message}"); }
            return szFileDir;
        }
        // Method to get the parent folder ID from the file path
        public string getISMSParentFolderID()
        {
            // Get the current directory of the file
            DirectoryInfo currentDir = fileInfo.Directory;

            // Check if the current directory matches _dir* or _appro*
            if (currentDir != null &&
                (currentDir.Name.StartsWith("_dist") || currentDir.Name.StartsWith("_appr")))
            {
                // Go up one level to get the grandparent directory
                currentDir = currentDir.Parent;
            }

            // If there's a valid parent directory, proceed to get the folder ID
            if (currentDir != null)
            {
                // Get the name of the last folder in the current directory path
                string parentFolderName = currentDir.Name;

                // Split the folder name by "_"
                string[] parts = parentFolderName.Split('_');

                // Return the first part if it exists
                return parts.Length > 1 ? parts[0] : string.Empty;
            }

            return string.Empty; // Return empty if no valid parent directory found
        }
        // Method to get the classification property from the document properties
        public string getClassificationIfPossible()
        {         
            string classification = getDocumentProperty(fileFullPathName, getFileExtension(), DocListManager.FileProp.eDocumentProperty.Classification);
            return classification;
        }
        // Method to get the subject property from the document properties
        public string getSubjectIfPossible()
        {
            string subject = getDocumentProperty(fileFullPathName, getFileExtension(), DocListManager.FileProp.eDocumentProperty.Subject);
            int parenthesisIndex = subject.IndexOf("(");
            if (parenthesisIndex > -1)
            {
                subject = subject.Substring(0, parenthesisIndex).Trim();
            }
            return subject;
        }
        //  Method to check if the parent directory is a distribution folder
        public bool checkParentDirIsDistributed(string fileFullPathName)
        {
            bool isDistr = false;
            try
            {
                DirectoryInfo parentDir = Directory.GetParent(fileFullPathName);
                string parentDirName = parentDir.Name;
                if (parentDirName.StartsWith("_dist"))
                {
                    isDistr = true;
                }
            }
            catch (Exception ex) { Logger.Instance.LogWrite($"Error: {ex.Message}"); }
            return isDistr;
        }
        //  Method to check if the parent directory is an approved folder
        public bool checkParentDirIsApproved(string fileFullPathName)
        {
            bool isApproved= false;
            try
            {
                DirectoryInfo parentDir = Directory.GetParent(fileFullPathName);
                string parentDirName = parentDir.Name;
                if (parentDirName.StartsWith("_appr"))
                {
                    isApproved = true;
                }
            }
            catch (Exception ex) { Logger.Instance.LogWrite($"Error: {ex.Message}"); }
            return isApproved;
        }

        // Method to get the last modification time of the file as a string
        public string getLastModTimeOfFileAsString()
        {
            string lastModTime = "";
            try
            {
                if (fileFullPathName != "")
                {
                    // Get Last modification time of the file
                    // Create a FileInfo object
                    FileInfo fileInfo = new FileInfo(fileFullPathName);
                    // Check if the file exists before deleting
                    if (fileInfo.Exists)
                    {
                        DateTime lastWriteTime = fileInfo.LastWriteTime;
                        lastModTime = lastWriteTime.ToString("dd/MM/yyyy");
                    }
                    else
                    {
                        Logger.Instance.LogWrite($"File does not exist. {fileFullPathName}");
                    }

                }
            } catch (Exception ex) { Logger.Instance.LogWrite($"Error: {ex.Message}"); }
            return lastModTime;
        }
        //      Method to extract the activity name from the domain string
        public enum eDocumentProperty
        {
            Title,
            Subject,
            Classification,
            Version,
            Status
        }
        // describe getDocumentProperty method            
        //          @Description: Retrieves the specified document property from a Word document (.docx) or Excel document (.xlsx).
        //          @ARG1 : "filePath": The full path of the document.
        //          @ARG2 : "extension": The file extension of the document (e.g., ".docx" or ".xlsx").
        //          @ARG3 : "attribute": The document property to retrieve (e.g., Title, Subject, Classification, Version, Status).
        //          Returns: The value of the specified document property as a string.
        //          If the property is not found or the file type is unsupported, an empty string is returned.
        //          If the file is not found, an error message is logged.
        //          If the file is not a Word or Excel document, a warning message is logged.   
        public string getDocumentProperty(string filePath, string extension, eDocumentProperty attribute)
        {
            string propertyValue = "";

            if (extension == ".docx")
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    switch (attribute)
                    {
                        case eDocumentProperty.Title:
                            propertyValue = wordDoc.PackageProperties.Title;
                            break;
                        case eDocumentProperty.Subject:
                            propertyValue = wordDoc.PackageProperties.Subject;
                            break;
                        case eDocumentProperty.Classification:
                            // Internal etc
                            propertyValue = wordDoc.PackageProperties.Description;
                            break;
                        case eDocumentProperty.Version:
                            propertyValue = wordDoc.PackageProperties.Version;
                            break;
                        case eDocumentProperty.Status:
                            propertyValue = wordDoc.PackageProperties.ContentStatus;
                            break;
                        default:
                            Logger.Instance.LogWrite($"Invalid attribute: {attribute}");
                            break;
                    }

                    // Log all properties (optional)
                    //Logger.Instance.LogWrite($"File {filePath} Title = {wordDoc.PackageProperties.Title}, Subject = {wordDoc.PackageProperties.Subject}, Description = {wordDoc.PackageProperties.Description}, Version = {wordDoc.PackageProperties.Version}, Status = {wordDoc.PackageProperties.ContentStatus}");
                }
            }
            else // TODO see how to retrieve meta data from Excel - Currently metadata is not stored example classification. this will need to be changed.
            {
                Logger.Instance.LogWrite($"Warning: Unsupported file extension: {extension} of {filePath}");
            }

            return propertyValue ?? ""; // Ensure null values are handled properly
        }

    }

    // describe DocIdent class
    //          @Description: Represents a document identifier with various fields and their values.
    public class DocIdent
    {
        // Dictionary to hold field values
        private readonly Dictionary<DocIdentGlobals.DocIdentFields, string> fieldValues;

        // Constructor accepting a dictionary
        public DocIdent(Dictionary<DocIdentGlobals.DocIdentFields, string> docEntryIdentifiers)
        {
            if (docEntryIdentifiers.Count != Enum.GetValues(typeof(DocIdentGlobals.DocIdentFields)).Length)
            {
                throw new ArgumentException("The dictionary must contain exactly 13 elements.");
            }

            fieldValues = new Dictionary<DocIdentGlobals.DocIdentFields, string>(docEntryIdentifiers);
        }

        // Method to access the dictionary values
        public string GetFieldValue(DocIdentGlobals.DocIdentFields field)
        {
            return fieldValues.TryGetValue(field, out var value) ? value : string.Empty;
        }

        // Setter method to modify the value for a given field
        public void SetFieldValue(DocIdentGlobals.DocIdentFields field, string value)
        {
            if (fieldValues.ContainsKey(field))
            {
                fieldValues[field] = value;  // Update the existing value
            }
            else
            {
                throw new ArgumentException("Field does not exist in the dictionary.");
            }
        }
    }
    // describe DocIdentGlobals class
    public class DocIdentGlobals
    {
        // Constants
        public static class DocIdentConsts
        {
            public const int DocExcelIdentFieldsConstOffsetBegin = 1; // A constant integer
            public const int DocExcelFieldsConstOffsetEnd = 32;  // A constant integer
        }

        private static Dictionary<DocIdentFieldsExcelColumns, DocIdentFields> excelFieldsDict;
        private static Dictionary<DocIdentFields, DocIdentFieldsExcelColumns> docIdentFieldDict;

        // Enum for DocIdentFields
        public enum DocIdentFields
        {
            eId = 1,
            eDomain = 2,
            eType = 3,
            eTitle = 4,
            eVersion = 5,
            eExtension = 6,
            eFilename = 7,
            eStatus = 8,
            eFolder = 9,
            eAuthor = 10,
            eChangeDate = 11,
            eChangeBy = 12,
            eErrors = 13,
            eClassification = 14,
            eParentFolderID = 15,
            eAcronymn = 16,// Ignore in Excel
            eComments = 17, // Comments only in existing doclist
            ePublished = 18 // Date of publishing the latest distributed file
        }

        // Enum for Excel columns mapping
        public enum DocIdentFieldsExcelColumns
        {
            eColId = 1,  // "Id."
            eColOldId = 2,  // "Old Id"
            eColI = 3,  // "i"
            eColJ = 4,  // "j"
            eColK = 5,  // "k"
            eColL = 6,  // "l"
            eColType = 7,  // "Type"
            eColTitle = 8,  // "Title"
            eColAcr = 9,  // "Acr."
            eColVersionOpt = 10, // "Ver. (opt)"
            eColExtension = 11, // "Ext."
            eColFilename = 12, // "Filename"
            eColOrganisation = 13, // "Organisation"
            eColOrgAcr = 14, // "OrgAcr"
            eColReference = 15, // "Ref."
            eColStatus = 16, // "Status"
            eColClassification = 17, // "Classif."
            eColManager = 18, // "Manager"
            eColOwner = 19, // "Owner"
            eColPublished = 20, // "Published"
            eColFolder = 21, // "Folder"
            eColLastCC = 22, // "last CC"
            eColNextReviewer = 23, // "Next reviewer"
            eColTicketOrDeadline = 24, // "Ticket or deadline"
            eColConfirmDate = 25, // "Confirm date"
            eColNextReviewPlan = 26, // "Next review plan"
            eColMonthToNextRevision = 27, // "Month to next revision"
            eColChangedOn = 28, // "Changed on"
            eColChangedBy = 29, // "By"
            eColComments = 30, // "Comments"
            eColPendingAction = 31,  // "Pending action"
            eColErrors = 32, // Additional Column from DocList only expected in Template
            eColIsFilled = 33 // Filed for additional isFIlled column 
        }

        // Method to map DocIdentFields to Excel columns
        // Method to get the corresponding Excel column for a given DocIdentFields
        public static DocIdentFieldsExcelColumns GetExcelColumnForDocIdentField(DocIdentFields docField)
        {
            var fieldMappings = GetFieldMappings();

            if (fieldMappings.TryGetValue(docField, out var excelColumn))
            {
                return excelColumn; // Return the corresponding Excel column
            }
            else
            {
                return (DocIdentFieldsExcelColumns)docField;
                //throw new ArgumentException($"No Excel column mapping found for DocIdentField: {docField}");
            }
        }

        public static DocIdentFields GetDocIdentFieldForExcelColumn(DocIdentFieldsExcelColumns docExcelField)
        {
            var fieldMappings = GetExcelFieldMappings();

            if (fieldMappings.TryGetValue(docExcelField, out var docIdenField))
            {
                return docIdenField; // Return the corresponding Excel column
            }
            return 0;
        }

        // Method to map DocIdentFields to Excel columns
        public static Dictionary<DocIdentFields, DocIdentFieldsExcelColumns> GetFieldMappings()
        {
            if (docIdentFieldDict == null)
            {
                docIdentFieldDict = new Dictionary<DocIdentFields, DocIdentFieldsExcelColumns>
                {
                    { DocIdentFields.eId, DocIdentFieldsExcelColumns.eColId },              // Id <-> eColId
                    { DocIdentFields.eDomain, DocIdentFieldsExcelColumns.eColOrgAcr }, // Domain <-> Organisation
                    { DocIdentFields.eType, DocIdentFieldsExcelColumns.eColType },           // Type <-> Type
                    { DocIdentFields.eAcronymn, DocIdentFieldsExcelColumns.eColAcr },        // Acronym -> Acronymn 
                    { DocIdentFields.eTitle, DocIdentFieldsExcelColumns.eColTitle },         // Title <-> Title
                    { DocIdentFields.eVersion, DocIdentFieldsExcelColumns.eColVersionOpt },  // Version <-> Ver. (opt)
                    { DocIdentFields.eExtension, DocIdentFieldsExcelColumns.eColExtension }, // Extension <-> Ext.
                    { DocIdentFields.eFilename, DocIdentFieldsExcelColumns.eColFilename },   // Filename <-> Filename
                    { DocIdentFields.eStatus, DocIdentFieldsExcelColumns.eColStatus },       // Status <-> Status
                    { DocIdentFields.eFolder, DocIdentFieldsExcelColumns.eColFolder },       // Folder <-> Folder
                    { DocIdentFields.eAuthor, DocIdentFieldsExcelColumns.eColOwner },        // Author <-> Owner
                    { DocIdentFields.eChangeDate, DocIdentFieldsExcelColumns.eColChangedOn },// Change Date <-> Changed on
                    { DocIdentFields.eChangeBy, DocIdentFieldsExcelColumns.eColChangedBy },  // Change By <-> By
                    { DocIdentFields.eErrors, DocIdentFieldsExcelColumns.eColErrors },   // Errors <-> Errors
                    { DocIdentFields.eClassification, DocIdentFieldsExcelColumns.eColClassification }, // Classification if read from file
                    { DocIdentFields.eComments, DocIdentFieldsExcelColumns.eColComments }, // Comments <-> Comments
                    { DocIdentFields.ePublished, DocIdentFieldsExcelColumns.eColPublished } // Publishing
                    // eParentFolderID is ignored as per your request
                };
            }
            return docIdentFieldDict;
        }

        public static Dictionary<DocIdentFieldsExcelColumns, DocIdentFields> GetExcelFieldMappings()
        {
            if (excelFieldsDict == null)
            {
                excelFieldsDict = new Dictionary<DocIdentFieldsExcelColumns, DocIdentFields>
                {
                    { DocIdentFieldsExcelColumns.eColId, DocIdentFields.eId },              // eColId <-> Id
                    { DocIdentFieldsExcelColumns.eColOrgAcr, DocIdentFields.eDomain },      // Organisation <-> Domain
                    { DocIdentFieldsExcelColumns.eColType, DocIdentFields.eType },          // Type <-> Type
                    { DocIdentFieldsExcelColumns.eColAcr, DocIdentFields.eAcronymn },          // Title <-> Title
                    { DocIdentFieldsExcelColumns.eColTitle, DocIdentFields.eTitle },          // Title <-> Title
                    { DocIdentFieldsExcelColumns.eColVersionOpt, DocIdentFields.eVersion },  // Ver. (opt) <-> Version
                    { DocIdentFieldsExcelColumns.eColExtension, DocIdentFields.eExtension }, // Ext. <-> Extension
                    { DocIdentFieldsExcelColumns.eColFilename, DocIdentFields.eFilename },   // Filename <-> Filename
                    { DocIdentFieldsExcelColumns.eColStatus, DocIdentFields.eStatus },       // Status <-> Status
                    { DocIdentFieldsExcelColumns.eColFolder, DocIdentFields.eFolder },       // Folder <-> Folder
                    { DocIdentFieldsExcelColumns.eColOwner, DocIdentFields.eAuthor },      // Owner <-> Author
                    { DocIdentFieldsExcelColumns.eColChangedOn, DocIdentFields.eChangeDate }, // Changed 
                    { DocIdentFieldsExcelColumns.eColComments, DocIdentFields.eComments },     // Comments
                    { DocIdentFieldsExcelColumns.eColErrors, DocIdentFields.eErrors },   // Errors
                    { DocIdentFieldsExcelColumns.eColPublished, DocIdentFields.ePublished },     // Published date
                    { DocIdentFieldsExcelColumns.eColClassification, DocIdentFields.eClassification }     // Classification of document

                };
            }
            return excelFieldsDict;
        }
    }

    public class DocListExistingEntries
    {
        private readonly Dictionary<DocIdentGlobals.DocIdentFieldsExcelColumns, string> fieldValues;

        // Constructor accepting a dictionary
        public DocListExistingEntries(Dictionary<DocIdentGlobals.DocIdentFieldsExcelColumns, string> docEntryIdentifiers)
        {
            if (docEntryIdentifiers.Count != Enum.GetValues(typeof(DocIdentGlobals.DocIdentFieldsExcelColumns)).Length)
            {
                throw new ArgumentException("The dictionary must contain exactly 13 elements.");
            }

            fieldValues = new Dictionary<DocIdentGlobals.DocIdentFieldsExcelColumns, string>(docEntryIdentifiers);
        }

        // Method to access the dictionary values
        public string GetFieldValue(DocIdentGlobals.DocIdentFieldsExcelColumns field)
        {
            return fieldValues.TryGetValue(field, out var value) ? value : string.Empty;
        }

        // Setter method to modify the value for a given field
        public void SetFieldValue(DocIdentGlobals.DocIdentFieldsExcelColumns field, string value)
        {
            if (fieldValues.ContainsKey(field))
            {
                fieldValues[field] = value;  // Update the existing value
            }
            else
            {
                throw new ArgumentException("Field does not exist in the dictionary.");
            }
        }
    }


    class CheckNC
    {
        //Check if filename conforms with naming convention:
        // ID_Domain_Type_TitleAcr_Version(-Kürzel).Extension, e.g. R001_PM_CryptoCeVerif_v1.0.0-hfr.xlsx
        // Domain can be extracted from folder. abbr. (-[a-z]{3})?
        private static string pattern = @"[A-Za-z0-9]+_[A-Z]+_[^_]+_v\d\.\d+(\.\d+)?(-[^\.]+)?\.\S+";
        private static string smallpattern = @"(_)?[A-Z0-9]{4}_";
        public static bool CheckConformity(string file)
        {
            string filename = System.IO.Path.GetFileName(file);
            Match match = Regex.Match(filename, pattern);
            
            return match.Success;
        }

        public static bool CheckSmallConformity(string file)
        {
            string filename = System.IO.Path.GetFileName(file);
            Match match = Regex.Match(filename, smallpattern);
            bool dmatch;
            if (match.Success) { dmatch = Regex.Match(match.Value, @"[A-z]\d+").Success; }
            else dmatch = false;
            return dmatch;
        }

        public static Dictionary<DocIdentGlobals.DocIdentFields, string> GetIdentifiers(string file, bool Conformity, bool inDistributionFolder)
        {
            FileProp pFileProp;
            try
            {
                pFileProp = new FileProp(file);
            }
            catch (Exception ex)
            {
                // Log or handle the exception as necessary
                Logger.Instance.LogWrite($"Failed to initialize FileProp: {ex.Message}");
                return new Dictionary<DocIdentGlobals.DocIdentFields, string>(); // Return an empty dictionary
            }

            string filename = pFileProp.getFileName();
            string[] temp = filename.Split("_");

            // Initialize the identifiers dictionary with default values
            var identifiers = new Dictionary<DocIdentGlobals.DocIdentFields, string>
            {
                { DocIdentGlobals.DocIdentFields.eId, string.Empty },
                { DocIdentGlobals.DocIdentFields.eDomain, string.Empty },
                { DocIdentGlobals.DocIdentFields.eType, string.Empty },
                { DocIdentGlobals.DocIdentFields.eTitle, pFileProp.getSubjectIfPossible() },
                { DocIdentGlobals.DocIdentFields.eAcronymn, string.Empty },
                { DocIdentGlobals.DocIdentFields.eVersion, string.Empty },
                { DocIdentGlobals.DocIdentFields.eExtension, pFileProp.getFileExtension() },
                { DocIdentGlobals.DocIdentFields.eFilename, filename },
                { DocIdentGlobals.DocIdentFields.eStatus, string.Empty },
                { DocIdentGlobals.DocIdentFields.eFolder, pFileProp.getISMSWorkDir() },
                { DocIdentGlobals.DocIdentFields.eAuthor, pFileProp.getFileAuthor() },
                { DocIdentGlobals.DocIdentFields.eChangeDate, pFileProp.getLastModTimeOfFileAsString() },
                { DocIdentGlobals.DocIdentFields.eChangeBy, string.Empty },
                { DocIdentGlobals.DocIdentFields.eComments, string.Empty },
                { DocIdentGlobals.DocIdentFields.eParentFolderID, pFileProp.getISMSParentFolderID() },
                { DocIdentGlobals.DocIdentFields.eClassification, pFileProp.getClassificationIfPossible() },
                { DocIdentGlobals.DocIdentFields.eErrors, string.Empty},
                { DocIdentGlobals.DocIdentFields.ePublished, string.Empty}
            };


            if (Conformity)
            {
                // Ensure that temp has the expected number of elements before accessing them
                if (temp.Length > 0)
                {
                    identifiers[DocIdentGlobals.DocIdentFields.eId] = temp.Length > 0 ? temp[0] : string.Empty; // id
                    identifiers[DocIdentGlobals.DocIdentFields.eDomain] = temp.Length > 2 ? ExtractactivityName(temp[2]) : "PS"; // domain
                    identifiers[DocIdentGlobals.DocIdentFields.eType] = temp.Length > 1 ? temp[1] : string.Empty; // Type
                    identifiers[DocIdentGlobals.DocIdentFields.eAcronymn] = temp.Length > 2 ? temp[2] : string.Empty; // Title
                    identifiers[DocIdentGlobals.DocIdentFields.eVersion] = temp.Length > 3 ? ExtractVersion(temp[3]) : string.Empty; // version
                    identifiers[DocIdentGlobals.DocIdentFields.eFilename] = filename; // filename

                    if (inDistributionFolder)
                    {
                        // Time stamp of file is the publishing date
                        identifiers[DocIdentGlobals.DocIdentFields.ePublished] = pFileProp.getLastModTimeOfFileAsString();
                    }

                    // Match and extract additional details
                    if (temp.Length > 3)
                    {
                        string tempPart = temp[3];
                        Match versionMatch = Regex.Match(tempPart, @"v\d+(\.\d+)+");
                        Match abbrMatch = Regex.Match(tempPart, @"-[a-z]{3}");

                        if (versionMatch.Success)
                        {
                            identifiers[DocIdentGlobals.DocIdentFields.eVersion] = versionMatch.Value;
                            identifiers[DocIdentGlobals.DocIdentFields.eStatus] = VersionClassifier.GetDocumentStatus(versionMatch.Value, inDistributionFolder);
                        }

                        if (abbrMatch.Success)
                        {
                            identifiers[DocIdentGlobals.DocIdentFields.eChangeBy] = abbrMatch.Value;
                        }
                    }
                }
            }
            else
            {
                // Handle non-conforming cases
                identifiers[DocIdentGlobals.DocIdentFields.eStatus] = ""; // status
            }
            return identifiers;
        }


        // Gets the ID of the directory path
        public static string GetDirId(string dirPath)
        {
            string identifier = "";
            // Extract the folder name from the full directory path
            string dirName = System.IO.Path.GetFileName(dirPath);
            string[] parts = dirName.Split("_");

            if (parts.Length > 0)
            {
                identifier = parts[0];                  
            }

            return identifier;
        }


        private static string ExtractVersion(string input)
        {
            Match versionMatch = Regex.Match(input, @"v\d+(\.\d+)+");
            return versionMatch.Success ? versionMatch.Value : string.Empty;
        }

        private static string ExtractactivityName(string input)
        {
            Match activityNameMatch = Regex.Match(input, @"^(.*?)-");
            return activityNameMatch.Success ? activityNameMatch.Groups[1].Value : "PS";
        }
           
        public static string FindAndVerifyFile(string filePath, string directoryPath, string expectedChecksum)
        {
            string szVerificationStr =string.Empty;
            // Get the file name from the path
            string fileName = System.IO.Path.GetFileName(filePath);

            // Search for the file in the directory
            string[] files = Directory.GetFiles(directoryPath, fileName, SearchOption.AllDirectories);

            if (files.Length == 0)
            {
                szVerificationStr = $"File '{fileName}' not found in published {directoryPath}.";
                return szVerificationStr;
            }

            foreach (string foundFile in files)
            {
                // Compute checksum of the found file
                string foundFileChecksum = ComputeFileChecksum(foundFile);

                // Compare the checksums
                if (foundFileChecksum == expectedChecksum)
                {
                    return szVerificationStr; // File found and checksum matched
                } else
                {
                    szVerificationStr = $"File '{fileName}' found in published {directoryPath} but checksum does not match.";
                    return szVerificationStr;
                }
            }

            szVerificationStr = $"File '{fileName}' not found in published {directoryPath}.";
            return szVerificationStr;
        }

        // Computes the SHA-256 checksum of a file at the specified path.
        public static string ComputeFileChecksum(string filePath)
        {
            using (var hashAlgorithm = SHA256.Create())
            {
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hashBytes = hashAlgorithm.ComputeHash(stream);
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }
        }
    }
    
}
