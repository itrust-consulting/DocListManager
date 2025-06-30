using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;

namespace DocListManager
{
    /// <summary>
    /// Deserializes input XML file (DirToMonitor) and creates a class object.
    /// </summary>
    public class XMLtoObject
    {
        /// <summary>
        /// Definition of the class object Directory
        /// </summary>
        public class Directory
        {
            public bool active { get; set; }
            public string ISMSWorkDir { get; set; }
            public string docListSavePath { get; set; }
            public string docListName { get; set; }
            public bool docListOverwrite { get; set; }
            public string activityName { get; set; }
            public string activityAcronymn { get; set; }
            public string ISMSPublishDir { get; set; }
            public string useExistingDocListEntries { get; set; }
            public string templateDoclistSheet { get; set; }
            public string templateMiscSheet { get; set; }
            public string templateMappingSheet { get; set; }
            public string existingDocListSheet { get; set; }

            [XmlElement("excludeList")]
            public ExcludeList Exclusions { get; set; }
        }

        /// <summary>
        /// Class to represent each excludeName element
        /// </summary>
        public class ExcludeName
        {
            [XmlAttribute("postFixWildCard")]
            public string PostFixWildCard { get; set; }  // Optional attribute

            [XmlText]
            public string Name { get; set; }  // The text content of the excludeName element
        }

        /// <summary>
        /// Class to represent the excludeList element
        /// </summary>
        public class ExcludeList
        {
            [XmlElement("excludeName")]
            public List<ExcludeName> ExcludeNames { get; set; }
        }

        /// <summary>
        /// Definition of the class object DirectoryList (used to match the XML structure)
        /// </summary>
        [XmlRoot("DirList")]
        public class DirectoryList
        {
            [XmlElement("Directory")]
            public List<Directory> Directories { get; set; }
        }

        /// <summary>
        /// Deserialization of the XML input file.
        /// </summary>
        /// <returns>List of all directories and their attributes provided in the XML input file.</returns>
        public static List<Directory> DeserializeObject(string szFilePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(DirectoryList));

            using (StreamReader streamReader = new StreamReader(szFilePath))
            {
                DirectoryList deserializedList = (DirectoryList)serializer.Deserialize(streamReader);
                return deserializedList.Directories;
            }
        }
    }
}
