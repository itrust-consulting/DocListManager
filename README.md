# About

This tool is designed to monitor directories and generate Excel documents based on a provided configuration.
It leverages an XML configuration file to specify directories and settings, uses an excel template to format the output, 
and logs its activities to a specified file. The tool has been developed as a C# application and is open for exploration. 
The tool is currently designed and developed to automatically handle generation of DocList for ISMS systems that are currently 
follow a specific structure however it can be tailored for specific use cases by modifying the source code.
The tool also allows to create a DocList based on an existing doclist which was followed using the tools conventions described in the user guide.

# Prerequisites, Build and Install
For the prerequisites, build and installation guide refer [Prerequisites, Build and Installation guide](./INSTALL.md)

# Usage Command

Below is a short explanation of the tool's arguments:
### `--xmlConfigFilePath`
**Description:** Specifies the path to the XML configuration file.

**Example:** `_input/DirToMonitor.xml`

**Usage:** This XML file contains the configuration for directory monitoring, including which directories to scan, file exclusions, and other settings. The tool uses this configuration to determine where to look for files and how to process them.
The schema for the XML file is specified in section 3.3 in the detailed [User Guide](./user-guide/5DCU_STA_CyFort-DocListManager-UserGuide_v1.0.pdf)

### `--docListTemplate`

**Description:** Specifies the path to the Excel template file.

**Example:** `_input/DocListTemplate.xlsm`

**Usage:** The Excel template defines the structure and formatting of the output document. The tool will populate this template with data from the monitored directories. Ensure that the template file is correctly set up with the necessary sheets and formatting.
**Note :** The default template can be customized in terms of design, but it is important to maintain the same sheet names and column structure as those in the sample template provided with the tool. This ensures that the tool can properly populate and process the data according to its expected format.
The _input folder provides an [English](_input/DocListTemplate.xlsm) and [French](_input/DocListTemplate-FR.xlsm) version of the Template.

### `--logdir`

**Description:** Specifies the path to the log file where the tool's operations and any issues encountered will be recorded.

**Example:** `log.txt`

**Usage:** The log file will capture details about the tool’s execution, including any errors, warnings, and general activity. This helps in troubleshooting and reviewing the tool's performance.

## Summary

- **`--xmlConfigFilePath`**: Path to the XML configuration file.
- **`--docListTemplate`**: Path to the Excel template file.
- **`--logdir`**: Path to the log file.
Ensure that all file paths are correct and accessible to facilitate smooth operation of the tool.

# Basic Example

Follow two basic [examples](./test-example/README.md) provided at ./test-example. 

# User Guide
[Detailed User Guide](./user-guide/5DCU_STA_CyFort-DocListManager-UserGuide_v1.0.pdf)

# License

Copyright © itrust consulting. All rights reserved. Licensed under the GNU Affero General Public License (AGPL) v3.0.

# Acknowledgment
This tool was co-funded by the Ministry of Economy and Foreign Trade of Luxembourg, within the project Cloud Cybersecurity Fortress of Open Resources and Tools for Resilience (CyFORT).

# Contact
For more information about the project, contact us at dev@itrust.lu.

