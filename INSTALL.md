# Project: DocListManager

## Install

To install the Doclist Manager application, download the DoclistManagerSetup.msi or setup.exe file available with the latest release and
make sure that you have administrator rights to install new sofware into your computer.
If you do not have this permission, contact your IT department.

To install the application, double click on the installer file and follow the onscreen 
instructions. 

The default installation folder is set to C:\Program Files\ITR\DocListManager
by default, but other path may be used, just make sure to register the path for later use.

The installation does not includes the application in the sytem path variables, which 
means that when using the application, it is necessary to navigate to the installtion
folder.

An alternative is to include the application in the PATH system variable. To do so, follow
the instructions below:

1. Wait for the installation process to finish and close the installation window.
1. Click on the start menu and type "env".
1. The "Edit the system environment variables" options should appear. Select this option
1. Click the "Environment Variables..." button in the window that appears;
1. On the lower half of the screen there is a list called "System Variables", select "Path" and clock on the "Edit..." button;
1. The "Edit environment variable" window will appear. Click on the button "New" and type in the path to the installation folder;
1. Click on the "OK" button and restart the machine to complete.

## Testing the installation

To test the installation, open PowerShell and navigate to the folder in which the
appliation was installed. Type the following comman on the PowerShell window:

```bash
.\DocListManager.exe --xmlConfigFilePath ${Env:ProgramFiles}\ITR\DocListManager\input\DirToMonitor.xml --docListTemplate ${Env:ProgramFiles}\ITR\DocListManager\input\DocList-Template.xlsm --logdir $Env:USERPROFILE\Desktop\
```

In case DocListManager was added to the system PATH, use the following command:

```bash
DocListManager.exe --xmlConfigFilePath ${Env:ProgramFiles}\ITR\DocListManager\input\DirToMonitor.xml --docListTemplate ${Env:ProgramFiles}\ITR\DocListManager\input\DocList-Template.xlsm --logdir $Env:USERPROFILE\Desktop\
```

This will create the doclist for the desired folder and create a log file in the current 
user desktop. Minimize all windows and review to content of the log file to find the newly
created document.

## Uninstall

To uninstall the application, you may execute the same file as that during installation,
select the "Uninstall" option and follow the on-screen instructions.

If the application was added to the sysem path variable, the following extra steps must
be followed to restore the PATH system variable:

1. Click on the start menu and type "env".
1. The "Edit the system environment variables" options should appear. Select this option
1. Click the "Environment Variables..." button in the window that appears;
1. On the lower half of the screen there is a list called "System Variables", select "Path" and clock on the "Edit..." button;
1. The "Edit environment variable" window will appear. 
1. Select the folder path in which the application was installed and click on the "Delete" button;
1. Click on the "OK" button and restart the machine to complete.

# Build

# Supported Operating System
  Windows 11 Pro 

# Building DocListManager Visual Studio Application with Setup Project

This document outlines the basic steps to build a Visual Studio application (in src directory) consisting of:
- `DocListManager.sln` (solution file)
- `DocListManager` (main application project)
- `DocListManagerSetup` (installer project)

---

## Prerequisites

### Install Visual Studio
- Download from: [https://visualstudio.microsoft.com](https://visualstudio.microsoft.com)
- Use Visual Studio 2022 or later

### Required Workloads
- .NET Desktop Development
- Visual Studio Installer Projects extension (Install via Extensions Manager or [Visual Studio Marketplace](https://marketplace.visualstudio.com/items?itemName=VisualStudioClient.MicrosoftVisualStudio2017InstallerProjects))

### Others
- Microsoft Excel (Desktop version)
  Version: Excel 2016 or later
- Open XML SDK (NuGet package: DocumentFormat.OpenXml)

---

## Project Structure

src (Source code repository)

   ─  DocListManager.sln (Solution)

      ─  DocListManager (Main Application Project) 

      ─ DocListManagerSetup (Installer / Setup Project)

---

## Build Steps

### 1. Open the Solution
- Launch Visual Studio
- Open `DocListManager.sln`

### 2. Set Build Configuration
- Configuration: `Release` or `Debug`
- Platform: `x86`, `x64`, or `Any CPU`

### 3. Build the Main Project
- Right-click `DocListManager` → **Build**
- Verify output in `bin\Release` or `bin\Debug`

### 4. Configure Setup Project
- Open `DocListManagerSetup`
- Add/Change project output if needed:
  - Right-click `Application Folder` → Add → Project Output → Select `Primary Output from DocListManager`

### 5. Build Setup Project
- Right-click `DocListManagerSetup` → **Build**
- Output will be in `DocListManagerSetup\Release` (or `Debug`) as `.msi` and `.exe`

---

## Optional: Include Prerequisites

1. Right-click `DocListManagerSetup` → Properties → Click **Prerequisites...**
2. Select required .NET Framework or runtime
3. Choose deployment method (web, same location, or custom path)

---

## Test the Installer

- Run the `.msi` or `setup.exe` from `Release` folder
- Ensure:
  - Installation is successful
  - All required files are included
  - Uninstall work correctly

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Setup project missing | Install "Visual Studio Installer Projects" extension |
| Build errors | Re-check references and project paths |
| Missing files in setup | Add them manually in setup project's file system editor |


