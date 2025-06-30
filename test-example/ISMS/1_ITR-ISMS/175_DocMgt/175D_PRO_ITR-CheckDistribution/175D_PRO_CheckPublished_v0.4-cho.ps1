# File: CheckPublished_v1.0.ps1
# Author: Camar Houssein
# Created on: 30/11/2023
# Description: This script checks if each itrust ISMS doc policy is published in its latest version in Internal repo


# When launching the script, it pompts to provided for some infos:
        # - The source path: 
            # ° You can provide source path with the following formats ==> "C:\Path\To\SourcePath" or you can just type enter without providing any date in which case it will take the default value, namely "N:\MG\ISMS"
        # - The shared path: 
            # ° You can provide source path with the following formats ==> "C:\Path\To\SharedPath" or you can just type enter without providing any date in which case it will take the default value, namely "N:\_INternal\ISMS"

# The script output will be displayed in the terminal window: The first table concerns published policies located in Internal. The second table lists policies that have been approved but not yet published in "N:\_INternal\ISMS".

# To execute the script:
# Just double-click on the 'launch_CheckPublished.bat' bat file at the root of the project.



Write-Host ""
Write-Host "The script output will be displayed in the terminal window: The first table concerns published policies located in Internal. The second table lists policies that have been approved but not yet published in N:\_INternal\ISMS."
Write-Host ""
Write-Host "The script output will be also saved in csv files at the root of the project: a csv file named `published-[currentDate].csv` for the published policies and another named `notPublished-[currentDate].csv` for those who have not been yet published"
Write-Host ""

$SourceRepo = Read-Host 'Enter a source folder path with the following format without the double quotes ==> "C:\Path\To\SourcePath" [Default is N:\MG\ISMS]'
$SharedRepo = Read-Host 'Enter a shared folder path with the following format without the double quotes ==> "C:\Path\To\SharedPath" [Default is N:\_INternal\ISMS]'

if ($SourceRepo -eq ""){$SourceRepo = "N:\MG\ISMS"}
if ($SharedRepo -eq ""){$SharedRepo = "N:\_INternal\ISMS"}

# Arrays to store information about found and not found files
$foundFiles = @()
$notFoundFiles = @()

$currentDate = Get-Date

$dateOfAction = Get-Date($currentDate)
# Format date for output csv file name
$dayDate = $dateOfAction.ToString("yyyyMMdd")

# Specify the path for the CSV output files
# Published
$foundCSV = "175D_REC_published-$dayDate.csv"
# Not Published
$notFoundCSV = "175D_REC_notPublished-$dayDate.csv" 

# Search for the folder named "_distributed" recursively
# $foundFolders = Get-ChildItem -Path $SourceRepo -Directory -Recurse | Where-Object { $_.Name -eq "_distributed" -and $_.FullName -notmatch '_WITHDRAWN|_old' }
$foundFolders = Get-ChildItem -Path $SourceRepo -Directory -Recurse | Where-Object { $_.Name -like "_dist*" -and $_.FullName -notmatch '_WITHDRAWN|_old' }

# Check if any matching folders were found
if ($foundFolders.Count -gt 0) {
    $foundFolders | ForEach-Object {
        $currentFolder = $_.FullName

        # Retrieve the last saved file without extension
        # $lastSavedFile = Get-ChildItem $currentFolder | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $lastSavedFile = Get-ChildItem -Path $currentFolder -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($lastSavedFile) {
            $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($lastSavedFile.Name)

            # Check if the file is present in N:\_INternal\ISMS folder
            $fileInSharedRepo = Get-ChildItem -Path $SharedRepo -Filter "$fileNameWithoutExtension.*" -Recurse | Select-Object -First 1
            if ($fileInSharedRepo) {
                # Add information about found file to the $foundFiles array
                $foundFiles += [PSCustomObject]@{
                    FileName = $fileInSharedRepo.Name
                    ModifiedDate = $fileInSharedRepo.LastWriteTime
					SharedRepoPath = $fileInSharedRepo.FullName
					SharedRepoLink = "file:///$($fileInSharedRepo.FullName -replace '\\', '/')"
                }
            } else {
                # Add information about not found file to the $notFoundFiles array
                $notFoundFiles += [PSCustomObject]@{
                    FileName = $fileNameWithoutExtension
                    Folder = $currentFolder
                    NotFoundLink = "file:///$($currentFolder -replace '\\', '/')"
                }
            }
        }
    }

    # Display the table of found files in N:\_INternal\ISMS
    if ($foundFiles.Count -gt 0) {
		Write-Host ""
        Write-Host "Found files in N:\_INternal\ISMS:"
		$foundFiles | Export-Csv -Path $foundCSV -Delimiter ";" -NoTypeInformation -Force
        $foundFiles | Format-Table
    } else {
        Write-Host "No files found in N:\_INternal\ISMS."
    }

    # Display the table of not found files in N:\_INternal\ISMS
    if ($notFoundFiles.Count -gt 0) {
		Write-Host ""
        Write-Host "Not found files in N:\_INternal\ISMS:"
        $notFoundFiles | Export-Csv -Path $notFoundCSV -Delimiter ";" -NoTypeInformation -Force
        $notFoundFiles | Format-Table		
    } else {
        Write-Host "All files found in N:\_INternal\ISMS."
    }
} else {
    Write-Host "No folders named '_distributed' found in the specified directory and its subfolders ($SourceRepo)."
}
