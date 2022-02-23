<#
.SYNOPSIS
  Name: Check-IfMicrosoftOfficeRunning.ps1
  The purpose of this script is to check if Microsoft Office is running and return $true or $false accordingly.
  
.DESCRIPTION
  See synopsis

.NOTES
    Release Date: 2022-02-23T09:36
    Last Updated: 2022-02-23T10:03
   
    Author: Luke Nichols
    Github link: https://github.com/jlukenichols/CheckIfMicrosoftOfficeRunning

.EXAMPLE
    Just run the script without parameters, it's not designed to be called like a function
#>

#-------------------------- Set any initial values --------------------------

#Clear the console for easier PowerShell ISE debugging
Clear-Host

[array]$ArrayOfExecutables = "MSACESS" #Microsoft Access 2016
$ArrayOfExecutables += "EXCEL"` #Microsoft Excel 2016
$ArrayOfExecutables += "MSPUB"` #Microsoft Publisher 2016
$ArrayOfExecutables += "ONENOTE"` #Microsoft OneNote 2016
$ArrayOfExecutables += "ONENOTEM"` #Microsoft OneNote 2016
$ArrayOfExecutables += "POWERPNT" #Microsoft PowerPoint 2016
$ArrayOfExecutables += "VISIO" #Microsoft Visio 2016
$ArrayOfExecutables += "WINWORD" #Microsoft Word 2016
$ArrayOfExecutables += "WINPROJ" #Microsoft Project 2016
$ArrayOfExecutables += "SETLANG" #Microsoft Office 15 Language Preferences
$ArrayOfExecutables += "SPREADSHEETCOMPARE" #Microsoft Office component (32 bit)
$ArrayOfExecutables += "DATABASECOMPARE" #Microsoft Office component (32 bit)
$ArrayOfExecutables += "MSOUC" #Microsoft Office Upload Center

#-------------------------- End setting initial values --------------------------

#-------------------------- Begin defining functions --------------------------

function Check-IfProcessesRunning {
    Param (
        #ArrayOfProcesses should contain an array of process names WITHOUT trailing .exe
        [array]$ArrayOfProcesses
    )
    :loopThroughProcesses foreach ($Process in $ArrayOfProcesses) {
        if (Get-Process $Process -ErrorAction SilentlyContinue) {
            $ProcessFound = $true
            $RunningProcess = $Process
            break loopThroughProcesses
        }
    }
    if ($ProcessFound -eq $true) {
        return $true
        Write-Output "Detected running MS Office process $RunningProcess"
    } else {
        return $false
        Write-Output "No running MS Office processes found."
    }
    
}

#-------------------------- End defining functions --------------------------

#-------------------------- Start main script body --------------------------

if ((Check-IfProcessesRunning -ArrayOfProcesses $ArrayOfExecutables) -eq $true) {
    #There were Microsoft Office processes running. Exit with error code 1 to indicate that it is NOT safe to continue with the Microsoft Office upgrade procedure.
    exit 1
} elseif ((Check-IfProcessesRunning -ArrayOfProcesses $ArrayOfExecutables) -eq $false) {
    #There were no Microsoft Office processes running. Exit with error code 0 to indicate that it is safe to continue with the Microsoft Office upgrade procedure.
    exit 0
}

#-------------------------- End main script body --------------------------