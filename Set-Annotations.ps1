#Requires -Version 4
Add-Type -AssemblyName System.Windows.Forms

# Define and compile Win32Helper if it hasn't been loaded in this session
if (-not ([System.Management.Automation.PSTypeName]'Win32Helper').Type) {
    $win32Definitions = @'
    using System;
    using System.Runtime.InteropServices;

    public class Win32Helper {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
'@
    Add-Type -TypeDefinition $win32Definitions
}

If($host.name -eq 'ConsoleHost'){ # or -notmatch 'ISE'
    $logPath = "$PSScriptRoot\logs\"
}
Else{
  $logPath = "$($env:USERPROFILE)\Documents\logs\"
}

If(!(Test-Path $logPath)){
    New-Item -ItemType directory -Path $logPath | Out-Null
}

$logFile = $logPath + "LogFile_$(get-date -f ddMMyyhhmm).log"

function Write-Log{
    [CmdletBinding()]

    Param(
        [Parameter(Mandatory=$False)]
        [ValidateSet('INFO','WARN','ERROR','FATAL','DEBUG')]
        [String]
        $Level='INFO',

        [Parameter(Mandatory=$True)]
        [string]
        $Message,
    
        [Parameter(Mandatory=$False)]
        [string]
        $Path,

        [Parameter(Mandatory=$False)]
        [switch]
        $Console
    )
    
    $Stamp=(Get-Date).toString('MM/dd/yyyy HH:mm:ss')
    $Line="$Stamp $Level"+": $Message"

    If($Path){
        Try{
            Add-Content $Path -Value $Line -ErrorAction Stop
        }
        Catch{
            $ErrorMessage = $_.Exception.Message
            Write-Log -Level ERROR -Message "Error writing log: $ErrorMessage"
        }
    }
    If(!$Path -or $Console){
        Write-Output $Line
    }
}

Write-Host "




Log will be written to: $logFile

"

function Get-FileName($initialDirectory, $filterString){
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = $filterString
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function Load-TrimSDK(){

    $path = "Hewlett-Packard\HP TRIM","Hewlett-Packard\Records Manager","Hewlett Packard Enterprise\Content Manager","Micro Focus\Content Manager"

    $i = 0

    While($i -lt $path.Count){   
        $keyPath = "HKLM:SOFTWARE\" + $path[$i] + "\MSISettings"
        if (Test-Path $keyPath){
            $key = Get-ItemProperty -Path $keyPath
            Break
        }
        $i++
    }

    $sdkPath = "$($key.INSTALLDIR)TRIM.SDK.dll"

    try{
        Add-Type -Path "$($key.INSTALLDIR)TRIM.SDK.dll"
    }
    catch{
        $msgBoxInput = [System.Windows.Forms.MessageBox]::Show("Unable to find TRIM.SDK.dll at '$sdkPath'. Would you like to browse for it?",'Load SDK','YesNo','Warning')

        switch ($msgBoxInput){
            'Yes' {
                $sdkPath = Get-FileName -initialDirectory "$($env:ProgramFiles)" -filterString "|TRIM.SDK.dll"
                Try{
                    Add-Type -Path $sdkPath
                }
                Catch{
                    $ErrorMessage = $_.Exception.Message
                    Write-Log -Level ERROR -Path $logFile -Message "$ErrorMessage"
                }
            }
            'No' {
                Write-Log -Level INFO -Path $logFile -Message "Operation cancelled"
                exit
            }
        }
    } 
}

function ConvertTo-PDF {
    param($TextDocumentPath)
    Add-Type -AssemblyName System.Drawing
    $doc = New-Object System.Drawing.Printing.PrintDocument
    $doc.DocumentName = $TextDocumentPath
    $doc.PrinterSettings = New-Object System.Drawing.Printing.PrinterSettings
    $doc.PrinterSettings.PrinterName = 'Microsoft Print to PDF'
    $doc.PrinterSettings.PrintToFile = $true
    $file = [io.fileinfo]$TextDocumentPath
    $pdf = ' C:\Users\trim.svc\Desktop\test.pdf'
    $doc.PrinterSettings.PrintFileName = $pdf
    $doc.Print()
    $doc.Dispose()
}

Load-TrimSDK

$db = New-Object TRIM.SDK.Database
$db.Connect()

$flatDir = "C:\FlattenedPDFs"
$search = New-Object TRIM.SDK.TrimMainObjectSearch($db, [TRIM.SDK.BaseObjectTypes]::Record)
$search.SetSearchString("extension:tif and revision:2")

$trimProc = Get-Process -Name "TRIM" | Select-Object -First 1
[Win32Helper]::SetForegroundWindow($trimProc.MainWindowHandle)

foreach ($record in $search) {
    $recordUri = $record.uri.UriAsString
    
    # 1. Trigger Print Document via context menu shortcut
    [System.Windows.Forms.SendKeys]::SendWait("+{F10}") # Shift + F10
    Start-Sleep -Milliseconds 250
    [System.Windows.Forms.SendKeys]::SendWait("r")      # Electronic
    Start-Sleep -Milliseconds 250
    [System.Windows.Forms.SendKeys]::SendWait("d")      # Print Document
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")# Confirm Print dialog
    
    # 2. Wait for clawPDF to output the new file
    $timeout = 10
    $found = $false
    while ($timeout -gt 0) {
        Start-Sleep -Seconds 1
        $newFile = Get-ChildItem -Path $flatDir -Filter "*_TIF File.pdf" | Select-Object -First 1
        if ($newFile) {
            Rename-Item -Path $newFile.FullName -NewName "$recordUri.pdf"
            Write-Host "Flattened and saved: $recordUri.pdf"
            $found = $true
            break
        }
        $timeout--
    }

    # 3. Move selection to the next record in the CM grid
    [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
    Start-Sleep -Milliseconds 300
}
