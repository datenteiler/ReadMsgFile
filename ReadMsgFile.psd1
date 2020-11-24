@{
RootModule = 'ReadMsgFile.dll'
ModuleVersion = '1.1'
GUID = '7f86d981-137d-4d72-a334-5303b95b2475'
Author = 'Christian Imhorst'
CompanyName = 'www.datenteiler.de'
Copyright = '(c) 2018 Christian Imhorst. Some rights reserved.'
Description = 'Read Outlook MSG files with this PowerShell cmdlet without the need for Outlook'
PowerShellVersion = '5.0'
CmdletsToExport=@("Read-MsgFile","Get-MsgAttachment")
PrivateData = @{
    PSData = @{
        Tags = @()
        LicenseUri = 'https://github.com/datenteiler/ReadMsgFile/blob/master/LICENSE'
        ProjectUri = 'https://github.com/datenteiler/ReadMsgFile'
        # IconUri = ''
        # ReleaseNotes = ''
    } # End of PSData hashtable
} # End of PrivateData hashtable
}
