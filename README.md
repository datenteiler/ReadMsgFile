# ReadMsgFile

Read Outlook MSG files with this PowerShell cmdlet without the need for Outlook. 


PowerShell binary cmdlets Read-MsgFile and Get-MsgAttachment
------------------------------------------------------------

* Read-MsgFile reads Microsoft Outlook MSG files without the need for Outlook.
* Get-MsgAttachment extracts attachments from them.


How to compile and start:
-------------------------

In PowerShell use the following commands to compile the cmdlet with .NET core. :

```
dotnet restore
dotnet publish -c Release
ipmo ./bin/Release/netstandard2.0/publish/ReadMsgFile.dll
Read-MsgFile -File sample.msg 
```

The cmdlet will show the content of the MSG file like sender and receiver, the email body etc. The output are objects, not only text.

The cmdlet is using https://github.com/Sicos1977/MSGReader and https://github.com/ironfede/openmcdf to extract the information from the MSG file. 

Installation from PowerShell Gallery (PSGallery)
------------------------------------------------

You can download the module from PSGallery in PowerShell with this command:

```Install-Module -Name ReadMsgFile```


![screenshot_20181127_221001]

