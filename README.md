# ReadMsgFile

Read Outlook MSG files with this PowerShell cmdlet without the need for Outlook.


PowerShell binary cmdlets Read-MsgFile and Get-MsgAttachment
------------------------------------------------------------

* Read-MsgFile reads Microsoft Outlook MSG files without the need for Outlook.
* Get-MsgAttachment extracts attachments from them.


How to compile and start:
-------------------------

In PowerShell use the following commands to compile the cmdlet with .NET core:

```
dotnet restore
dotnet publish -c Release
ipmo ./bin/Release/net7.0/publish/ReadMsgFile.dll
Read-MsgFile -File sample.msg 
```

The cmdlet will show the content of the MSG file like sender and receiver, the email body etc. The output are objects, not only text.

The cmdlet is using https://github.com/Sicos1977/MSGReader and https://github.com/ironfede/openmcdf to extract the information from the MSG file. 

Installation from PowerShell Gallery (PSGallery)
------------------------------------------------

You can download the module from PSGallery in PowerShell with this command:

```Install-Module -Name ReadMsgFile```

This works for Linux ...

![read-msgfile](https://user-images.githubusercontent.com/3180008/49112709-c68c5500-f293-11e8-839e-26b8df7b1248.png)

... and Windows ...

![screenshot_20181127_221001](https://user-images.githubusercontent.com/3180008/49111855-9a6fd480-f291-11e8-8899-b2b0ef9a53da.png)

... and should works for MacOS X, too.
