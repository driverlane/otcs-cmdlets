cscommandlets
=============

PowerShell commandlets for Content Server

It's just a PowerShell wrapper to the Content Server SOAP based web services. It's written to support a few admin and SharePoint integration requirements. Current calls are:

### Open-CSConnection
Simple Content Server based login, no OTDS or single sign-in support.

Parameters:
- Username **mandatory**
- Password **mandatory**
- ServicesDirectory, e.g. http://server.domain/cws/ **mandatory**

### Add-CSFolder
Creates a folder. Just support for naming at the moment.

Parameters:
- Name **mandatory**
- ParentID **mandatory**

### Add-CSProjectWorkspace
Creates a project workspace.

Parameters:
- Name **mandatory**
- ParentID **mandatory**
- TemplateID optional

### Add-CSUser
Creates a user

Parameters:
- Login **mandatory**
- DepartmentGroupID **mandatory** 1001 is DefaultGroup
- Password optional
- FirstName optional
- MiddleName optional
- LastName optional
- Email optional
- Fax optional
- OfficeLocation optional
- Phone optional
- Title optional

### Remove-CSUser
Deletes a user. Can't find a disable function in the API.

Parameters:
- UserID **mandatory**

### Remove-CSNode
Deletes most Content Server objects

Parameters:
- NodeID **mandatory**

### ConvertTo-EncryptedPassword
Outputs an encrypted version of a password when you want to store your password in a script

Parameters:
- Password **mandatory**

I'll be adding more as there's a need.

Build
-------
Basically the core of the project is the cscommandlets.dll. I haven't created snapins or anything clever yet, so just load it using `ImportModule \path\to\dll`

If you're getting an error about running scripts, run `Set-ExecutionPolicy RemoteSigned`

Testing
-------
It's a bit sketchy at the moment. Build the solution (it needs to be built to copy over the test script). Then run it in debug/release, which should bring up PowerShell. Change to the debug/release folder for the project and then run the TestHarness.ps1 script.

If you're implementing a new call write up your own tests (and one day I might get around to writing scripts that actually test my code, rather than best path testing).