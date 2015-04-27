# cscmdlets

PowerShell cmdlets for Content Server.

These cmdlets wrap the Content Server SOAP and REST based web services. It's written to support a few admin and SharePoint integration requirements. Current calls are listed [here](../../wiki/cmdlets-list).

There's some very basic instructions on using piped csv values to run the cmdlets in the wiki - [link](../../wiki). If that doesn't do it, the googles are your friend (and probably stack overflow in the results of the googles).

Version history here - [link](../../releases)

Basically this project is the cscmdlets.dll, which is in the dist folder. In the dll there's a snapin called, unsuprisingly, cscmdlets. Feel free to use it if you want to install the cmdlets. InstallUtil under .NET 2 for prior to Windows 8.1, .NET 4 for Windows 8.1. Personally I just copy the dll to the folder where I'm writing the script and use `Import-Module .\cscmdlets.dll` to access the cmdlets.

If you're getting an error about running scripts, try `Set-ExecutionPolicy RemoteSigned`

### Testing  
The CmdletsTests project has a corresponding unit/system test for any cmdlets that are mentioned in the wiki. Please create one for any cmdlets you want published.

The NoConnection tests will only work if there's not been a connection opened prior to testing. So run them first and then ignore them. They also don't care which version of server you're pointing to or even if there's a server available.

### Build  
Please make sure the project is using the CS 10.5 or cws service definitions before any push or merge requests.

Before build please get all the tests in the cscmdlets.tests project to pass. Then build a release version of the dll and copy it to the CS 10.5 folder in the dist folder.

To build the 32 bit version (i.e. CS 10 or les-services) point the service definitions at a CS 10 instance. Resolve the errors (basically just do a search and replace on Int64 to Int32 in CmdletsTests, SoapApi, RestApi, CSMetadata and Cmdlets files) and run the tests (change the TestGlobals class so the CS 10 version is current). Once they've passed build a release version and add it to the CS 10 folder in the dist folder. Then return the service definitions to the CS 10.5 version and update any Int32 to Int64 in the files.