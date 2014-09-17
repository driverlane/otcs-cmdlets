# cscmdlets

PowerShell cmdlets for Content Server. The master branch is for 64 bit IDs (10.5), the master-32bit branch is for the older ID numbers (i.e. 9.7.1 to 10.0).

It's just a PowerShell wrapper to the Content Server SOAP based web services. It's written to support a few admin and SharePoint integration requirements. Current calls are listed [here](../../wiki/cmdlets-list).

There's a link to using piped csv values to run the cmdlets in the wiki, as well as some version history - [link](../../wiki).

Basically this project is the cscmdlets.dll. In the dll there's a snapin called, unsuprisingly, cscmdlets. Feel free to use it if you want to install the cmdlets. InstallUtil under .NET 2 for prior to Windows 8.1, .NET 4 for Windows 8.1. Personally I just use `ImportModule \path\to\dll`

If you're getting an error about running scripts, run `Set-ExecutionPolicy RemoteSigned`

The CS 10 version is in dist\CS 10. The CS 10.5 version is in dist\CS 10.5

### Testing

There's a project in the solution for testing. It's more system test than unit test, so creating unit testing one day would be nice. In the meantime please create a system test for any cmdlets you create.

### Build

Please make sure the project is using the CS 10.5 or cws service definitions before any push or merge requests.

Before build please get all the tests in the cscmdlets.tests project to pass. Then build a release version of the dll and copy it to the CS 10.5 folder in the dist folder.

To build the 32 bit version (i.e. CS 10 or les-services) point the service definitions at a CS 10 instance. Resolve the errors (basically just do a search and replace on Int64 to Int32 in CmdletsTests, Server, CSMetadata and Cmdlets files) and run the tests (possibly changing the configuration object to a CS 10 specific version). Again build with the release profile and add it to the CS 10 folder in the dist folder. Then return the service definitions to the CS 10.5 version and update any Int32 to Int64 in the files.

The NoConnection tests will only work if there's not been a connection opened prior to testing. So run them first and then ignore them. They also don't care which server you're pointing to or even if there's a server available.