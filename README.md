cscmdlets
=============

PowerShell cmdlets for Content Server. The master branch is for 64 bit IDs (10.5), the master-32bit branch is for the older ID numbers (i.e. 9.7.1 to 10.0).

It's just a PowerShell wrapper to the Content Server SOAP based web services. It's written to support a few admin and SharePoint integration requirements. Current calls are listed [here](../../wiki/cmdlets-list).

There's a link to using piped csv values to run the cmdlets in the wiki, as well as some version history - [link](../../wiki).

Build
-------
Basically this project is the cscmdlets.dll. In the dll there's a snapin called, unsuprisingly, cscmdlets. Feel free to use it if you want to install the cmdlets. InstallUtil under .NET 2 for prior to Windows 8.1, .NET 4 for Windows 8.1. Personally I just use `ImportModule \path\to\dll`

If you're getting an error about running scripts, run `Set-ExecutionPolicy RemoteSigned`

Testing
-------
There's a project in the solution for testing. It's more system test than unit test, so creating unit testing one day would be nice. In the meantime please create a system test for any cmdlets you create.