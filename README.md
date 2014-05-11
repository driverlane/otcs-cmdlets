cscmdlets
=============

PowerShell cmdlets for Content Server. The master branch is for 64 bit IDs (10.5), the master-32bit branch is for the older ID numbers (i.e. 9.7.1 to 10.0).

It's just a PowerShell wrapper to the Content Server SOAP based web services. It's written to support a few admin and SharePoint integration requirements. Current calls are listed [here](../../wiki/cmdlets-list).

I'll be adding more as there's a need.

Build
-------
Basically this project is the cscmdlets.dll. In the dll there's a snapin called, unsuprisingly, cscmdlets. Feel free to use it if you want to install the cmdlets. InstallUtil under .NET 2 for prior to Windows 8.1, .NET 4 for Windows 8.1. Personally I just use `ImportModule \path\to\dll`

If you're getting an error about running scripts, run `Set-ExecutionPolicy RemoteSigned`

Testing
-------
There's a project in the solution for testing. It does the basics, but i'm not that happy with it. It's more end to end test than unit test, so will have to create proper unit testing.

Build 32 bit branch
-------
```
git fetch origin   
git pull
git checkout master-32bit  
git reset --hard master  
git merge -s ours origin/master-32bit
```
Then:

- replace all Int64 with Int32
- uncomment/comment the globals
- repoint the service references at the CS 10 definitions
- run the tests and if they pass add all, commit and push to the origin