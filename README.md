cscommandlets
=============

PowerShell commandlets for Content Server

It's just a PowerShell wrapper to the Content Server SOAP based web services. It's written to support a few admin and SharePoint integration requirements. Current calls supported (not many at the moment):
- simple login
- create folder
- create project workspace (including from template)
- delete item

I'll be adding more as there's a need.

Testing
-------
Build the solution (it needs to be built to copy over the test script). Then run it in debug/release, which should bring up PowerShell. Change to the debug/release folder for the project and then run the TestHarness.ps1 script.

If you're implementing a new call write up your own tests (and one day I might get around to writing scripts that actually test my code, rather than best path testing).

If you're getting an error about running scripts, run `Set-ExecutionPolicy RemoteSigned`