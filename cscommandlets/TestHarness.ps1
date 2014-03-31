$parentFolder = 2000
$templateId = 123456

$DebugPreference = "Continue"
$dllPath = ".\cscommandlets.dll"
Write-Debug "** Importing module"
Import-Module $dllPath

Write-Debug "** Creating AGANode object"
$newItem = New-Object -TypeName cscommandlets.Node

Write-Debug "** Trying to add without connection opened"
$newItem = Add-CSFolder -Name Tester123 -ParentID $parentFolder
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Opening connection with unencrypted password"
Open-CSConnection -Username otadmin@otds.admin -ServicesDirectory http://content.cgi.demo/cws/

Write-Debug "** Encrypting the password"
$encryptedPassword = ConvertTo-CGIEncryptedPassword
Write-Debug "Encrypted password: $encryptedPassword"

Write-Debug "** Opening connection with encrypted password"
Open-CSConnection -Username otadmin@otds.admin -password $encryptedPassword -ServicesDirectory http://content.cgi.demo/cws/

Write-Debug "** Trying to add a folder with connection opened"
$newName = "Tester"
$newName += Get-Date -format "hm"
$newItem = Add-CSFolder -Name $newName -ParentID $parentFolder
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Trying to add a folder with the same name"
$newItem = Add-CSFolder -Name $newName -ParentID $parentFolder
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Deleting the folder"
$response = Remove-CSNode -NodeID $newItem.ID
$response

Write-Debug "** Trying to add a project workspace"
$newName += " PWS"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID $parentFolder
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Trying to add a project workspace with the same name"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID $parentFolder
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Deleting the project workspace"
$response = Remove-CSNode -NodeID $newItem.ID
$response

Write-Debug "** Trying to add a project workspace and populate from template"
$newName += " from template"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID $parentFolder -TemplateID $templateId
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Deleting the project workspace from template"
$response = Remove-CSNode -NodeID $newItem.ID
$response