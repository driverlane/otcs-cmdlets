$DebugPreference = "Continue"
$dllPath = ".\CGIPSTools.dll"
Write-Debug "** Importing module"
Import-Module $dllPath

Write-Debug "** Creating AGANode object"
$newItem = New-Object -TypeName CGIPSTools.AGANode

Write-Debug "** Trying to add without connection opened"
$newItem = Add-CSFolder -Name Tester123 -ParentID 118594
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Opening connection with unencrypted password"
Open-CSConnection -Username otadmin@otds.admin -password p@ssw0rd -ServicesDirectory http://content.logica.demo/les-services/

Write-Debug "** Encrypting the password"
$encryptedPassword = ConvertTo-CGIEncryptedPassword -Password p@ssw0rd
Write-Debug "Encrypted password: $encryptedPassword"

Write-Debug "** Opening connection with encrypted password"
Open-CSConnection -Username otadmin@otds.admin -password $encryptedPassword -ServicesDirectory http://content.logica.demo/les-services/

Write-Debug "** Trying to add a folder with connection opened"
$newName = "Tester"
$newName += Get-Date -format "hm"
$newItem = Add-CSFolder -Name $newName -ParentID 118594
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Trying to add a folder with the same name"
$newItem = Add-CSFolder -Name $newName -ParentID 118594
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Trying to add a project workspace"
$newName += " PWS"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 118594
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Trying to add a project workspace with the same name"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 118594
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "** Trying to add a project workspace and populate from template"
$newName += " from template"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 118594 -TemplateID 486237
$newItem.ID
$newItem.NodeValue
$newItem.Message
