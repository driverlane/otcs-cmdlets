$DebugPreference = "Continue"
$dllPath = ".\CGIPSTools.dll"
Write-Debug "Importing module"
Import-Module $dllPath

Write-Debug "Creating AGANode object"
$newItem = New-Object -TypeName CGIPSTools.AGANode

Write-Debug "Trying to add folder without connection opened"
$newItem = Add-CSFolder -Name Tester123 -ParentID 207812
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "Opening connection with unencrypted password"
Open-CSConnection -Username MagnetIntegration@MMGLibrary -password "V:M'7])=TK)$%kp" -ServicesDirectory http://qmmglibrary.myintranet.local/les-services/

#Write-Debug "Encrypting the password"
#$encryptedPassword = ConvertTo-CGIEncryptedPassword -Password OpenText1
#Write-Debug "Encrypted password: $encryptedPassword"

#Write-Debug "Opening connection with encrypted password"
#Open-CSConnection -Username otadmin@otds.admin -password $encryptedPassword -ServicesDirectory http://qmmglibrary.myintranet.local/les-services/

Write-Debug "Trying to add folder with connection opened"
$newName = "Tester"
$newName += Get-Date -format "hm"
$newItem = Add-CSFolder -Name $newName -ParentID 207812
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "Trying to add folder with the same name"
$newItem = Add-CSFolder -Name $newName -ParentID 207812
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "Trying to add project workspace"
$newName += " PWS"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 207812
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "Trying to add project workspace with the same name"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 207812
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "Trying to add project workspace with it template"
$newName += " from it template"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 207812 -TemplateID 208450
$newItem.ID
$newItem.NodeValue
$newItem.Message

Write-Debug "Trying to add project workspace with m&a template"
$newName += " from ma template"
$newItem = Add-CSProjectWorkspace -Name $newName -ParentID 207812 -TemplateID 207806
$newItem.ID
$newItem.NodeValue
$newItem.Message