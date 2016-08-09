#region Vault-SDK Intellisense
#[System.Reflection.Assembly]::LoadFrom("C:\Program Files (x86)\Autodesk\Autodesk Vault 2015 R2 SDK\bin\Autodesk.Connectivity.WebServices.dll")
#$cred = New-Object Autodesk.Connectivity.WebServicesTools.UserPasswordCredentials("localhost", "MKOE_CAx", "Administrator", "")
#$vault = New-Object Autodesk.Connectivity.WebServicesTools.WebServiceManager($cred)
#endregion

$folderId=$vaultContext.CurrentSelectionSet[0].Id
$vaultContext.ForceRefresh = $true
$dialog = $dsCommands.GetCreateFolderDialog($folderId)
$xamlFile = New-Object CreateObject.WPF.XamlFile "testxaml", "%ProgramData%\Autodesk\Vault 2015 R2\Extensions\DataStandard\Vault\Configuration\Folder.xaml"
$dialog.XamlFile = $xamlFile

$result = $dialog.Execute()
$dsDiag.Trace($result)

if($result)
{
	#new folder can be found in $dialog.CurrentFolder
	$folder = $vault.DocumentService.GetFolderById($folderId)
	$path=$folder.FullName+"/"+$dialog.CurrentFolder.Name
	
	$ParentFolder = $vault.DocumentService.GetFolderByPath($path)
	$_ret = $vault.DocumentService.AddFolder("CAD", $ParentFolder.Id, $false)
	
	#add property values to the last created folder by the API method UpdateFolderProperties 
	# see SDK help for parameters required
	
	$_FldIDs = @() #create the list of folders to update
	$_FldIDs += $_ret.Id
	
	$_FldProIds = @()#add property definitions to update
	$_FldPropVals = @() #add list of values
	
	$PropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
	$propDefIds = @()
	$PropDefs | ForEach-Object {
		$propDefIds += $_.Id
	}
	$_mPropID = $propDefs | Where-Object { $_.DispName -eq "Beschreibung"} #note you might replace by UIString.Value
	If ($_mPropID) 
	{
		$_FldProIds += $_mPropID.Id
		$_FldPropVals += "MyTestDescription"
	}
	#alternatively use the system name of the property to find the definition id
	$_PropSysNames = @()
	$_PropSysNames += "Title"
	$_TitlePropDefId = $vault.PropertyService.FindPropertyDefinitionsBySystemNames("FLDR", $_PropSysNames)
	If ($_TitlePropDefId)	
	{
		$_FldProIds += $_TitlePropDefId.SyncRoot[0].Id
		$_FldPropVals += "MyTestTitle"
	}
	#the API methods offers a list of overloads; PowerShell is not easily able to call the appropriate overload
	#to make this easier the method's call is wrapped in a dll and we call the overload 1 (of 2)
	[System.Reflection.Assembly]::LoadFrom('C:\ProgramData\Autodesk\Vault 2015 R2\Extensions\DataStandard\Vault\addinVault\VDSUtils.dll')
	$_mVltHelpers = New-Object VDSUtils.VltHelpers
	$_Result = $_mVltHelpers.UpdateFolderProp1($vault, $_FldIDs, $_FldProIds, $_FldPropVals)
	
	[System.Reflection.Assembly]::LoadFrom("C:\Program Files\Autodesk\Vault Professional 2015 R2\Explorer\Autodesk.Connectivity.Explorer.Extensibility.dll")
	$selectionId = [Autodesk.Connectivity.Explorer.Extensibility.SelectionTypeId]::Folder
	$location = New-Object Autodesk.Connectivity.Explorer.Extensibility.LocationContext $selectionId, $path
	$vaultContext.GoToLocation = $location
}