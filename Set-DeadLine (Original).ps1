$WSUSServer = 'SERVERNAME'
$Group = 'CLUSTER_SERVERS'
[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WSUSServer,$False)
$ComputerTargetGroups = $wsus.GetComputerTargetGroups()

$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
$UpdateScope.ApprovedStates = 'NotApproved'
$UpdateScope.IncludedInstallationStates = 'NotInstalled'

##Classifications
#Get all Classifications for specific Classifications
$updateClassifications = $wsus.GetUpdateClassifications() | Where {
  $_.Title -Match "Critical Updates|Updates|Security Updates"
}
$UpdateScope.Classifications.AddRange($updateClassifications)

$ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
$TargetGroup = $ComputerTargetGroups | Where {$_.Name -eq $Group}
[void]$computerscope.ComputerTargetGroups.Add($TargetGroup)

$Clients = $WSUS.GetComputerTargets($computerscope)
$Updates = $Clients.GetUpdateInstallationInfoPerUpdate($UpdateScope) | Select -Unique -ExpandProperty UpdateID
ForEach ($Item in $Updates) {    
    $Update = $wsus.GetUpdate($Item)
    Write-Verbose "Install $($Update.title) on $($TargetGroup.name)" -Verbose
    $Update.Approve('Install',$TargetGroup)
}