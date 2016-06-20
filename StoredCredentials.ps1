
install-module 


$credentials = Get-StoredCredential -Type GENERIC -AsPsCredential:$false 
$credentials = $credentials | where { $_.UserName -like '*' -and $_.Type -eq 'GENERIC'} | select -Property UserName, TargetName, Type, TargetAlias, CredentialBlob
$credentials | FL *
