foreach($comp in $computers){
    icm -cn $comp -ScriptBlock{
    Get-NetTCPConnection | select localaddress, localport, remoteaddress, remoteport, state, owningprocess
    } | Export-Csv -Path "$env:USERPROFILE\desktop\output\$comp`_netconnections.csv" -NoTypeInformation -Append
}