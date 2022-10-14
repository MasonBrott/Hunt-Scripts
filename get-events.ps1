$ids = 4738,4728,4670,4768,4769,400,403,4104,4688,4624,4460,4463,4698,4697,4657
foreach ($comp in $computers){
    icm -cn $comp -ScriptBlock{
        get-winevent -FilterHashtable @{LogName='security';ID=$ids} | select timecreated, id, message
    } | Export-Csv -Path "$env:USERPROFILE\desktop\output\$comp`_events.csv" -NoTypeInformation -Append 
}