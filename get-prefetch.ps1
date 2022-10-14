foreach($comp in $computers){
    icm -cn $comp -ScriptBlock{
    gci C:\Windows\Prefetch\ | select name, creationtime, lastwritetime
    } | Export-Csv -Path "$env:USERPROFILE\desktop\output\$comp`_prefetch.csv" -NoTypeInformation -Append
}