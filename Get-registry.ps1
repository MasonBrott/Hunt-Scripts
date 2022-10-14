$keys =  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\','HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\lsass.exe', 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Security packages', 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\OSConfig\security packages'

foreach($computer in $computers){
    Invoke-Command -ComputerName $computer -ScriptBlock {
        foreach($key in $keys){
            $key = get-item $key -ErrorAction SilentlyContinue
            $reg = @{}
            $counter = 1

            foreach($i in ($key | select -ExpandProperty property)){
                $reg.add($key.Name + "_" + $counter, $i)
                $counter++
    
            }
            $reg.GetEnumerator()
        }
    } | select name, value | Export-Csv -Path "$env:USERPROFILE\desktop\output\$comp`_registry.csv" -NoTypeInformation -Append
}