# Changing directories to the location that your scripts are located. (Could be changed to hardcode each script but there are 6 run from this one)
Set-Location $env:USERPROFILE\desktop\scripts\
# Loads contents of computers.txt to be used as a target list for remote commands
$computers = Get-Content $env:USERPROFILE\desktop\computers.txt

# Setting up environment. Script outputs will break if these two directories do not exist.
if (Test-Path $env:USERPROFILE\desktop\output){
    "Output path exists"
} else {
       New-Item -ItemType Directory -Path $env:USERPROFILE\desktop\output
       "Output path created"
}
if (Test-Path $env:USERPROFILE\desktop\reports){
    "Report path exists"
} else {
       New-Item -ItemType Directory -Path $env:USERPROFILE\desktop\reports
       "Report path created"
}

# Starts the scripts that actually complete the hunt tasks and produce the output.
.\get-registry.ps1
.\processes.ps1
.\Get-events.ps1
.\get-prefetch.ps1
.\get-netconnections.ps1
.\get-files.ps1

# Path to where the CSVs are located (Should be output\<computername>\)
$path = "$env:USERPROFILE\desktop\output"


$reports = "$env:USERPROFILE\desktop\reports"

# Here's where the actual report building begins after the setup is complete and the data collection finishes
# 4 nested foreach loops to parse down to each computer then each csv then each line of the csv and then each cell to be input the the new excel file.
foreach ($computer in $computers){
    $csvs = Get-ChildItem -path $path\$comp\* -Recurse -Include *.csv
    $count = $csvs.count
    Write-Host "$computer, $count CSVs"

    $outputfilename = "$reports\$computer`_report.xlsx"
    # Setting up the excel document
    $excelapp = new-object -comobject Excel.Application
    $excelapp.sheetsInNewWorkbook = $count
    $xlsx = $excelapp.Workbooks.Add()
    $sheet=1
    # Sets up a new worksheet for each csv found in the output directory
    foreach($csv in $csvs){
        $row=1
        $column=1
        $worksheet = $xlsx.Worksheets.Item($sheet)
        $worksheet.Name = $csv.Name
        $file = (Get-Content $csv)

        foreach($line in $file){
            $linecontents=$line -split ',(?!\s*\w+")'

            foreach($cell in $linecontents){
                $worksheet.Cells.Item($row,$column) = $cell
                $column++
            }
            $column = 1
            $row++
        }
        $sheet++
    }
    $xlsx.SaveAs($outputfilename)
    $excelapp.quit()
}