The ComputerReport_V2 script will call 6 other scripts to call out to a list of remote computers and grab various information. The information that those scripts grab is based off of what
was needed for one hunt phase of a mission and should be updated to reflect what specific things you are looking for.
After the data collection is complete ComputerReport_V2 will combine the csv outputs from each computer in to an excel file (one for each computer). Each csv that it finds will be one
worksheet within the excel doc.
This is best suited for a small network or a small subset of computers due to the amount of files being pulled back and being created.
