#DMT Automation Example 2

#Extract Data From From BAQ -> CSV File -> Load in with DMT

$DMTPath = "C:\Epicor\ERP10\LocalClients\ERP10\DMT.exe"
$User = "epicor"
$Pass = "epicor"

$Source = "C:\Temp\CustomerList.csv"
$completeLog = $source + ".CompleteLog.txt"

Write-Output "Extracting Data via BAQ $(get-date)"

Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Export -BAQ CustomerDetails -Target $Source -NoUI -ConfigValue=ERP10"
Write-Output "Loading Data $(get-date) " $Source 

#Load Data
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Add -Update -Import Customer -Source $Source "

#Check Results
select-string -Path $completeLog -Pattern "Records:\D*(\d+\/\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
select-string -Path $completeLog -Pattern "Errors:\D*(\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }