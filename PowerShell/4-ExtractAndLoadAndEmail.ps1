#DMT Automation Example 4

#Extract Data From DB -> CSV File -> Load in with DMT -> Send Email on Completion

Import-module "sqlps" -DisableNameChecking
$DMTPath = "C:\Epicor\ERP10\LocalClients\ERP10\DMT.exe"
$User = "epicor"
$Pass = "epicor"

$Source = "C:\Temp\Product.csv"
$completeLog = $source + ".CompleteLog.txt"

Write-Output "Extracting Data"
#Extract Data
Invoke-SqlCmd -Query "SELECT TOP 50 
                        'EPIC06' As Company,
                        ProductNumber As PartNum,
                        Name As PartDescription,
                        (Case when [MakeFlag] = 0 then 'P' else 'M' END) as TypeCode
                        FROM [AdventureWorks2012].[Production].[Product]" | Export-Csv -NoType $Source


Write-Output "Extracted : " Import-Csv $Source | Measure-Object | % {$_.Count}

Write-Output "Loading Data " $Source 

#Load Data
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Add -Update -Import Part -Source $Source "

#Check Results
select-string -Path $completeLog -Pattern "Records:\D*(\d+\/\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
select-string -Path $completeLog -Pattern "Errors:\D*(\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }

Send-MailMessage -From DMT@YourCompany.com -Subject "Job Done" -To administrator@yourcompany.com -Attachments $completeLog -Body "DMT Completed" -SmtpServer localhost

