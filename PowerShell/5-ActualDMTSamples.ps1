#Composed by Rick Bird
#Aligned Solutions Consulting, LLC.
#rick@getaligned.solutions
#www.getaligned.solutions
#######################
#Extract Data From From BAQ -> CSV File -> Load in with DMT -> Email results to Socialcast
#Common variables for all DMT Loads
$DMTPath = "E:\Epicor\ERP10\LocalClients\E10Live\DMT.exe"
$User = "epicordmt"
$Pass = "password"
$datestring = (Get-Date -Format yyyyMMdd-HH)

########################
#Qty Adjust Bin 1SALESS#
########################
#Set Source & Complete Log
$Source = "E:\Epicor\EpicorData\DMT\AdjustOut1SALESS_$datestring.csv"
$completeLog = $source + ".CompleteLog.txt"

#Extract via BAQ AdjustOut1SALESS
Write-Output "Extracting Data via BAQ $(get-date)"
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Export -BAQ AdjustOut1SALESS -Target $Source -Company CMP01 -NoUI -ConfigValue=E10Live"

#Load Data
Write-Output "Loading Data $(get-date)" $Source 
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Add -Import `"Quantity Adjustment`" -Source $Source "

#Check Results
select-string -Path $completeLog -Pattern "Records:\D*(\d+\/\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
select-string -Path $completeLog -Pattern "Errors:\D*(\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
Add-Content $completeLog "`r`n#DMTQtyAdj"

#Email Results to Socialcast
$emailBody = (Get-Content $completeLog -Raw| out-string)
Send-MailMessage -From epicor@yourcompany.com -Subject "Epicor DMT 1SALESS Qty Adj Complete" -To share@socialcast.com -Body $emailBody -Attachments $completeLog -SmtpServer exchange.yourcompany.local

##########################
#Lock Qty on Printed Jobs#
##########################
#Set Source & Complete Log
$Source = "E:\Epicor\EpicorData\DMT\JobsPrinted_$datestring.csv"
$completeLog = $source + ".CompleteLog.txt"

#Extract Data via BAQ JobsPrinted
Write-Output "Extracting Data via BAQ $(get-date)"
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Export -BAQ JobsPrinted -Target $Source -Company CMP01 -NoUI -ConfigValue=E10Live"

#Load Data
Write-Output "Loading Data $(get-date)" $Source
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Update -Import `"Job Header`" -Source $Source "

#Check Results
select-string -Path $completeLog -Pattern "Records:\D*(\d+\/\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
select-string -Path $completeLog -Pattern "Errors:\D*(\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
Add-Content $completeLog "`r`n#JobPrintedQtyLock"

#Email Results to Socialcast
$emailBody = (Get-Content $completeLog -Raw| out-string)
Send-MailMessage -From epicor@yourcompany.com -Subject "Epicor DMT Jobs Printed Qty Locked Complete" -To share@socialcast.com -Body $emailBody -Attachments $completeLog -SmtpServer exchange.yourcompany.local

###########################
#Correct ShipHead Invoiced#
###########################
#Set Source & Complete Log
$Source = "E:\Epicor\EpicorData\DMT\PackHdInvdCorr_$datestring.csv"
$completeLog = $source + ".CompleteLog.txt"

#Extract via BAQ PackHdInvdCorrection
Write-Output "Extracting Data via BAQ $(get-date)"
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Export -BAQ PackHdInvdCorrection -Target $Source -Company CMP01 -NoUI -ConfigValue=E10Live"


#Load Data
Write-Output "Loading Data $(get-date)" $Source 
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Update -Import `"Ship Head`" -Source $Source "

#Check Results
select-string -Path $completeLog -Pattern "Records:\D*(\d+\/\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
select-string -Path $completeLog -Pattern "Errors:\D*(\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
Add-Content $completeLog "`r`n#PackHdInvdCorr"

#Email Results to Socialcast
$emailBody = (Get-Content $completeLog -Raw| out-string)
Send-MailMessage -From epicor@yourcompany.com -Subject "Epicor DMT Ship Head Corrections Complete" -To share@socialcast.com -Body $emailBody -Attachments $completeLog -SmtpServer exchange.yourcompany.local

#############################
#SEPERATE FILE - EVERY 3 HRS#
#############################
#################################
#Update OrderHed UD Order Status#
#################################
#Extract Data From From BAQ -> CSV File -> Load in with DMT

$DMTPath = "E:\Epicor\ERP10\LocalClients\E10Live\DMT.exe"
$User = "epicordmt"
$Pass = "password"
$datestring = (Get-Date -Format yyyyMMdd-HH)
$Source = "E:\Epicor\EpicorData\DMT\OrderHedStatusUpdate_$datestring.csv"
$completeLog = $source + ".CompleteLog.txt"

#Extract to csv via BAQ WebOrderStatus-DMT
Write-Output "Extracting Data via BAQ $(get-date)"
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Export -BAQ WebOrderStatus-DMT -Target $Source -Company CMP01 -NoUI -ConfigValue=E10Live"

#Load Data
Write-Output "Loading Data $(get-date)" $Source 
Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Update -Import `"Sales Order Header`" -Source $Source "

#Check Results
select-string -Path $completeLog -Pattern "Records:\D*(\d+\/\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
select-string -Path $completeLog -Pattern "Errors:\D*(\d+)" -AllMatches | % { $_.Matches.Groups[0].Value }
Add-Content $completeLog "`r`n#SalesOrderUpdate"

#Email Results to Socialcast
$emailBody = (Get-Content $completeLog -Raw| out-string)
Send-MailMessage -From epicor@yourcompany.com -Subject "Epicor DMT Order Status Update Complete" -To share@socialcast.com -Body $emailBody -Attachments $completeLog -SmtpServer exchange.yourcompany.local