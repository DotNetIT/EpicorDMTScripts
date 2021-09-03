#DMT Automation Example 1

$DMTPath = "C:\Epicor\ERP10\LocalClients\ERP10\DMT.exe"
$User = "epicor"
$Pass = "epicor"
$Source = "C:\Temp\Product.csv"

#Load Data

Start-Process -Wait -FilePath $DMTPath -ArgumentList "-User $User -Pass $Pass -Delete -Import Part -Source $Source "