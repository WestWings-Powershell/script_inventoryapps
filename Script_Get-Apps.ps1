
$File = "c:\Temp.json"
$AppFilters = @("Microsoft Office 2010 Standard", `
    "Microsoft Office 2010 стандартный", `
    "Microsoft Office 2010 Publisher ", `
    "Microsoft Office 365", `
    "Kaspersky Endpoint Security for Windows", `
    "PDFsam Enhanced", `
    "Microsoft Visio", `
    "Kaspersky Security Center", `
    "ABBYY FineReader", `
    "ACDSee Pro", `
    "Adobe Acrobat", `
    "Adobe After Effects", `
    "Adobe AIR", `
    "Adobe Photoshop", `
    "Adobe Premiere", `
    "Movavi", `
    "PDF Architect", `
    "PDFCreator", `
    "КриптоПро CSP")

$DataApps = Get-CimInstance -ClassName Win32_Product
$PayApps = ForEach ($Filter in $AppFilters){
    $DataApps | where {$_.name -like "$Filter*"}   
}

$array = [PSCustomObject]@{
    PCName = $Env:COMPUTERNAME
    IP = (Get-NetIPAddress | where {$_.IPAddress -match '[\d]{2,3}.[\d]{1,3}.[\d]{1,3}.[\d]{1,3}'}).IPAddress
    Users = (Get-ChildItem C:\Users\*).Name
    AllApps = ($DataApps).Name
    CountAllApps = ($DataApps).Count
    PayApps = ($PayApps).Name
    CountPayApps = ($PayApps).count
}