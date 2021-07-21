
# Заполняем переменные для работы 
$Path = "c:"
$FileName = "Temp.json"
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

# Проводим проверку установленного ПО     
$DataApps = Get-CimInstance -ClassName Win32_Product
$PayApps = ForEach ($Filter in $AppFilters){
    $DataApps.where{$_.name -like "$Filter*"}   
}

# Формируем хэштаблицу с полученными данными 
$array = [PSCustomObject]@{
    PCName = $Env:COMPUTERNAME
    IP = (Get-NetIPAddress.IPAddress).where{$_ -like '192*'} 
    Users = (Get-ChildItem C:\Users\*).Name
    AllApps = ($DataApps).Name
    CountAllApps = ($DataApps).Count
    PayApps = ($PayApps).Name
    CountPayApps = ($PayApps).count
}

# Делаем проверку наличия файа и записываем данные в файл в JSON формате
$JSONFind = Get-ChildItem -Path $Path\$FileName
if ($null -eq $JSONFind){
    $array | ConvertTo-Json | Out-File -FilePath $Path\$FileName 
}

<# 
    если файл есть - тогда забираем данные из него и 
    преобразуем его обратно в хеш таблицу и 
    перезаписываем обратно уникальные занчения 
#>
else {
    $OldData = Get-Content $Path\$FileName
    $OldObjectData = $OldData | Select-Object
    $OldJSONData = $OldObjectData | ConvertFrom-Json
    $UniqueData = $OldJSONData + $array | Get-Unique
    $UniqueData | Out-File $Path\$FileName
}

