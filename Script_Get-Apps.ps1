# Заполняем переменные для работы 
$Path = "\\SMR01\Share_Z\IT"
$FileName = "APPsInventory.json"

# Фильтр условно платных приложений
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
$DataApps = Get-WMIObject -ClassName Win32_Product
$PayApps = ForEach ($Filter in $AppFilters){
    $DataApps.where{$_.name -like "$Filter*"}   
}

# Формируем хэштаблицу с полученными данными 
$NewData = [PSCustomObject]@{
    PCName = $Env:COMPUTERNAME
    IP = ((Get-NetIPAddress).where{$_.IPAddress -like '192*'}).IPAddress
    Users = (Get-ChildItem C:\Users\*).Name
    AllApps = ($DataApps).Name
    CountAllApps = ($DataApps).name.Count
    PayApps = ($PayApps).Name
    CountPayApps = ($PayApps).name.Count
}

# Делаем проверку наличия файа и записываем данные в файл в JSON формате
$JSONFind = Get-ChildItem -Path $Path\$FileName
if ($null -eq $JSONFind){
    $NewData | ConvertTo-Json | Out-File -FilePath $Path\$FileName 
}

<# 
    если файл есть - тогда забираем данные из него и 
    преобразуем его обратно в хеш таблицу и 
    перезаписываем обратно уникальные занчения 
#>
else {
    $OldData = Get-Content $Path\$FileName
    $OldObjectData = $OldData | Select-Object
    $uniqueComps = ($OldObjectData + $NewData).PCName | Get-Unique 
    $uniqueData = foreach ($uniqueComp in $uniqueComps) {
        $Data = $NewData | Where-Object {$_.PCName -like "$uniqueComp"} 
        $Data
    }
    
    # Конвертируем обратно и записываем в файл 
    ConvertTo-Json -InputObject $uniqueData | Out-File $Path\$FileName  -Encoding utf8
}