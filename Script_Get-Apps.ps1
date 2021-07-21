# Функция обновления ключей хэштаблицы - взято отсюда 
# https://coderoad.ru/8800375/%D0%A1%D0%BB%D0%B8%D1%8F%D0%BD%D0%B8%D0%B5-%D1%85%D1%8D%D1%88%D1%82%D0%B0%D0%B1%D0%BE%D0%B2-%D0%B2-PowerShell-%D0%BA%D0%B0%D0%BA
Function Merge-Hashtables {
    $Output = @{}
    ForEach ($Hashtable in ($Input + $Args)) {
        If ($Hashtable -is [Hashtable]) {
            ForEach ($Key in $Hashtable.Keys) {$Output.$Key = $Hashtable.$Key}
        }
    }
    $Output
}

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
$NewData = [PSCustomObject]@{
    PCName = $Env:COMPUTERNAME
    IP = ((Get-NetIPAddress).where{$_.IPAddress -like '192*'}).IPAddress
    Users = (Get-ChildItem C:\Users\*).Name
    AllApps = ($DataApps).Name
    CountAllApps = ($DataApps).Count
    PayApps = ($PayApps).Name
    CountPayApps = ($PayApps).count
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
    $OldJSONData = $OldObjectData | ConvertFrom-Json
    $uniqueComps = ($OldJSONData + $NewData).computername | Get-Unique 
    $uniqueData = foreach ($uniqueComp in $uniqueComps) {
        $Data = $NewData | Where-Object {$_.computername -like "$uniqueComp"} 
        $Data
    }
    
    ConvertTo-Json -InputObject $uniqueData | Out-File $JSONPath -Encoding utf8
}

