param (
    [string]$name = '*',
    [string]$path = 'LDAP://OU=mailboxes;DC=dvfu;DC=ru',
    [bool]$printPropNames = $false
)

# возьмём константы, чтобы удобнее отлаживаться
$name = "klenin*"
$path = 'LDAP://DC=mydomain;DC=loc'#'LDAP://DC=dvfu;DC=ru'

$root = New-Object DirectoryServices.DirectoryEntry($path)
$selector = New-Object DirectoryServices.DirectorySearcher
$selector.SearchRoot = $root

# если не указать "PageSize", то вычитает первую 1000 объектов
# если указать больше 1000, то всё равно вычитает первую 1000
# это ограничение указывается в настройках Active Directory
# а если указать 100, 200 или даже 500, то оно само собой будет в цикле получать хоть 20000 объектов
$selector.PageSize = 100

# вот такой код найдёт все объекты, вычитает их, а потом будет фильтровать
#$adObj = $selector.findall() | `
#    where {($_.properties.objectcategory -Match "CN=Person") -and ($_.properties.samaccountname -Like $name)}
# вот такой код сразу найдёт только нужные объекты. ничего лишнего вычитываться не будет
# "*" после $name нужна для того, чтобы просто указывать маску для имён.
# $name может быть "klenin.a", "klenin.a*", "kle*", ....
$selector.Filter = "(&(objectCategory=Person)(samaccountname=$name*))"
$adObj = $selector.findall()
#"$($adObj.count) user(s) found"

if ($adObj.count) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Cells.Item(1, 1) = "$($adObj.count) user(s) found"
    $counter = 0

    if ($printPropNames) {
        $counter++
    }

    foreach ($person in $adObj) {
        $counter++
        $sheet.Cells.Item(1, $counter) = $person.Path
        # безумный код для того, чтобы получить два списка свойств
        $p2 = $person.GetDirectoryEntry() # надо будет освободить объект в конце цикла
        $schema = $p2.SchemaEntry # надо будет освободить объект в конце цикла
        $t = $schema.NativeObject.GetType()

        # получаем имена обязательных свойств
        $mandatoryPropNames = $t.InvokeMember("MandatoryProperties", [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::GetProperty, $null, $schema.NativeObject, $null)

        # получаем имена опциональных свойств
        $optionalPropNames = $t.InvokeMember("OptionalProperties", [System.Reflection.BindingFlags]::Public -bor [System.Reflection.BindingFlags]::GetProperty, $null, $schema.NativeObject, $null)
        $propNames = $mandatoryPropNames + $optionalPropNames | sort
        $j = 1
        foreach ($name in $propNames) {
            $j++
            if ($mandatoryPropNames.Contains($name)) {
                $textMarker = "m"
            } else {
                $textMarker = "o"
            }

            $propertyName = $name

            # штатный способ получения свойств в ADSI
            # они всегда сидят как массив, даже простые числа и строки
            # но в этом случае надо особо обрабатывать массивы
            if ($person.Properties[$name].Count -gt 1) {
                # массив значений
                # например, это "memberOf"
                $propertyValue = "{0}" -f [string] $person.Properties[$name]
            }
            else {
                # одинарное значение
                # [string] заодно превращает значения "массив байтов" в последовательность чисел
                # например, это objectGUID или objectSid
                # если добавить условие и чуток кода, то можно красивенько форматировать "массив байтов" в виде "0x01 0xAB" или "01 AB"
                $propertyValue = "{0}" -f [string] $person.Properties[$name][0]
            }
            if ($propertyValue) {
                $sheet.Cells.Item($j, $counter) = $propertyValue
            }
            if ($printPropNames) {
                $sheet.Cells.Item($j, 1) = $propertyName
            }
        }

        # надо освободить объекты.
        # они сами потом освободятся, но могут быть нюансы при больших объёмах или работе в интерактивном режиме
        $schema.Dispose()
        $p2.Dispose()
        $printPropNames = $false
    }
    $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
    $excel.ActiveWorkbook.SaveAs("C:\Users\Misha\Desktop\ActiveDirectory\myfile.xls", $xlFixedFormat)
    $excel.Workbooks.Close()
    $excel.Quit()    
}
