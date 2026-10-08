# Minimal XLSX writer: builds an Office Open XML workbook with System.IO.Compression.
# No Excel, no COM, no ImportExcel module. Windows PowerShell 5.1 compatible.
#
#   Export-Xlsx -Path report.xlsx -Sheets @(
#       @{ Name = 'Data'; Columns = @(@{ Key = 'Name'; Width = 20; Type = 'text' }, ...); Rows = $objects; Freeze = $true; Filter = $true }
#   )
#
# Column types: text, int, num, date, datetime, auto. A row property "__Highlight" = $true paints the row light red.

function ConvertTo-XlsxColumnName {
    param([int]$Index)
    $name = ''
    $n = $Index
    while ($n -gt 0) {
        $m = ($n - 1) % 26
        $name = [string][char](65 + $m) + $name
        $n = [int](($n - 1 - $m) / 26)
    }
    return $name
}

function ConvertTo-XlsxText {
    param($Value)
    $text = [string]$Value -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    return [System.Security.SecurityElement]::Escape($text)
}

function Add-XlsxCell {
    param([System.Text.StringBuilder]$Builder, [string]$Ref, $Value, [string]$Type, [bool]$Highlight)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    if ($Type -eq 'auto') {
        if ($Value -is [datetime]) { $Type = 'datetime' }
        elseif ($Value -is [int] -or $Value -is [long]) { $Type = 'int' }
        elseif ($Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) { $Type = 'num' }
        else { $Type = 'text' }
    }
    # Style indexes in styles.xml: 2..6 normal, 7..11 highlighted.
    $style = @{ text = 2; date = 3; datetime = 4; num = 5; int = 6 }[$Type]
    if ($null -eq $style) { $style = 2 }
    if ($Highlight) { $style += 5 }

    if ($null -eq $Value -or ([string]$Value) -eq '') {
        [void]$Builder.Append("<c r=""$Ref"" s=""$style""/>")
        return
    }
    switch ($Type) {
        { $_ -in 'date', 'datetime' } {
            $v = ([datetime]$Value).ToOADate().ToString('R', $inv)
            [void]$Builder.Append("<c r=""$Ref"" s=""$style""><v>$v</v></c>")
        }
        { $_ -in 'num', 'int' } {
            $v = ([double]$Value).ToString('R', $inv)
            [void]$Builder.Append("<c r=""$Ref"" s=""$style""><v>$v</v></c>")
        }
        default {
            [void]$Builder.Append("<c r=""$Ref"" s=""$style"" t=""inlineStr""><is><t xml:space=""preserve"">$(ConvertTo-XlsxText $Value)</t></is></c>")
        }
    }
}

function New-XlsxSheetXml {
    param([hashtable]$Sheet, [bool]$Selected)
    $columns = @($Sheet.Columns)
    $rows = @($Sheet.Rows)
    $lastCol = ConvertTo-XlsxColumnName $columns.Count
    $lastRow = $rows.Count + 1
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append("<dimension ref=""A1:$lastCol$lastRow""/>")
    $tabSelected = if ($Selected) { ' tabSelected="1"' } else { '' }
    if ($Sheet.Freeze) {
        [void]$sb.Append("<sheetViews><sheetView workbookViewId=""0""$tabSelected><pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/><selection pane=""bottomLeft"" activeCell=""A2"" sqref=""A2""/></sheetView></sheetViews>")
    }
    else {
        [void]$sb.Append("<sheetViews><sheetView workbookViewId=""0""$tabSelected/></sheetViews>")
    }

    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
    for ($i = 0; $i -lt $columns.Count; $i++) {
        $width = ([double]$columns[$i].Width).ToString($inv)
        [void]$sb.Append("<col min=""$($i + 1)"" max=""$($i + 1)"" width=""$width"" customWidth=""1""/>")
    }
    [void]$sb.Append('</cols><sheetData><row r="1" ht="30" customHeight="1">')
    for ($i = 0; $i -lt $columns.Count; $i++) {
        [void]$sb.Append("<c r=""$(ConvertTo-XlsxColumnName ($i + 1))1"" s=""1"" t=""inlineStr""><is><t xml:space=""preserve"">$(ConvertTo-XlsxText $columns[$i].Key)</t></is></c>")
    }
    [void]$sb.Append('</row>')

    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $highlight = [bool]($row.PSObject.Properties['__Highlight'] -and $row.__Highlight)
        [void]$sb.Append("<row r=""$rowNumber"">")
        for ($i = 0; $i -lt $columns.Count; $i++) {
            Add-XlsxCell -Builder $sb -Ref ((ConvertTo-XlsxColumnName ($i + 1)) + $rowNumber) `
                -Value $row.($columns[$i].Key) -Type $columns[$i].Type -Highlight $highlight
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($Sheet.Filter) { [void]$sb.Append("<autoFilter ref=""A1:$lastCol$lastRow""/>") }
    [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    return $sb.ToString()
}

function Export-Xlsx {
    <#
    .SYNOPSIS
        Writes one or more sheets of objects to an .xlsx file without Excel.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][hashtable[]]$Sheets
    )
    Add-Type -AssemblyName System.IO.Compression
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $nsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $nsPkg = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

    $contentTypes = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $contentTypes += "<Override PartName=""/xl/worksheets/sheet$i.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>"
    }
    $contentTypes += '</Types>'

    $rootRels = $head + "<Relationships xmlns=""$nsPkg""><Relationship Id=""rId1"" Type=""$nsRel/officeDocument"" Target=""xl/workbook.xml""/></Relationships>"

    $workbook = $head + "<workbook xmlns=""$ns"" xmlns:r=""$nsRel""><bookViews><workbookView activeTab=""0""/></bookViews><sheets>"
    $workbookRels = $head + "<Relationships xmlns=""$nsPkg"">"
    $definedNames = ''
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $sheet = $Sheets[$i - 1]
        $workbook += "<sheet name=""$(ConvertTo-XlsxText $sheet.Name)"" sheetId=""$i"" r:id=""rId$i""/>"
        $workbookRels += "<Relationship Id=""rId$i"" Type=""$nsRel/worksheet"" Target=""worksheets/sheet$i.xml""/>"
        if ($sheet.Filter) {
            $range = "'{0}'!`$A`$1:`${1}`${2}" -f $sheet.Name, (ConvertTo-XlsxColumnName @($sheet.Columns).Count), (@($sheet.Rows).Count + 1)
            $definedNames += "<definedName name=""_xlnm._FilterDatabase"" localSheetId=""$($i - 1)"" hidden=""1"">$(ConvertTo-XlsxText $range)</definedName>"
        }
    }
    $workbook += '</sheets>'
    if ($definedNames) { $workbook += "<definedNames>$definedNames</definedNames>" }
    $workbook += '</workbook>'
    $workbookRels += "<Relationship Id=""rId$($Sheets.Count + 1)"" Type=""$nsRel/styles"" Target=""styles.xml""/></Relationships>"

    # Cell formats: 0 default, 1 header, 2-6 text/date/datetime/num/int, 7-11 the same with highlight fill.
    $border = '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'
    $xf = ''
    foreach ($fill in 0, 3) {
        foreach ($fmt in 0, 164, 165, 2, 1) {
            $xf += "<xf numFmtId=""$fmt"" fontId=""0"" fillId=""$fill"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1""><alignment vertical=""top""/></xf>"
        }
    }
    $styles = $head + "<styleSheet xmlns=""$ns"">" +
        '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm"/></numFmts>' +
        '<fonts count="2"><font><sz val="10"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>' +
        '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFBE2E2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' + $border + '</borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="12"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
        $xf + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

    $parts = [ordered]@{
        '[Content_Types].xml'        = $contentTypes
        '_rels/.rels'                = $rootRels
        'xl/workbook.xml'            = $workbook
        'xl/_rels/workbook.xml.rels' = $workbookRels
        'xl/styles.xml'              = $styles
    }
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        $parts["xl/worksheets/sheet$i.xml"] = New-XlsxSheetXml -Sheet $Sheets[$i - 1] -Selected ($i -eq 1)
    }

    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::CreateNew)
    $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) {
            $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
            try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
        }
    }
    finally {
        $zip.Dispose()
        $stream.Dispose()
    }
}
