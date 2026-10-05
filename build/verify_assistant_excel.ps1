$ErrorActionPreference = "Stop"
$path = "D:\Portfolio\app-excel\dist\Sistema_Financiero_Empresa_2026.xlsx"
$excel = $null
$book = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $book = $excel.Workbooks.Open($path, 0, $false)

    $wsInicio = $book.Worksheets.Item("INICIO")
    $first = $book.Worksheets.Item(1)
    if ($wsInicio.Index -ne 1) { $wsInicio.Move($first) }
    $wsInicio.Activate()

    $homeLinks = @(
        @("A6", "EMPRESA", "1. CONFIGURAR EMPRESA"),
        @("F6", "MOVIMIENTOS", "2. REGISTRAR MOVIMIENTO"),
        @("A10", "ASISTENTE", "3. ASISTENTE LOCAL"),
        @("F10", "DOCUMENTOS", "4. REVISAR DOCUMENTOS"),
        @("A14", "DASHBOARD", "5. VER INDICADORES"),
        @("F14", "REPORTE_MENSUAL", "6. ABRIR REPORTES"),
        @("A26", "MENU", "Ver todos los módulos"),
        @("F26", "REGLAS_IA", "Editar reglas del asistente")
    )
    foreach ($item in $homeLinks) {
        $cell = $wsInicio.Range($item[0])
        if ($cell.Hyperlinks.Count -eq 0) {
            $cell.Value2 = $item[2]
            $wsInicio.Hyperlinks.Add($cell, "", "'" + $item[1] + "'!A1", "", $item[2]) | Out-Null
        }
        $actualTarget = $cell.Hyperlinks.Item(1).SubAddress.Replace("'", "")
        if ($cell.Hyperlinks.Count -ne 1 -or $actualTarget -ne ($item[1] + "!A1")) {
            throw "Wrong INICIO link at $($item[0]): $actualTarget"
        }
    }

    $assistant = $book.Worksheets.Item("ASISTENTE")
    foreach ($item in @(
        @("A26", "REGLAS_IA", "Editar las reglas y crear patrones propios"),
        @("A27", "INICIO", "Volver al inicio")
    )) {
        $cell = $assistant.Range($item[0])
        if ($cell.Hyperlinks.Count -eq 0) {
            $cell.Value2 = $item[2]
            $assistant.Hyperlinks.Add($cell, "", "'" + $item[1] + "'!A1", "", $item[2]) | Out-Null
        }
        $actualTarget = $cell.Hyperlinks.Item(1).SubAddress.Replace("'", "")
        if ($cell.Hyperlinks.Count -ne 1 -or $actualTarget -ne ($item[1] + "!A1")) {
            throw "Wrong ASISTENTE link at $($item[0]): $actualTarget"
        }
    }

    $menu = $book.Worksheets.Item("MENU")
    if ($menu.Range("A6").Value2 -ne "INICIO") {
        $menu.Rows("6:6").Insert() | Out-Null
        $menu.Rows("14:15").Insert() | Out-Null
        foreach ($item in @(
            @("A6", "INICIO", "INICIO"),
            @("A14", "ASISTENTE", "ASISTENTE DE CLASIFICACION"),
            @("A15", "REGLAS_IA", "REGLAS DEL ASISTENTE")
        )) {
            $cell = $menu.Range($item[0])
            $cell.Value2 = $item[2]
            $menu.Hyperlinks.Add($cell, "", "'" + $item[1] + "'!A1", "", $item[2]) | Out-Null
            $cell.Font.Color = 0xCC5511
            $cell.Font.Underline = 2
            $cell.Font.Bold = $true
            $cell.Interior.Color = 0xFBF1EA
        }
    }
    foreach ($item in @(
        @("A6", "INICIO"), @("A14", "ASISTENTE"), @("A15", "REGLAS_IA")
    )) {
        $cell = $menu.Range($item[0])
        $actualTarget = $cell.Hyperlinks.Item(1).SubAddress.Replace("'", "")
        if ($actualTarget -ne ($item[1] + "!A1")) {
            throw "Wrong MENU link at $($item[0]): $actualTarget"
        }
    }

    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()

    $rules = $book.Worksheets.Item("REGLAS_IA")
    if ($rules.ListObjects.Item("REGLAS_IA").Range.Address() -ne '$A$4:$L$104') {
        throw "REGLAS_IA table does not cover rows 4:104."
    }
    foreach ($address in @("B5", "D5", "E5", "G5", "H5")) {
        if ($rules.Range($address).Validation.Type -ne 3) { throw "List validation missing at $address." }
    }
    if ($rules.Range("J5").Validation.Type -ne 1) {
        throw "Priority whole-number validation is missing."
    }

    $assistant.Range("B5:B6").ClearContents()
    $excel.CalculateFullRebuild()
    if ($assistant.Range("B9").Text -ne "Escriba un concepto") {
        throw "Empty assistant input state failed."
    }

    $assistant.Range("B5").Value2 = "Proveedor de prueba"
    $assistant.Range("B6").Value2 = "Pago de arriendo local"
    $excel.CalculateFullRebuild()
    $rent = @($assistant.Range("B9").Text, $assistant.Range("B10").Text, $assistant.Range("B11").Text, $assistant.Range("B12").Text, $assistant.Range("B15").Text, $assistant.Range("B16").Text)
    if (($rent -join "|") -ne "Sugerencia lista para revisar|Gasto|Administracion|Arriendo|arriendo|20") {
        throw "Arriendo rule failed: $($rent -join ' | ')"
    }

    $assistant.Range("B6").Value2 = "Venta de productos y venta de servicios"
    $excel.CalculateFullRebuild()
    if ($assistant.Range("B9").Text -ne "Conflicto: ajuste prioridad" -or $assistant.Range("B10").Text -ne "") {
        throw "Equal-priority conflict was not detected."
    }

    $assistant.Range("B6").Value2 = "concepto sin regla conocida"
    $excel.CalculateFullRebuild()
    if ($assistant.Range("B9").Text -ne "Sin coincidencia") {
        throw "No-match state failed."
    }

    $rules.Range("A21").Value2 = "IA-017"
    $rules.Range("B21").Value2 = "Sí"
    $rules.Range("C21").Value2 = "parqueadero"
    $rules.Range("D21").Value2 = "Gasto"
    $rules.Range("E21").Value2 = "Administracion"
    $rules.Range("F21").Value2 = "Arriendo"
    $rules.Range("J21").Value2 = [double]35
    $rules.Range("L21").Value2 = "Regla personalizada de prueba"
    $assistant.Range("B6").Value2 = "Pago de parqueadero"
    $excel.CalculateFullRebuild()
    if ($assistant.Range("B9").Text -ne "Sugerencia lista para revisar" -or
        $assistant.Range("B10").Text -ne "Gasto" -or
        $assistant.Range("B11").Text -ne "Administracion" -or
        $assistant.Range("B12").Text -ne "Arriendo" -or
        $assistant.Range("B15").Text -ne "parqueadero" -or
        $assistant.Range("B16").Value2 -ne 35) {
        throw "Custom rule test failed."
    }

    $assistant.Range("B5:B6").ClearContents()
    $rules.Range("A21:J21").ClearContents()
    $rules.Range("L21").ClearContents()
    $excel.CalculateFullRebuild()

    $formulaErrors = 0
    foreach ($sheet in $book.Worksheets) {
        try {
            $errors = $sheet.UsedRange.SpecialCells(-4123, 16)
            $formulaErrors += $errors.CountLarge
        } catch {
            # Excel raises when the range contains no formula errors.
        }
    }
    if ($formulaErrors -ne 0) {
        throw "Workbook contains $formulaErrors formula errors after recalculation."
    }

    if ($book.Worksheets.Count -ne 28) { throw "Expected 28 sheets; found $($book.Worksheets.Count)." }
    $book.Save()
    Write-Output "Excel COM QA passed: home navigation, rule dropdowns, rent match, conflict, no-match, custom rule, and zero formula errors."
    Write-Output "Clean final state saved with blank assistant inputs and no test rule. Active sheet: $($excel.ActiveSheet.Name)."
} finally {
    if ($book -ne $null) {
        try {
            $book.Worksheets.Item("ASISTENTE").Range("B5:B6").ClearContents()
            $book.Worksheets.Item("REGLAS_IA").Range("A21:J21").ClearContents()
            $book.Worksheets.Item("REGLAS_IA").Range("L21").ClearContents()
            $book.Save()
        } catch { }
        $book.Close($true) | Out-Null
    }
    if ($excel -ne $null) { $excel.Quit() }
    if ($book -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
