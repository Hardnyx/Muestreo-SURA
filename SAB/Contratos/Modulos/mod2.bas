Option Explicit

Public Sub ExportarMuestra()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSrc As Worksheet
    Dim tbl As ListObject
    Dim cntPN As Long, cntPJ As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' Hoja y tabla origen
    Set wsSrc = wb.Worksheets("Contratos")
    Set tbl = wsSrc.ListObjects("Contratos")
    
    If tbl Is Nothing Or tbl.DataBodyRange Is Nothing Then
        MsgBox "No se encontró la tabla 'Contratos' o está vacía.", vbCritical
        GoTo Salir
    End If
    
    ' Leer tamaño de muestra
    cntPN = wb.Names("TamañoMuestraPN").RefersToRange.Value
    cntPJ = wb.Names("TamañoMuestraPJ").RefersToRange.Value
    
    ' Exportar muestras
    ExportFiltered tbl, "N", "Muestra_Contratos_PN", cntPN
    ExportFiltered tbl, "J", "Muestra_Contratos_PJ", cntPJ
    
    MsgBox "Exportación completada." & vbCrLf & _
           "PN: " & cntPN & " fila(s)." & vbCrLf & _
           "PJ: " & cntPJ & " fila(s).", vbInformation

Salir:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Salir
End Sub

'---------------------------------------------------
Private Sub ExportFiltered(tbl As ListObject, tipo As String, hojaDestino As String, tamaño As Long)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDest As Worksheet
    Dim tipoCol As Long, fechaCol As Long
    Dim db As Range
    Dim rowList() As Range
    Dim k As Long, destRow As Long, i As Long
    Dim wsMuestra As Worksheet
    Dim mesFiltro As String, añoFiltro As Long, tipoInforme As String
    Dim fechaVal As Date, fechaTexto As String
    Dim mesNumFiltro As Long
    
    ' Leer parámetros de filtros
    Set wsMuestra = wb.Worksheets("Muestra")
    mesFiltro = Trim(wsMuestra.Range("Mes").Value)
    añoFiltro = CLng(wsMuestra.Range("Año").Value)
    tipoInforme = Trim(wsMuestra.Range("TipoInforme").Value)
    mesNumFiltro = MonthNumberFromName(mesFiltro)
    
    ' Crear/limpiar hoja destino
    On Error Resume Next
    Set wsDest = wb.Worksheets(hojaDestino)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDest.name = hojaDestino
    Else
        wsDest.Cells.Clear
    End If
    wsDest.Visible = xlSheetVeryHidden
    
    ' Copiar encabezado
    tbl.HeaderRowRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    destRow = 2
    
    ' Columnas
    tipoCol = GetListColumnIndex(tbl, "Tipo")
    fechaCol = GetListColumnIndex(tbl, "Fecha de Ingreso")
    Set db = tbl.DataBodyRange
    
    ' Construir lista de filas que cumplen filtros
    k = 0
    For i = 1 To db.Rows.Count
        If CStr(db.Cells(i, tipoCol).Value) = tipo Then
            fechaTexto = Trim(CStr(db.Cells(i, fechaCol).Value))
            If Len(fechaTexto) >= 7 Then
                On Error Resume Next
                fechaVal = DateValue(Left(fechaTexto, 2) & "-" & Mid(fechaTexto, 3, 3) & "-20" & Right(fechaTexto, 2))
                On Error GoTo 0
                If Year(fechaVal) = añoFiltro Then
                    If UCase(tipoInforme) = "ANUAL" Or _
                       (UCase(tipoInforme) = "MENSUAL" And Month(fechaVal) = mesNumFiltro) Then
                        k = k + 1
                        ReDim Preserve rowList(1 To k)
                        Set rowList(k) = db.Rows(i)
                        If k >= tamaño Then Exit For ' limitar al tamaño exacto
                    End If
                End If
            End If
        End If
    Next i
    
    ' Copiar filas
    For i = 1 To k
        rowList(i).Copy
        wsDest.Cells(destRow, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        destRow = destRow + 1
    Next i
    
    ' Crear tabla
    If k > 0 Then
        On Error Resume Next
        wsDest.ListObjects(hojaDestino).Delete
        On Error GoTo 0
        Dim lo As ListObject
        Set lo = wsDest.ListObjects.Add(xlSrcRange, wsDest.Range("A1").CurrentRegion, , xlYes)
        lo.name = hojaDestino
    End If
    
    wsDest.Cells.EntireColumn.AutoFit
    wsDest.Rows.RowHeight = wsDest.StandardHeight
    wsDest.Visible = xlSheetVisible
End Sub

'---------------------------------------------------
Private Function GetListColumnIndex(lo As ListObject, colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).name, colName, vbTextCompare) = 0 Then
            GetListColumnIndex = i
            Exit Function
        End If
    Next i
    GetListColumnIndex = 0
End Function

'---------------------------------------------------
Private Function MonthNumberFromName(mes As String) As Long
    Select Case UCase(Trim(mes))
        Case "ENERO": MonthNumberFromName = 1
        Case "FEBRERO": MonthNumberFromName = 2
        Case "MARZO": MonthNumberFromName = 3
        Case "ABRIL": MonthNumberFromName = 4
        Case "MAYO": MonthNumberFromName = 5
        Case "JUNIO": MonthNumberFromName = 6
        Case "JULIO": MonthNumberFromName = 7
        Case "AGOSTO": MonthNumberFromName = 8
        Case "SEPTIEMBRE": MonthNumberFromName = 9
        Case "OCTUBRE": MonthNumberFromName = 10
        Case "NOVIEMBRE": MonthNumberFromName = 11
        Case "DICIEMBRE": MonthNumberFromName = 12
        Case Else: MonthNumberFromName = 0
    End Select
End Function

