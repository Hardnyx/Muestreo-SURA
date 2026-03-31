Option Explicit

Public Sub TamañoPoblacion()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsContratos As Worksheet
    Dim tbl As ListObject
    Dim fechaCol As Long, cuentaCol As Long, tipoCol As Long
    Dim tipoInforme As String, mesTexto As String
    Dim anioFiltro As Long, mesFiltro As Long
    Dim db As Range
    Dim i As Long, contadorTotal As Long, contadorN As Long, contadorJ As Long
    Dim s As String, monthTxt As String, yearTxt As String
    Dim monthNum As Long, yearNum As Long, yearFull As Long
    Dim cuentaValor As String, tipoValor As String

    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Hojas y tabla
    Set wsContratos = wb.Worksheets("Contratos")
    Set tbl = wsContratos.ListObjects("Contratos")

    ' Localizar columnas
    fechaCol = GetListColumnIndex(tbl, "Fecha de Ingreso")
    If fechaCol = 0 Then fechaCol = GetListColumnIndex(tbl, "FechaIngreso")
    cuentaCol = GetListColumnIndex(tbl, "Cuenta")
    tipoCol = GetListColumnIndex(tbl, "Tipo") ' <-- NUEVO: columna Tipo

    ' Leer filtros
    tipoInforme = UCase$(Trim$(CStr(wb.Names("TipoInforme").RefersToRange.Value)))
    anioFiltro = CLng(wb.Names("Año").RefersToRange.Value)
    If tipoInforme = "MENSUAL" Then
        mesTexto = Trim$(CStr(wb.Names("Mes").RefersToRange.Value))
        mesFiltro = MonthNumberFromNameSpanish(mesTexto)
    Else
        mesFiltro = 0
    End If

    ' Validaciones mínimas
    If fechaCol = 0 Then
        MsgBox "No se encontró la columna 'Fecha de Ingreso' en la tabla 'Contratos'.", vbCritical
        GoTo Cleanup
    End If
    If cuentaCol = 0 Then
        MsgBox "No se encontró la columna 'Cuenta' en la tabla 'Contratos'.", vbCritical
        GoTo Cleanup
    End If
    ' tipoCol puede ser 0 (fallback más abajo)

    ' Recorrer filas
    Set db = tbl.DataBodyRange
    contadorTotal = 0: contadorN = 0: contadorJ = 0

    For i = 1 To db.Rows.Count
        s = Trim$(CStr(db.Cells(i, fechaCol).Value))
        If Len(s) >= 5 Then
            monthTxt = UCase$(Mid$(s, 3, 3))
            monthNum = MonthNumberFromAbbrevSpanish(monthTxt)
            yearTxt = Mid$(s, 6)
            If IsNumeric(yearTxt) Then
                yearNum = CLng(yearTxt)
                If Len(Trim$(yearTxt)) >= 4 Then
                    yearFull = yearNum
                Else
                    yearFull = 2000 + yearNum
                End If

                If yearFull = anioFiltro And (mesFiltro = 0 Or monthNum = mesFiltro) Then
                    cuentaValor = Trim$(CStr(db.Cells(i, cuentaCol).Value))

                    ' Obtener tipo real (preferir la columna "Tipo"; si no existe, fallback a primera letra de Cuenta)
                    If tipoCol <> 0 Then
                        tipoValor = Trim$(CStr(db.Cells(i, tipoCol).Value))
                    Else
                        If Len(cuentaValor) > 0 Then
                            tipoValor = Left$(cuentaValor, 1)
                        Else
                            tipoValor = ""
                        End If
                    End If

                    If Len(cuentaValor) > 0 Then
                        contadorTotal = contadorTotal + 1
                        Select Case UCase$(Left$(tipoValor & " ", 1))
                            Case "N": contadorN = contadorN + 1
                            Case "J": contadorJ = contadorJ + 1
                        End Select
                    End If
                End If
            End If
        End If
    Next i

    ' Guardar resultados
    wb.Names("TamañoPob").RefersToRange.Value = contadorTotal
    If Not NameNotFound(wb, "UniversoPN") Then wb.Names("UniversoPN").RefersToRange.Value = contadorN
    If Not NameNotFound(wb, "UniversoPJ") Then wb.Names("UniversoPJ").RefersToRange.Value = contadorJ

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ---------- Funciones auxiliares ----------

' Busca una columna en ListObject por nombre (intenta exacto, luego por palabras contenidas)
Private Function GetListColumnIndex(lo As ListObject, colName As String) As Long
    Dim i As Long, tokens As Variant, t As Variant, nameLower As String, matchAll As Boolean

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), colName, vbTextCompare) = 0 Then
            GetListColumnIndex = i: Exit Function
        End If
    Next i

    ' Si no hubo coincidencia exacta, intentar por tokens (p. ej. "Fecha" + "Ingreso")
    tokens = Split(LCase$(colName))
    For i = 1 To lo.ListColumns.Count
        nameLower = LCase$(lo.ListColumns(i).name)
        matchAll = True
        For Each t In tokens
            If InStr(nameLower, t) = 0 Then
                matchAll = False: Exit For
            End If
        Next t
        If matchAll Then
            GetListColumnIndex = i: Exit Function
        End If
    Next i
    GetListColumnIndex = 0
End Function

' Convierte abreviatura de mes español (3 letras) a número
Private Function MonthNumberFromAbbrevSpanish(abbrev As String) As Long
    Select Case UCase$(Left$(abbrev & "   ", 3))
        Case "ENE": MonthNumberFromAbbrevSpanish = 1
        Case "FEB": MonthNumberFromAbbrevSpanish = 2
        Case "MAR": MonthNumberFromAbbrevSpanish = 3
        Case "ABR": MonthNumberFromAbbrevSpanish = 4
        Case "MAY": MonthNumberFromAbbrevSpanish = 5
        Case "JUN": MonthNumberFromAbbrevSpanish = 6
        Case "JUL": MonthNumberFromAbbrevSpanish = 7
        Case "AGO": MonthNumberFromAbbrevSpanish = 8
        Case "SEP", "SET": MonthNumberFromAbbrevSpanish = 9 ' acepta SEP o SET
        Case "OCT": MonthNumberFromAbbrevSpanish = 10
        Case "NOV": MonthNumberFromAbbrevSpanish = 11
        Case "DIC": MonthNumberFromAbbrevSpanish = 12
        Case Else: MonthNumberFromAbbrevSpanish = 0
    End Select
End Function

' Convierte nombre de mes en español ("Enero", "ene",...) a número
Private Function MonthNumberFromNameSpanish(name As String) As Long
    If Len(Trim$(name)) = 0 Then MonthNumberFromNameSpanish = 0: Exit Function
    Select Case UCase$(Left$(Trim$(name) & "   ", 3))
        Case "ENE": MonthNumberFromNameSpanish = 1
        Case "FEB": MonthNumberFromNameSpanish = 2
        Case "MAR": MonthNumberFromNameSpanish = 3
        Case "ABR": MonthNumberFromNameSpanish = 4
        Case "MAY": MonthNumberFromNameSpanish = 5
        Case "JUN": MonthNumberFromNameSpanish = 6
        Case "JUL": MonthNumberFromNameSpanish = 7
        Case "AGO": MonthNumberFromNameSpanish = 8
        Case "SEP", "SET": MonthNumberFromNameSpanish = 9
        Case "OCT": MonthNumberFromNameSpanish = 10
        Case "NOV": MonthNumberFromNameSpanish = 11
        Case "DIC": MonthNumberFromNameSpanish = 12
        Case Else: MonthNumberFromNameSpanish = 0
    End Select
End Function

' Comprueba si el nombre existe en el libro
Private Function NameNotFound(wb As Workbook, nm As String) As Boolean
    Dim tmp As name
    On Error Resume Next
    Set tmp = wb.Names(nm)
    NameNotFound = (tmp Is Nothing)
    On Error GoTo 0
End Function

