Option Explicit

Sub SeleccionMuestra()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Está seguro de que desea generar nuevas muestras PN y PJ?", vbYesNo + vbQuestion, "Confirmar")
    If resp <> vbYes Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' PN
    GenerarMuestraOrdenada "Muestra1_PN", "TamañoMuestraPN", "UniversoPN"
    ' PJ
    GenerarMuestraOrdenada "Muestra1_PJ", "TamañoMuestraPJ", "UniversoPJ"
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Muestras PN y PJ generadas correctamente.", vbInformation
End Sub

Private Sub GenerarMuestraOrdenada(nombreInicio As String, nombreTamano As String, nombreUniverso As String)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim rngInicio As Range
    Dim ws As Worksheet
    Dim tamano As Long, universo As Long
    Dim coll As Collection
    Dim numeros() As Long
    Dim i As Long, j As Long, tmp As Long
    Dim startRow As Long, startCol As Long
    Dim lastRowUsed As Long, lastColUsed As Long
    Dim c As Long, r As Long
    Dim rnum As Long
    Dim fila As Long, col As Long
    
    ' Obtener la celda de inicio (plantilla)
    On Error Resume Next
    Set rngInicio = wb.Names(nombreInicio).RefersToRange
    On Error GoTo 0
    If rngInicio Is Nothing Then
        MsgBox "No existe el nombre definido '" & nombreInicio & "'.", vbCritical
        Exit Sub
    End If
    Set ws = rngInicio.Parent   ' ? hoja donde está la muestra
    
    ' Leer tamaño y universo
    On Error Resume Next
    tamano = CLng(wb.Names(nombreTamano).RefersToRange.Value)
    universo = CLng(wb.Names(nombreUniverso).RefersToRange.Value)
    On Error GoTo 0
    
    If tamano <= 0 Or universo <= 0 Then
        MsgBox "Los valores de '" & nombreTamano & "' y '" & nombreUniverso & "' deben ser > 0.", vbExclamation
        Exit Sub
    End If
    If tamano > universo Then
        MsgBox "'" & nombreTamano & "' no puede ser mayor que '" & nombreUniverso & "'.", vbExclamation
        Exit Sub
    End If
    
    startRow = rngInicio.Row
    startCol = rngInicio.Column
    
    ' Determinar el área usada anteriormente dentro del bloque de 5 columnas
    lastColUsed = startCol + 4
    lastRowUsed = startRow
    For c = startCol To lastColUsed
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > lastRowUsed Then lastRowUsed = r
    Next c
    If lastRowUsed < startRow Then lastRowUsed = startRow
    
    ' Limpiar contenidos previos (preservando formato de la celda plantilla)
    With ws
        .Range(.Cells(startRow, startCol), .Cells(lastRowUsed, lastColUsed)).ClearContents
        ' borrar formatos a la derecha de la plantilla
        If lastColUsed > startCol Then
            .Range(.Cells(startRow, startCol + 1), .Cells(lastRowUsed, lastColUsed)).ClearFormats
        End If
        ' borrar formatos debajo de la plantilla en su misma columna
        If lastRowUsed > startRow Then
            .Range(.Cells(startRow + 1, startCol), .Cells(lastRowUsed, startCol)).ClearFormats
        End If
        rngInicio.ClearContents   ' la plantilla pierde valor, conserva formato
    End With
    
    ' Generar números aleatorios ÚNICOS
    Set coll = New Collection
    Randomize
    Do While coll.Count < tamano
        rnum = Int(universo * Rnd) + 1
        On Error Resume Next
        coll.Add rnum, CStr(rnum)   ' key = número -> evita duplicados
        On Error GoTo 0
    Loop
    
    ' Pasar a array y ordenar ascendente (menor primero)
    ReDim numeros(1 To coll.Count)
    For i = 1 To coll.Count
        numeros(i) = coll(i)
    Next i
    For i = 1 To UBound(numeros) - 1
        For j = i + 1 To UBound(numeros)
            If numeros(i) > numeros(j) Then
                tmp = numeros(i): numeros(i) = numeros(j): numeros(j) = tmp
            End If
        Next j
    Next i
    
    ' Escribir en 5 columnas de ancho y luego hacia abajo, copiando formato de la plantilla
    fila = startRow
    col = startCol
    For i = 1 To UBound(numeros)
        ws.Cells(fila, col).Value = numeros(i)
        rngInicio.Copy
        ws.Cells(fila, col).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        
        col = col + 1
        If col > startCol + 4 Then
            col = startCol
            fila = fila + 1
        End If
    Next i
End Sub

