Option Explicit

'----------------------------------------
' Evento de cambio en cualquier hoja
'----------------------------------------
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim lo As ListObject
    Dim nmNames As Variant
    Dim nmRange As Range
    Dim i As Long

    ' Evitar llamadas recursivas
    Application.EnableEvents = False

    ' 1) Si el cambio ocurrió en la hoja "Contratos" y afecta la tabla "Contratos" -> recalcular
    If Sh.name = "Contratos" Then
        On Error Resume Next
        Set lo = Sh.ListObjects("Contratos")
        On Error GoTo 0
        If Not lo Is Nothing Then
            ' Si el cambio impacta en cualquier celda de la tabla (cabecera, cuerpo o total)
            If Not Intersect(Target, lo.Range) Is Nothing Then
                TamañoPoblacion
                GoTo ExitHandler
            End If
        End If
    End If

    ' 2) Si el cambio ocurrió en la hoja "Muestra" y toca alguna de las celdas control -> recalcular
    If Sh.name = "Muestra" Then
        nmNames = Array("Mes", "Año", "TipoInforme")
        For i = LBound(nmNames) To UBound(nmNames)
            On Error Resume Next
            Set nmRange = wb.Names(nmNames(i)).RefersToRange
            On Error GoTo 0
            If Not nmRange Is Nothing Then
                If Not Intersect(Target, nmRange) Is Nothing Then
                    TamañoPoblacion
                    Exit For
                End If
            End If
        Next i
    End If

ExitHandler:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "Error en Workbook_SheetChange: " & Err.Number & " - " & Err.Description, vbCritical
    Resume ExitHandler
End Sub

'----------------------------------------
' Macro de carga de datos con actualización automática
'----------------------------------------
Public Sub CargarDatos()
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tbl As ListObject
    Dim rutaArchivo As String
    Dim ultimaFila As Long, ultimaColumna As Long
    Dim rngDatos As Range
    Dim fila As Long
    
    ' Seleccionar archivo XLS
    rutaArchivo = Application.GetOpenFilename("Archivos Excel (*.xls), *.xls", , "Seleccionar archivo con datos")
    If rutaArchivo = "Falso" Or rutaArchivo = "False" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Abrir archivo origen
    Set wbOrigen = Workbooks.Open(rutaArchivo, ReadOnly:=True)
    Set wsOrigen = wbOrigen.Sheets(1)
    
    ' Detectar última fila y columna con datos
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    ultimaColumna = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
    
    ' Definir rango completo
    Set rngDatos = wsOrigen.Range(wsOrigen.Cells(1, 1), wsOrigen.Cells(ultimaFila, ultimaColumna))
    
    ' Eliminar filas con "Número de Cuentas:" en columna "Nombre"
    For fila = ultimaFila To 1 Step -1
        If Trim(CStr(wsOrigen.Cells(fila, Application.Match("Nombre", rngDatos.Rows(1), 0)).Value)) Like "*Número de Cuentas:*" Then
            wsOrigen.Rows(fila).Delete
        End If
    Next fila
    
    ' Actualizar rango después de borrar filas
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    Set rngDatos = wsOrigen.Range(wsOrigen.Cells(1, 1), wsOrigen.Cells(ultimaFila, ultimaColumna))
    
    ' Copiar datos al destino
    Set wsDestino = ThisWorkbook.Sheets("Contratos") ' hoja con la tabla Contratos
    Set tbl = wsDestino.ListObjects("Contratos")
    
    ' Borrar filas existentes (mantener tabla)
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    
    ' Redimensionar tabla al nuevo rango
    tbl.Resize tbl.Range.Resize(rngDatos.Rows.Count, rngDatos.Columns.Count)
    
    ' Pegar encabezado y datos
    tbl.HeaderRowRange.Value = rngDatos.Rows(1).Value
    tbl.DataBodyRange.Value = rngDatos.Offset(1, 0).Resize(rngDatos.Rows.Count - 1).Value
    
    ' Cerrar archivo origen sin guardar
    wbOrigen.Close False
    
    ' Actualizar variables dependientes automáticamente
    Call TamañoPoblacion
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Datos cargados correctamente en la tabla Contratos y variables actualizadas.", vbInformation
End Sub