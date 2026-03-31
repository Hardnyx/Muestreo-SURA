Sub CargarDatos()
    Dim wbOrigen As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tbl As ListObject
    Dim rutaArchivo As String
    Dim ultimaFila As Long, ultimaColumna As Long
    Dim rngDatos As Range
    Dim fila As Long, colNombre As Long
    
    ' Seleccionar archivo XLS
    rutaArchivo = Application.GetOpenFilename("Archivos Excel (*.xls;*.xlsx), *.xls;*.xlsx", , "Seleccionar archivo con datos")
    If rutaArchivo = "Falso" Or rutaArchivo = "False" Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Abrir el archivo origen
    Set wbOrigen = Workbooks.Open(rutaArchivo, ReadOnly:=True)
    Set wsOrigen = wbOrigen.Sheets(1)
    
    ' Detectar última fila y columna con datos
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    ultimaColumna = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
    
    ' Definir rango de datos completo
    Set rngDatos = wsOrigen.Range(wsOrigen.Cells(1, 1), wsOrigen.Cells(ultimaFila, ultimaColumna))
    
    ' Encontrar columna "Nombre"
    On Error Resume Next
    colNombre = Application.Match("Nombre", rngDatos.Rows(1), 0)
    On Error GoTo 0
    If colNombre = 0 Then
        MsgBox "No se encontró la columna 'Nombre' en el archivo origen.", vbCritical
        GoTo Salir
    End If
    
    ' Eliminar filas con "Número de Cuentas:" en columna "Nombre"
    For fila = ultimaFila To 2 Step -1 ' asumimos encabezado en fila 1
        If Trim(CStr(wsOrigen.Cells(fila, colNombre).Value)) Like "*Número de Cuentas:*" Then
            wsOrigen.Rows(fila).Delete
        End If
    Next fila
    
    ' Actualizar última fila después de borrar
    ultimaFila = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
    Set rngDatos = wsOrigen.Range(wsOrigen.Cells(1, 1), wsOrigen.Cells(ultimaFila, ultimaColumna))
    
    ' Copiar datos al destino
    Set wsDestino = ThisWorkbook.Sheets("Contratos") ' Cambia por la hoja donde esté la tabla Contratos
    Set tbl = wsDestino.ListObjects("Contratos")
    
    ' Borrar todas las filas actuales de la tabla (mantener la tabla)
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    
    ' Redimensionar tabla a nuevo rango
    tbl.Resize tbl.Range.Resize(rngDatos.Rows.Count, rngDatos.Columns.Count)
    
    ' Copiar encabezado y datos
    tbl.HeaderRowRange.Value = rngDatos.Rows(1).Value
    If rngDatos.Rows.Count > 1 Then
        tbl.DataBodyRange.Value = rngDatos.Offset(1, 0).Resize(rngDatos.Rows.Count - 1).Value
    End If
    
    ' Cerrar archivo origen sin guardar
    wbOrigen.Close False
    
    MsgBox "Datos cargados correctamente en la tabla Contratos.", vbInformation

Salir:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

