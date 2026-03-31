Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngMonitoreo As Range
    
    ' Unir las 3 celdas nombradas en un solo rango
    Set rngMonitoreo = Union( _
        Me.Parent.Names("Mes").RefersToRange, _
        Me.Parent.Names("Año").RefersToRange, _
        Me.Parent.Names("TipoInforme").RefersToRange)
    
    ' Si cambia alguna de esas celdas, ejecuta TamañoPoblacion
    If Not Intersect(Target, rngMonitoreo) Is Nothing Then
        Application.EnableEvents = False
        Call TamañoPoblacion
        Application.EnableEvents = True
    End If
End Sub
