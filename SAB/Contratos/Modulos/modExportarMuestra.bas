' ========== modPQ_Contratos.bas ==========
Option Explicit

' ============================================================
'  ENTRADA DEL BOTÓN "Importar Datos"
' ============================================================
Public Sub ImportarDatos()
    CargarContratos_PQ
End Sub

Public Sub CargarContratos_PQ()
    Dim ruta As String
    Dim ws As Worksheet

    ruta = PickFilePath()
    If Len(ruta) = 0 Then Exit Sub

    On Error GoTo ERR_HANDLER
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.CutCopyMode = False

    ResetContratosEnvironment ws
    UpsertQuery "Contratos", M_Contratos_PQ(ruta)

    Dim connStr As String, cmdText As String
    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Contratos;Extended Properties=" & Chr(34) & Chr(34)
    cmdText = "SELECT * FROM [Contratos]"

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connStr, Destination:=ws.Range("A1"))
    With lo
        .name = "Contratos"
        With .QueryTable
            .CommandType = xlCmdSql
            .CommandText = cmdText
            .AdjustColumnWidth = True
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .MaintainConnection = True
            .RefreshStyle = xlOverwriteCells
            .SaveData = False
            .Refresh
        End With
    End With

    On Error Resume Next
    TamañoPoblacion
    On Error GoTo 0

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Consulta 'Contratos' cargada correctamente.", vbInformation
    Exit Sub

ERR_HANDLER:
    Dim errDesc As String: errDesc = Err.Description
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error al cargar los datos:" & vbCrLf & vbCrLf & errDesc, _
           vbCritical, "Error de importaci" & Chr(243) & "n"
End Sub

' ============================================================
'  RESET
' ============================================================
Private Sub ResetContratosEnvironment(ByRef wsOut As Worksheet)
    Dim ws As Worksheet, lo As ListObject, cN As WorkbookConnection

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        Do While ws.QueryTables.Count > 0
            If InStr(1, ws.QueryTables(1).Connection, "Location=Contratos", vbTextCompare) > 0 _
               Or InStr(1, ws.QueryTables(1).Connection, "Microsoft.Mashup.OleDb.1", vbTextCompare) > 0 Then
                ws.QueryTables(1).Delete
            Else
                Exit Do
            End If
        Loop
        Set lo = Nothing
        Set lo = ws.ListObjects("Contratos")
        If Not lo Is Nothing Then lo.Unlist
    Next ws

    For Each cN In ThisWorkbook.Connections
        If cN.Type = xlConnectionTypeOLEDB Then
            If InStr(1, cN.OLEDBConnection.Connection, "Location=Contratos", vbTextCompare) > 0 Then
                cN.Delete
            End If
        End If
    Next cN

    ThisWorkbook.Queries("Contratos").Delete
    On Error GoTo 0

    Set wsOut = Nothing
    On Error Resume Next
    Set wsOut = ThisWorkbook.Worksheets("Contratos")
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.name = "Contratos"
    Else
        wsOut.Cells.Clear
    End If
End Sub

Private Sub UpsertQuery(ByVal qName As String, ByVal mCode As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = ThisWorkbook.Queries(qName)
    On Error GoTo 0
    If q Is Nothing Then
        ThisWorkbook.Queries.Add name:=qName, Formula:=mCode
    Else
        q.Formula = mCode
    End If
End Sub

Private Function PickFilePath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo de Contratos (.XLS, .CSV, .TXT)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Archivos comunes", "*.xls; *.xlsx; *.xlsm; *.xlsb; *.csv; *.txt"
        If .Show <> -1 Then Exit Function
        PickFilePath = .SelectedItems(1)
    End With
End Function

' ============================================================
'  M FORMULA
'  27 columnas reales del archivo de Contratos.
'  Columnas clave: "Tipo" (N/J), "Fecha de Ingreso" (date).
'  Sort: Fecha de Ingreso ASC, Cuenta ASC.
' ============================================================
Private Function M_Contratos_PQ(ByVal ruta As String) As String
    Dim m As String, p As String
    p = Replace(ruta, """", """""")

    ' Columnas con caracteres especiales construidas con Chr()
    Dim cClasif1  As String: cClasif1 = "Clasificaci" & Chr(243) & "n 1"
    Dim cClasif2  As String: cClasif2 = "Clasificaci" & Chr(243) & "n 2"
    Dim cDirPrec  As String: cDirPrec = "Direcci" & Chr(243) & "n Precisa"
    Dim cDirCont  As String: cDirCont = "Direcci" & Chr(243) & "n de Contacto"
    Dim cTelefono As String: cTelefono = "Tel" & Chr(233) & "fono"
    Dim cPais     As String: cPais = "Pa" & Chr(237) & "s"
    Dim cEnvio    As String: cEnvio = "Lugar de Env" & Chr(237) & "o de Correspondencia"
    Dim cFechaIng As String: cFechaIng = "Fecha de Ingreso"
    Dim cFecBloq  As String: cFecBloq = "Fecha de Bloqueo"

    m = "let" & vbCrLf
    m = m & "  path = """ & p & """," & vbCrLf & vbCrLf

    ' 27 columnas en el orden exacto del archivo
    m = m & "  expected = {""Cuenta"",""Tipo"",""Nombre"",""RUC/NIT""," & vbCrLf
    m = m & "               """ & cClasif1 & """,""" & cClasif2 & """," & vbCrLf
    m = m & "               """ & cDirPrec & """,""" & cDirCont & """," & vbCrLf
    m = m & "               """ & cTelefono & """,""Celular"",""Fax"",""Casilla"",""Email""," & vbCrLf
    m = m & "               """ & cEnvio & """," & vbCrLf
    m = m & "               ""Oficial de Cuenta"",""Referencia"",""" & cFechaIng & """,""" & cPais & """,""Distrito""," & vbCrLf
    m = m & "               ""C Entero"",""Conoc Merc"",""Estado"",""Tipo Bloqueo"",""" & cFecBloq & """," & vbCrLf
    m = m & "               ""Observaciones del Agente"",""Tipo de Cliente"",""Vinculado a Agente""}," & vbCrLf & vbCrLf

    ' Canon: normaliza a mayúsculas sin tildes ni separadores
    m = m & "  Canon = (s as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      t0 = Text.Upper(Text.Trim(s))," & vbCrLf
    m = m & "      t1 = Text.Replace(t0,""" & Chr(193) & """,""A"")," & vbCrLf
    m = m & "      t2 = Text.Replace(t1,""" & Chr(201) & """,""E"")," & vbCrLf
    m = m & "      t3 = Text.Replace(t2,""" & Chr(205) & """,""I"")," & vbCrLf
    m = m & "      t4 = Text.Replace(t3,""" & Chr(211) & """,""O"")," & vbCrLf
    m = m & "      t5 = Text.Replace(t4,""" & Chr(218) & """,""U"")," & vbCrLf
    m = m & "      t6 = Text.Replace(t5,""" & Chr(209) & """,""N"")," & vbCrLf
    m = m & "      t7 = Text.Replace(Text.Replace(t6,""N" & Chr(186) & """,""N""),""N" & Chr(176) & """,""N"")," & vbCrLf
    m = m & "      t8 = Text.Replace("" "" & t7 & "" "","" DE "","" "")," & vbCrLf
    m = m & "      out = Text.Remove(Text.Trim(t8), {"" "",""_"",""-"",""."",""/"",""\""})" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      out," & vbCrLf & vbCrLf

    m = m & "  expectedCanon = List.Transform(expected, each Canon(_))," & vbCrLf
    m = m & "  lenExp = List.Count(expected)," & vbCrLf & vbCrLf

    m = m & "  bin = Binary.Buffer(File.Contents(path))," & vbCrLf
    m = m & "  encodings = {65001, 1252}," & vbCrLf
    m = m & "  delims    = {" & Chr(34) & "," & Chr(34) & "," & Chr(34) & "#(tab)" & Chr(34) & "," & Chr(34) & "|" & Chr(34) & "}," & vbCrLf & vbCrLf

    m = m & "  MakeTable = (enc as number, delim as text) as nullable table =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      t = try Csv.Document(bin, [Delimiter=delim, Columns=null, Encoding=enc, QuoteStyle=QuoteStyle.Csv]) otherwise null" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      if t = null then null else t," & vbCrLf & vbCrLf

    m = m & "  RowIsExactHeader = (row as list) as logical =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      rTxt  = List.Transform(row, each if _ = null then """" else Text.From(_))," & vbCrLf
    m = m & "      rCan  = List.Transform(rTxt, each Canon(_))," & vbCrLf
    m = m & "      okLen = List.Count(rCan) >= lenExp," & vbCrLf
    m = m & "      slice = if okLen then List.FirstN(rCan, lenExp) else rCan," & vbCrLf
    m = m & "      eqAll = okLen and List.AllTrue(List.Transform({0..lenExp-1}, each slice{_} = expectedCanon{_}))" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      eqAll," & vbCrLf & vbCrLf

    m = m & "  FindHeaderRowIndex = (tbl as table, maxScan as number) as any =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      sample = Table.FirstN(tbl, maxScan)," & vbCrLf
    m = m & "      rows   = Table.ToRows(sample)," & vbCrLf
    m = m & "      idx    = List.PositionOf(List.Transform(rows, each RowIsExactHeader(_)), true)" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      idx," & vbCrLf & vbCrLf

    m = m & "  TryAll =" & vbCrLf
    m = m & "    List.First(" & vbCrLf
    m = m & "      List.RemoveNulls(" & vbCrLf
    m = m & "        List.Transform(" & vbCrLf
    m = m & "          encodings," & vbCrLf
    m = m & "          (enc) =>" & vbCrLf
    m = m & "            List.First(" & vbCrLf
    m = m & "              List.RemoveNulls(" & vbCrLf
    m = m & "                List.Transform(" & vbCrLf
    m = m & "                  delims," & vbCrLf
    m = m & "                  (delim) =>" & vbCrLf
    m = m & "                    let" & vbCrLf
    m = m & "                      t0  = MakeTable(enc, delim)," & vbCrLf
    m = m & "                      idx = if t0 = null then -1 else FindHeaderRowIndex(t0, 120)," & vbCrLf
    m = m & "                      rec = if t0 <> null and idx >= 0 then [Enc=enc, Delim=delim, Tbl=t0, HeaderIdx=idx] else null" & vbCrLf
    m = m & "                    in" & vbCrLf
    m = m & "                      rec" & vbCrLf
    m = m & "                )" & vbCrLf
    m = m & "              )" & vbCrLf
    m = m & "            )" & vbCrLf
    m = m & "        )" & vbCrLf
    m = m & "      )," & vbCrLf
    m = m & "      null" & vbCrLf
    m = m & "    )," & vbCrLf & vbCrLf

    m = m & "  tblAll = if TryAll = null" & vbCrLf
    m = m & "           then error ""No se encontr" & Chr(243) & " la fila con las 27 cabeceras esperadas en el archivo.""" & vbCrLf
    m = m & "           else TryAll[Tbl]," & vbCrLf
    m = m & "  hIdx   = TryAll[HeaderIdx]," & vbCrLf & vbCrLf

    m = m & "  tblAfterSkip = Table.Skip(tblAll, hIdx)," & vbCrLf
    m = m & "  promoted     = Table.PromoteHeaders(tblAfterSkip, [PromoteAllScalars=true])," & vbCrLf & vbCrLf

    m = m & "  curNames = Table.ColumnNames(promoted)," & vbCrLf
    m = m & "  pairs    = List.Transform({0..lenExp-1}, each {curNames{_}, expected{_}})," & vbCrLf
    m = m & "  renamed  = Table.RenameColumns(promoted, pairs, MissingField.Ignore)," & vbCrLf & vbCrLf

    m = m & "  only27   = Table.SelectColumns(renamed, expected, MissingField.UseNull)," & vbCrLf & vbCrLf

    ' Fecha de Ingreso y Fecha de Bloqueo como date; todo lo demás como text
    m = m & "  allTextCols = List.RemoveItems(expected, {""" & cFechaIng & """,""" & cFecBloq & """})," & vbCrLf
    m = m & "  AsText = Table.TransformColumnTypes(only27, List.Transform(allTextCols, each {_, type text}), ""es-PE"")," & vbCrLf & vbCrLf

    ' Parseo de fecha DDMMMYYYY
    m = m & "  MonthPairs = {" & vbCrLf
    m = m & "    ""ene"",""enero"",""feb"",""febrero"",""mar"",""marzo"",""abr"",""abril"",""may"",""mayo"",""jun"",""junio""," & vbCrLf
    m = m & "    ""jul"",""julio"",""ago"",""agosto"",""set"",""septiembre"",""sep"",""septiembre"",""sept"",""septiembre""," & vbCrLf
    m = m & "    ""oct"",""octubre"",""nov"",""noviembre"",""dic"",""diciembre""" & vbCrLf
    m = m & "  }," & vbCrLf & vbCrLf

    m = m & "  NormalizeMonthText = (t as text) as text =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      s0 = Text.Trim(t)," & vbCrLf
    m = m & "      s1 = Text.Lower(s0)," & vbCrLf
    ' Detectar formato compacto [1-2 dígitos][3 letras mes][2-4 dígitos año]
    ' Primero intenta día de 2 dígitos
    m = m & "      validMons = {""ene"",""feb"",""mar"",""abr"",""may"",""jun"",""jul"",""ago"",""set"",""sep"",""oct"",""nov"",""dic""}," & vbCrLf
    m = m & "      mon2 = Text.Middle(s1, 2, 3)," & vbCrLf
    m = m & "      yr2  = Text.End(s1, Text.Length(s1) - 5)," & vbCrLf
    m = m & "      ok2  = Text.Length(s1) >= 7 and (try Number.From(Text.Start(s1,2)) >= 1 otherwise false)" & vbCrLf
    m = m & "             and List.Contains(validMons, mon2) and (try Number.From(yr2) >= 1 otherwise false)," & vbCrLf
    ' Luego intenta día de 1 dígito
    m = m & "      mon1 = Text.Middle(s1, 1, 3)," & vbCrLf
    m = m & "      yr1  = Text.End(s1, Text.Length(s1) - 4)," & vbCrLf
    m = m & "      ok1  = Text.Length(s1) >= 6 and (try Number.From(Text.Start(s1,1)) >= 1 otherwise false)" & vbCrLf
    m = m & "             and List.Contains(validMons, mon1) and (try Number.From(yr1) >= 1 otherwise false)," & vbCrLf
    m = m & "      dayPart = if ok2 then Text.Start(s1, 2) else if ok1 then Text.Start(s1, 1) else """"," & vbCrLf
    m = m & "      monPart = if ok2 then mon2 else if ok1 then mon1 else """"," & vbCrLf
    m = m & "      yrRaw   = if ok2 then yr2 else if ok1 then yr1 else """"," & vbCrLf
    ' Normalizar año: 2 dígitos ? 20XX o 19XX; 3 dígitos ? "026"?"2026"; 4 dígitos ? tal cual
    m = m & "      yrNorm  = if Text.Length(yrRaw) = 2 then" & vbCrLf
    m = m & "                  (if Number.FromText(yrRaw) < 50 then ""20"" else ""19"") & yrRaw" & vbCrLf
    m = m & "                else if Text.Length(yrRaw) = 3 then" & vbCrLf
    m = m & "                  ""20"" & Text.End(yrRaw, 2)" & vbCrLf
    m = m & "                else yrRaw," & vbCrLf
    m = m & "      sExp    = if ok2 or ok1 then dayPart & "" "" & monPart & "" "" & yrNorm else s1," & vbCrLf
    ' Normalización estándar del resto de formatos
    m = m & "      s2 = Text.Replace(Text.Replace(Text.Replace(sExp, ""/"", "" ""), ""-"", "" ""), ""."", """")," & vbCrLf
    m = m & "      parts = List.Select(Text.Split(s2, "" ""), each _ <> """")," & vbCrLf
    m = m & "      s3 = Text.Combine(parts, "" "")," & vbCrLf
    m = m & "      s4 = "" "" & s3 & "" ""," & vbCrLf
    m = m & "      s5 = List.Accumulate({0..List.Count(MonthPairs)/2-1}, s4, (state, i) => Text.Replace(state, "" "" & MonthPairs{2*i} & "" "", "" "" & MonthPairs{2*i+1} & "" ""))," & vbCrLf
    m = m & "      res = Text.Trim(s5)" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      res," & vbCrLf & vbCrLf

    m = m & "  ParseFechaES = (v as any) as nullable date =>" & vbCrLf
    m = m & "    let" & vbCrLf
    m = m & "      out =" & vbCrLf
    m = m & "        if v is date then Date.From(v) else" & vbCrLf
    m = m & "        if v is datetime then Date.From(v) else" & vbCrLf
    m = m & "        if v is number then Date.From(#datetime(1899,12,30,0,0,0) + #duration(Number.From(v),0,0,0)) else" & vbCrLf
    m = m & "        if v is text then" & vbCrLf
    m = m & "          let" & vbCrLf
    m = m & "            n   = NormalizeMonthText(v)," & vbCrLf
    m = m & "            d1  = try Date.FromText(n, ""es-PE"") otherwise try Date.FromText(n, ""es-ES"") otherwise null," & vbCrLf
    m = m & "            res = if d1 <> null then d1 else" & vbCrLf
    m = m & "                    let" & vbCrLf
    m = m & "                      parts = Text.Split(n, "" "")," & vbCrLf
    m = m & "                      dS = if List.Count(parts) > 0 then parts{0} else """"," & vbCrLf
    m = m & "                      mS = if List.Count(parts) > 1 then parts{1} else """"," & vbCrLf
    m = m & "                      yS = if List.Count(parts) > 2 then parts{2} else """"," & vbCrLf
    m = m & "                      dN = try Number.FromText(dS) otherwise null," & vbCrLf
    m = m & "                      months = {""enero"",""febrero"",""marzo"",""abril"",""mayo"",""junio"",""julio"",""agosto"",""septiembre"",""octubre"",""noviembre"",""diciembre""}," & vbCrLf
    m = m & "                      mPos = List.PositionOf(months, mS)," & vbCrLf
    m = m & "                      mN = if mPos >= 0 then mPos + 1 else null," & vbCrLf
    m = m & "                      yN0 = try Number.FromText(yS) otherwise null," & vbCrLf
    m = m & "                      yN  = if yN0 is number and Text.Length(yS)=2 then (if yN0 < 50 then 2000 + yN0 else 1900 + yN0) else yN0," & vbCrLf
    m = m & "                      d2 = if dN <> null and mN <> null and yN <> null then #date(yN, mN, dN) else null" & vbCrLf
    m = m & "                    in" & vbCrLf
    m = m & "                      d2" & vbCrLf
    m = m & "          in" & vbCrLf
    m = m & "            res" & vbCrLf
    m = m & "        else" & vbCrLf
    m = m & "          null" & vbCrLf
    m = m & "    in" & vbCrLf
    m = m & "      out," & vbCrLf & vbCrLf

    ' Aplicar parseo a Fecha de Ingreso y Fecha de Bloqueo
    m = m & "  FechaIngFix = Table.TransformColumns(AsText, {{""" & cFechaIng & """, each ParseFechaES(_), type date}})," & vbCrLf
    m = m & "  FechaFix    = Table.TransformColumns(FechaIngFix, {{""" & cFecBloq & """, each ParseFechaES(_), type date}})," & vbCrLf & vbCrLf

    ' Eliminar filas donde Fecha de Ingreso es null (totales o filas vacías)
    m = m & "  NoNulls = Table.SelectRows(FechaFix, each Record.Field(_, """ & cFechaIng & """) <> null)," & vbCrLf & vbCrLf

    ' Ordenar por Fecha de Ingreso ASC, Cuenta ASC como tiebreaker
    m = m & "  Sorted = Table.Sort(NoNulls, {{""" & cFechaIng & """, Order.Ascending}, {""Cuenta"", Order.Ascending}})" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "  Sorted" & vbCrLf

    M_Contratos_PQ = m
End Function


