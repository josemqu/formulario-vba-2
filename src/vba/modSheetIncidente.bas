Attribute VB_Name = "modSheetIncidente"
Option Explicit

Private Function EnsureFormSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Formulario")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Formulario"
    End If
    Set EnsureFormSheet = ws
End Function

Private gAppEv As clsAppEvents

Private Sub LayoutForm(ws As Worksheet)
    ws.Range("B2").Value = "ID incidente"
    ws.Range("B3").Value = "Fecha/hora ocurrencia"
    ws.Range("B4").Value = "País"
    ws.Range("B5").Value = "Provincia"
    ws.Range("B6").Value = "Localidad/Zona"
    ws.Range("B7").Value = "Coordenadas"
    ws.Range("B8").Value = "Lugar específico"
    ws.Range("B9").Value = "UO incidente"
    ws.Range("B10").Value = "UO accidentado"
    ws.Range("B11").Value = "Descripción"
    ws.Range("B12").Value = "Denuncia policial"
    ws.Range("B13").Value = "Examen alcoholemia"
    ws.Range("B14").Value = "Examen sustancias"
    ws.Range("B15").Value = "Entrevistas testigos"
    ws.Range("B16").Value = "Acción inmediata"
    ws.Range("B17").Value = "Consecuencias seguridad"
    ws.Range("B18").Value = "Fecha/hora reporte"
    ws.Range("B19").Value = "Cantidad personas"
    ws.Range("B20").Value = "Cantidad vehículos"
    ws.Range("B21").Value = "Clase evento"
    ws.Range("B22").Value = "Tipo colisión"
    ws.Range("B23").Value = "Nivel severidad"
    ws.Range("B24").Value = "Clasificación ESV"
    If LenB(CStr(ws.Range("C3").Value)) = 0 Then ws.Range("C3").Value = Now
    If LenB(CStr(ws.Range("C18").Value)) = 0 Then ws.Range("C18").Value = Now
    ws.Columns("B:B").ColumnWidth = 26
    ws.Columns("C:C").ColumnWidth = 50
    ws.Range("B2:B24").WrapText = True
    ws.Range("C3,C18").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("C19:C20").NumberFormat = "0"
End Sub

Private Sub ApplyValidations(ws As Worksheet)
    Dim r As Range
    On Error Resume Next
    Set r = ws.Range("C4"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_PAIS"
    Set r = ws.Range("C5"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_PROVINCIA"
    Set r = ws.Range("C6"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_LOCALIDAD_ZONA"
    Set r = ws.Range("C9"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_UO_INCIDENTE"
    Set r = ws.Range("C10"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_UO_ACCIDENTADO"
    Set r = ws.Range("C12"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = ws.Range("C13"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = ws.Range("C14"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = ws.Range("C15"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = ws.Range("C17"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = ws.Range("C21"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_CLASE_EVENTO"
    Set r = ws.Range("C22"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_TIPO_COLISION"
    Set r = ws.Range("C23"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_NIVEL_SEVERIDAD"
    Set r = ws.Range("C24"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_CLASIFICACION_ESV"
    On Error GoTo 0
End Sub

Private Sub EnsureGuardarButton(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("btnGuardarIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("B26").Left, ws.Range("B26").Top, 160, 32)
        shp.Name = "btnGuardarIncidente"
        shp.TextFrame.Characters.Text = "Guardar incidente"
        shp.OnAction = "GuardarIncidenteDesdeHoja"
    Else
        shp.OnAction = "GuardarIncidenteDesdeHoja"
    End If
    On Error Resume Next
    Set shp = ws.Shapes("btnNuevoIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D26").Left, ws.Range("D26").Top, 120, 32)
        shp.Name = "btnNuevoIncidente"
        shp.TextFrame.Characters.Text = "Nuevo"
        shp.OnAction = "NuevoIncidenteEnHoja"
    Else
        shp.OnAction = "NuevoIncidenteEnHoja"
    End If
End Sub

Public Sub AbrirFormularioIncidenteEnHoja()
    SetupESVWorkbook
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    LayoutForm ws
    ApplyValidations ws
    EnsureGuardarButton ws
    If gAppEv Is Nothing Then
        Set gAppEv = New clsAppEvents
        Set gAppEv.App = Application
    End If
    EstilizarFormularioIncidente
    ws.Activate
End Sub

Private Function ReadIncidenteFromSheet(ws As Worksheet) As clsIncidente
    Dim e As New clsIncidente
    e.id_incidente = CStr(ws.Range("C2").Value)
    e.fecha_hora_ocurrencia = ws.Range("C3").Value
    e.pais = CStr(ws.Range("C4").Value)
    e.provincia = CStr(ws.Range("C5").Value)
    e.localidad_zona = CStr(ws.Range("C6").Value)
    e.coordenadas_geograficas = CStr(ws.Range("C7").Value)
    e.lugar_especifico = CStr(ws.Range("C8").Value)
    e.uo_incidente = CStr(ws.Range("C9").Value)
    e.uo_accidentado = CStr(ws.Range("C10").Value)
    e.descripcion_esv = CStr(ws.Range("C11").Value)
    e.denuncia_policial = CStr(ws.Range("C12").Value)
    e.examen_alcoholemia = CStr(ws.Range("C13").Value)
    e.examen_sustancias = CStr(ws.Range("C14").Value)
    e.entrevistas_testigos = CStr(ws.Range("C15").Value)
    e.accion_inmediata = CStr(ws.Range("C16").Value)
    e.consecuencias_seguridad = CStr(ws.Range("C17").Value)
    e.fecha_hora_reporte = ws.Range("C18").Value
    e.cantidad_personas = ws.Range("C19").Value
    e.cantidad_vehiculos = ws.Range("C20").Value
    e.clase_evento = CStr(ws.Range("C21").Value)
    e.tipo_colision = CStr(ws.Range("C22").Value)
    e.nivel_severidad = CStr(ws.Range("C23").Value)
    e.clasificacion_esv = CStr(ws.Range("C24").Value)
    Set ReadIncidenteFromSheet = e
End Function

Private Sub ClearForm(ws As Worksheet)
    ws.Range("C2:C24").ClearContents
    ws.Range("C3").Value = Now
    ws.Range("C18").Value = Now
End Sub

Private Function ValidateForm(ws As Worksheet, ByRef messages As Collection) As Boolean
    Dim ok As Boolean: ok = True
    Set messages = New Collection
    If LenB(CStr(ws.Range("C3").Value)) = 0 Then ok = False: messages.Add("Fecha/hora ocurrencia es requerida.")
    If LenB(CStr(ws.Range("C4").Value)) = 0 Then ok = False: messages.Add("País es requerido.")
    If LenB(CStr(ws.Range("C21").Value)) = 0 Then ok = False: messages.Add("Clase de evento es requerida.")
    If LenB(CStr(ws.Range("C19").Value)) > 0 Then If Not IsNumeric(ws.Range("C19").Value) Then ok = False: messages.Add("Cantidad personas debe ser numérico.")
    If LenB(CStr(ws.Range("C20").Value)) > 0 Then If Not IsNumeric(ws.Range("C20").Value) Then ok = False: messages.Add("Cantidad vehículos debe ser numérico.")
    ValidateForm = ok
End Function

Public Sub GuardarIncidenteDesdeHoja()
    SetupESVWorkbook
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    Dim msgs As Collection
    If Not ValidateForm(ws, msgs) Then
        Dim t As String: t = "No se puede guardar. Corrige los siguientes puntos:" & vbCrLf
        Dim it As Variant
        For Each it In msgs
            t = t & "- " & CStr(it) & vbCrLf
        Next it
        MsgBox t, vbExclamation
        Exit Sub
    End If
    Dim e As clsIncidente
    Set e = ReadIncidenteFromSheet(ws)
    Dim id As String
    id = clsIncidenteRepo.SaveEntity(e)
    ws.Range("C2").Value = id
    MsgBox "Incidente guardado: " & id, vbInformation
End Sub

Public Sub NuevoIncidenteEnHoja()
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    ClearForm ws
End Sub

Public Sub LoadIncidenteEnHojaDesdeIdActual()
    SetupESVWorkbook
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    Dim id As String: id = CStr(ws.Range("C2").Value)
    If LenB(id) = 0 Then Exit Sub
    Dim e As clsIncidente
    Set e = clsIncidenteRepo.FindById(id)
    If e Is Nothing Then Exit Sub
    ws.Range("C3").Value = e.fecha_hora_ocurrencia
    ws.Range("C4").Value = e.pais
    ws.Range("C5").Value = e.provincia
    ws.Range("C6").Value = e.localidad_zona
    ws.Range("C7").Value = e.coordenadas_geograficas
    ws.Range("C8").Value = e.lugar_especifico
    ws.Range("C9").Value = e.uo_incidente
    ws.Range("C10").Value = e.uo_accidentado
    ws.Range("C11").Value = e.descripcion_esv
    ws.Range("C12").Value = e.denuncia_policial
    ws.Range("C13").Value = e.examen_alcoholemia
    ws.Range("C14").Value = e.examen_sustancias
    ws.Range("C15").Value = e.entrevistas_testigos
    ws.Range("C16").Value = e.accion_inmediata
    ws.Range("C17").Value = e.consecuencias_seguridad
    ws.Range("C18").Value = e.fecha_hora_reporte
    ws.Range("C19").Value = e.cantidad_personas
    ws.Range("C20").Value = e.cantidad_vehiculos
    ws.Range("C21").Value = e.clase_evento
    ws.Range("C22").Value = e.tipo_colision
    ws.Range("C23").Value = e.nivel_severidad
    ws.Range("C24").Value = e.clasificacion_esv
End Sub

Public Sub EstilizarFormularioIncidente()
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    ws.Range("B2:B24").Font.Bold = True
    ws.Range("B2:B24").Interior.Color = RGB(245, 245, 245)
    ws.Range("C2:C24").Interior.Color = RGB(255, 255, 255)
    With ws.Range("B2:C24").Borders
        .LineStyle = xlContinuous
        .Color = RGB(220, 220, 220)
        .Weight = xlThin
    End With
    ws.Range("B2:C24").Borders(xlInsideHorizontal).Color = RGB(235, 235, 235)
    ws.Range("B2:C24").Borders(xlInsideVertical).Color = RGB(235, 235, 235)
    ws.Rows("2:24").RowHeight = 20
    ws.Columns("B:C").HorizontalAlignment = xlLeft
    ws.Columns("C:C").HorizontalAlignment = xlLeft
    ws.Columns("C:C").VerticalAlignment = xlCenter
    ws.Range("B1:C1").Merge
    ws.Range("B1").Value = "Registro de Incidente"
    ws.Range("B1").Font.Size = 16
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Color = RGB(32, 32, 32)
    ws.Range("B1").Interior.Color = RGB(255, 255, 255)
    ws.Range("B1").EntireRow.RowHeight = 28
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes("btnGuardarIncidente")
    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = RGB(0, 120, 215)
        shp.Line.ForeColor.RGB = RGB(0, 84, 153)
        shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        shp.TextFrame.Characters.Font.Bold = True
    End If
    Set shp = ws.Shapes("btnNuevoIncidente")
    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = RGB(0, 153, 51)
        shp.Line.ForeColor.RGB = RGB(0, 102, 34)
        shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        shp.TextFrame.Characters.Font.Bold = True
    End If
    Set shp = Nothing
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

