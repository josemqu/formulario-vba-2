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
End Sub

Private Sub ApplyValidations(ws As Worksheet)
    Dim r As Range
    On Error Resume Next
    ' IDs existentes
    Set r = ws.Range("C2"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_ID_INCIDENTE"
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
End Sub

Public Sub AbrirFormularioIncidenteEnHoja()
    SetupESVWorkbook
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    LayoutForm ws
    EnsureIncidentIdCatalog
    ApplyValidations ws
    EnsureGuardarButton ws
    EnsureNuevoButton ws
    EnsureCargarButton ws
    FormatForm ws
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

Public Sub GuardarIncidenteDesdeHoja()
    SetupESVWorkbook
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    Dim e As clsIncidente
    Set e = ReadIncidenteFromSheet(ws)
    Dim msg As String
    msg = ValidateIncidente(ws)
    If LenB(msg) > 0 Then
        MsgBox msg, vbExclamation, "Validación"
        Exit Sub
    End If
    Dim id As String
    id = clsIncidenteRepo.SaveEntity(e)
    ws.Range("C2").Value = id
    EnsureIncidentIdCatalog
    ApplyValidations ws ' asegura lista de IDs actualizada
    MsgBox "Incidente guardado: " & id, vbInformation
End Sub

Private Sub ClearForm(ws As Worksheet)
    ws.Range("C2:C24").ClearContents
    ws.Range("C3").Value = Now
    ws.Range("C18").Value = Now
End Sub

Public Sub NuevoIncidenteDesdeHoja()
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    ClearForm ws
End Sub

Private Sub EnsureNuevoButton(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("btnNuevoIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("D26").Left, ws.Range("D26").Top, 120, 32)
        shp.Name = "btnNuevoIncidente"
        shp.TextFrame.Characters.Text = "Nuevo"
        shp.OnAction = "NuevoIncidenteDesdeHoja"
    Else
        shp.OnAction = "NuevoIncidenteDesdeHoja"
    End If
End Sub

Public Sub CargarIncidenteDesdeHoja()
    Dim ws As Worksheet
    Set ws = EnsureFormSheet()
    Dim id As String
    id = CStr(ws.Range("C2").Value)
    If LenB(id) = 0 Then
        MsgBox "Ingrese un ID en C2 para cargar.", vbExclamation
        Exit Sub
    End If
    Dim e As clsIncidente
    Set e = clsIncidenteRepo.FindById(id)
    If e Is Nothing Then
        MsgBox "ID no encontrado: " & id, vbExclamation
        Exit Sub
    End If
    WriteIncidenteToSheet ws, e
End Sub

Private Sub EnsureCargarButton(ws As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes("btnCargarIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("F26").Left, ws.Range("F26").Top, 120, 32)
        shp.Name = "btnCargarIncidente"
        shp.TextFrame.Characters.Text = "Cargar por ID"
        shp.OnAction = "CargarIncidenteDesdeHoja"
    Else
        shp.OnAction = "CargarIncidenteDesdeHoja"
    End If
End Sub

Private Sub WriteIncidenteToSheet(ws As Worksheet, ByVal e As clsIncidente)
    ws.Range("C2").Value = e.id_incidente
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

Private Function ValidateIncidente(ws As Worksheet) As String
    Dim errs As String
    Dim v
    ' Requeridos mínimos
    If LenB(CStr(ws.Range("C3").Value)) = 0 Then errs = errs & "- Fecha/hora ocurrencia es requerido" & vbCrLf
    If LenB(CStr(ws.Range("C4").Value)) = 0 Then errs = errs & "- País es requerido" & vbCrLf
    If LenB(CStr(ws.Range("C21").Value)) = 0 Then errs = errs & "- Clase de evento es requerido" & vbCrLf
    ' Tipos básicos
    v = ws.Range("C3").Value
    If LenB(CStr(v)) > 0 Then If Not IsDate(v) Then errs = errs & "- Fecha/hora ocurrencia debe ser fecha" & vbCrLf
    v = ws.Range("C18").Value
    If LenB(CStr(v)) > 0 Then If Not IsDate(v) Then errs = errs & "- Fecha/hora reporte debe ser fecha" & vbCrLf
    v = ws.Range("C19").Value
    If LenB(CStr(v)) > 0 Then If Not IsNumeric(v) Then errs = errs & "- Cantidad personas debe ser numérica" & vbCrLf
    v = ws.Range("C20").Value
    If LenB(CStr(v)) > 0 Then If Not IsNumeric(v) Then errs = errs & "- Cantidad vehículos debe ser numérica" & vbCrLf
    ValidateIncidente = errs
End Function

Private Sub FormatForm(ws As Worksheet)
    With ws
        .Range("A:Z").ColumnWidth = 15
        .Range("B:B").ColumnWidth = 24
        .Range("C:C").ColumnWidth = 40
        .Range("B2:B24").Font.Bold = True
        .Range("B2:B24").WrapText = True
        .Range("C11").WrapText = True
        .Range("C16").WrapText = True
    End With
End Sub
