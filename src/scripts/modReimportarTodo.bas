Attribute VB_Name = "modSheetIncidente"
Option Explicit

Private Function EnsureFormSheet() As Worksheet
    Dim WS As Worksheet
    On Error Resume Next
    Set WS = ThisWorkbook.Worksheets("Formulario")
    On Error GoTo 0
    If WS Is Nothing Then
        Set WS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        WS.name = "Formulario"
    End If
    Set EnsureFormSheet = WS
End Function

Private Sub LayoutForm(WS As Worksheet)
    WS.Range("B2").value = "ID incidente"
    WS.Range("B3").value = "Fecha/hora ocurrencia"
    WS.Range("B4").value = "Pais"
    WS.Range("B5").value = "Provincia"
    WS.Range("B6").value = "Localidad/Zona"
    WS.Range("B7").value = "Coordenadas"
    WS.Range("B8").value = "Lugar especifico"
    WS.Range("B9").value = "UO incidente"
    WS.Range("B10").value = "UO accidentado"
    WS.Range("B11").value = "Descripcion"
    WS.Range("B12").value = "Denuncia policial"
    WS.Range("B13").value = "Examen alcoholemia"
    WS.Range("B14").value = "Examen sustancias"
    WS.Range("B15").value = "Entrevistas testigos"
    WS.Range("B16").value = "Accion inmediata"
    WS.Range("B17").value = "Consecuencias seguridad"
    WS.Range("B18").value = "Fecha/hora reporte"
    WS.Range("B19").value = "Cantidad personas"
    WS.Range("B20").value = "Cantidad vehiculos"
    WS.Range("B21").value = "Clase evento"
    WS.Range("B22").value = "Tipo colision"
    WS.Range("B23").value = "Nivel severidad"
    WS.Range("B24").value = "Clasificacion ESV"
    If LenB(CStr(WS.Range("C3").value)) = 0 Then WS.Range("C3").value = Now
    If LenB(CStr(WS.Range("C18").value)) = 0 Then WS.Range("C18").value = Now
    WS.Columns("B:B").ColumnWidth = 26
    WS.Columns("C:C").ColumnWidth = 50
    WS.Range("B2:B24").WrapText = True
    WS.Range("C3,C18").NumberFormat = "dd/mm/yyyy hh:mm"
    WS.Range("C19:C20").NumberFormat = "0"
End Sub

Private Sub ApplyValidations(WS As Worksheet)
    Dim r As Range
    On Error Resume Next
    Set r = WS.Range("C4"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_PAIS"
    Set r = WS.Range("C5"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_PROVINCIA"
    Set r = WS.Range("C6"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_LOCALIDAD_ZONA"
    Set r = WS.Range("C9"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_UO_INCIDENTE"
    Set r = WS.Range("C10"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_UO_ACCIDENTADO"
    Set r = WS.Range("C12"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = WS.Range("C13"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = WS.Range("C14"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = WS.Range("C15"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = WS.Range("C17"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_SI_NO_NA"
    Set r = WS.Range("C21"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_CLASE_EVENTO"
    Set r = WS.Range("C22"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_TIPO_COLISION"
    Set r = WS.Range("C23"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_NIVEL_SEVERIDAD"
    Set r = WS.Range("C24"): r.Validation.Delete: r.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=CAT_CLASIFICACION_ESV"
    On Error GoTo 0
End Sub

Private Sub EnsureGuardarButton(WS As Worksheet)
    Dim shp As Shape
    On Error Resume Next
    Set shp = WS.Shapes("btnGuardarIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = WS.Shapes.AddShape(msoShapeRoundedRectangle, WS.Range("B26").Left, WS.Range("B26").Top, 160, 32)
        shp.name = "btnGuardarIncidente"
        shp.TextFrame.Characters.Text = "Guardar incidente"
        shp.OnAction = "GuardarIncidenteDesdeHoja"
    Else
        shp.OnAction = "GuardarIncidenteDesdeHoja"
    End If
    On Error Resume Next
    Set shp = WS.Shapes("btnNuevoIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = WS.Shapes.AddShape(msoShapeRoundedRectangle, WS.Range("D26").Left, WS.Range("D26").Top, 120, 32)
        shp.name = "btnNuevoIncidente"
        shp.TextFrame.Characters.Text = "Nuevo"
        shp.OnAction = "NuevoIncidenteEnHoja"
    Else
        shp.OnAction = "NuevoIncidenteEnHoja"
    End If
    On Error Resume Next
    Set shp = WS.Shapes("btnEliminarIncidente")
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = WS.Shapes.AddShape(msoShapeRoundedRectangle, WS.Range("F26").Left, WS.Range("F26").Top, 120, 32)
        shp.name = "btnEliminarIncidente"
        shp.TextFrame.Characters.Text = "Eliminar"
        shp.OnAction = "EliminarIncidenteDesdeHoja"
    Else
        shp.OnAction = "EliminarIncidenteDesdeHoja"
    End If
End Sub

Public Sub AbrirFormularioIncidenteEnHoja()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    LayoutForm WS
    ApplyValidations WS
    EnsureGuardarButton WS
    EstilizarFormularioIncidente
    WS.Activate
End Sub

Private Function ReadIncidenteFromSheet(WS As Worksheet) As clsIncidente
    Dim e As New clsIncidente
    e.id_incidente = CStr(WS.Range("C2").value)
    e.fecha_hora_ocurrencia = WS.Range("C3").value
    e.pais = CStr(WS.Range("C4").value)
    e.provincia = CStr(WS.Range("C5").value)
    e.localidad_zona = CStr(WS.Range("C6").value)
    e.coordenadas_geograficas = CStr(WS.Range("C7").value)
    e.lugar_especifico = CStr(WS.Range("C8").value)
    e.uo_incidente = CStr(WS.Range("C9").value)
    e.uo_accidentado = CStr(WS.Range("C10").value)
    e.descripcion_esv = CStr(WS.Range("C11").value)
    e.denuncia_policial = CStr(WS.Range("C12").value)
    e.examen_alcoholemia = CStr(WS.Range("C13").value)
    e.examen_sustancias = CStr(WS.Range("C14").value)
    e.entrevistas_testigos = CStr(WS.Range("C15").value)
    e.accion_inmediata = CStr(WS.Range("C16").value)
    e.consecuencias_seguridad = CStr(WS.Range("C17").value)
    e.fecha_hora_reporte = WS.Range("C18").value
    e.cantidad_personas = WS.Range("C19").value
    e.cantidad_vehiculos = WS.Range("C20").value
    e.clase_evento = CStr(WS.Range("C21").value)
    e.tipo_colision = CStr(WS.Range("C22").value)
    e.nivel_severidad = CStr(WS.Range("C23").value)
    e.clasificacion_esv = CStr(WS.Range("C24").value)
    Set ReadIncidenteFromSheet = e
End Function

Private Sub ClearForm(WS As Worksheet)
    WS.Range("C2:C24").ClearContents
    WS.Range("C3").value = Now
    WS.Range("C18").value = Now
End Sub

Private Function ValidateForm(WS As Worksheet, ByRef messages As Collection) As Boolean
    Dim ok As Boolean: ok = True
    Set messages = New Collection
    If LenB(CStr(WS.Range("C3").value)) = 0 Then ok = False: messages.Add ("Fecha/hora ocurrencia es requerida.")
    If LenB(CStr(WS.Range("C4").value)) = 0 Then ok = False: messages.Add ("País es requerido.")
    If LenB(CStr(WS.Range("C21").value)) = 0 Then ok = False: messages.Add ("Clase de evento es requerida.")
    If LenB(CStr(WS.Range("C19").value)) > 0 Then If Not IsNumeric(WS.Range("C19").value) Then ok = False: messages.Add ("Cantidad personas debe ser numérico.")
    If LenB(CStr(WS.Range("C20").value)) > 0 Then If Not IsNumeric(WS.Range("C20").value) Then ok = False: messages.Add ("Cantidad vehículos debe ser numérico.")
    ValidateForm = ok
End Function

Public Sub GuardarIncidenteDesdeHoja()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    Dim msgs As Collection
    If Not ValidateForm(WS, msgs) Then
        Dim t As String: t = "No se puede guardar. Corrige los siguientes puntos:" & vbCrLf
        Dim it As Variant
        For Each it In msgs
            t = t & "- " & CStr(it) & vbCrLf
        Next it
        MsgBox t, vbExclamation
        Exit Sub
    End If
    Dim e As clsIncidente
    Set e = ReadIncidenteFromSheet(WS)
    Dim id As String
    id = clsIncidenteRepo.SaveEntity(e)
    WS.Range("C2").value = id
    MsgBox "Incidente guardado: " & id, vbInformation
End Sub

Public Sub NuevoIncidenteEnHoja()
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    ClearForm WS
End Sub

Public Sub EliminarIncidenteDesdeHoja()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    Dim id As String: id = CStr(WS.Range("C2").value)
    If LenB(id) = 0 Then
        MsgBox "No hay ID en C2 para eliminar.", vbExclamation
        Exit Sub
    End If
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Eliminar el incidente " & id & " de forma permanente?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar eliminación")
    If resp <> vbYes Then Exit Sub
    clsIncidenteRepo.DeleteById id
    ClearForm WS
    MsgBox "Incidente eliminado.", vbInformation
End Sub

Public Sub LoadIncidenteEnHojaDesdeIdActual()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    Dim id As String: id = CStr(WS.Range("C2").value)
    If LenB(id) = 0 Then Exit Sub
    Dim e As clsIncidente
    Set e = clsIncidenteRepo.FindById(id)
    If e Is Nothing Then Exit Sub
    WS.Range("C3").value = e.fecha_hora_ocurrencia
    WS.Range("C4").value = e.pais
    WS.Range("C5").value = e.provincia
    WS.Range("C6").value = e.localidad_zona
    WS.Range("C7").value = e.coordenadas_geograficas
    WS.Range("C8").value = e.lugar_especifico
    WS.Range("C9").value = e.uo_incidente
    WS.Range("C10").value = e.uo_accidentado
    WS.Range("C11").value = e.descripcion_esv
    WS.Range("C12").value = e.denuncia_policial
    WS.Range("C13").value = e.examen_alcoholemia
    WS.Range("C14").value = e.examen_sustancias
    WS.Range("C15").value = e.entrevistas_testigos
    WS.Range("C16").value = e.accion_inmediata
    WS.Range("C17").value = e.consecuencias_seguridad
    WS.Range("C18").value = e.fecha_hora_reporte
    WS.Range("C19").value = e.cantidad_personas
    WS.Range("C20").value = e.cantidad_vehiculos
    WS.Range("C21").value = e.clase_evento
    WS.Range("C22").value = e.tipo_colision
    WS.Range("C23").value = e.nivel_severidad
    WS.Range("C24").value = e.clasificacion_esv
End Sub

Public Sub EstilizarFormularioIncidente()
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    WS.Cells.Font.name = "Calibri"
    WS.Cells.Font.Size = 11
    WS.Range("B2:B24").Font.Bold = True
    WS.Range("B2:B24").Interior.Color = RGB(245, 245, 245)
    WS.Range("C2:C24").Interior.Color = RGB(255, 255, 255)
    With WS.Range("B2:C24").Borders
        .LineStyle = xlContinuous
        .Color = RGB(220, 220, 220)
        .Weight = xlThin
    End With
    WS.Range("B2:C24").Borders(xlInsideHorizontal).Color = RGB(235, 235, 235)
    WS.Range("B2:C24").Borders(xlInsideVertical).Color = RGB(235, 235, 235)
    WS.Rows("2:24").RowHeight = 20
    WS.Columns("B:C").HorizontalAlignment = xlLeft
    WS.Columns("C:C").HorizontalAlignment = xlLeft
    WS.Columns("C:C").VerticalAlignment = xlCenter
    WS.Range("B1:C1").Merge
    WS.Range("B1").value = "Registro de Incidente"
    WS.Range("B1").Font.Size = 16
    WS.Range("B1").Font.Bold = True
    WS.Range("B1").Font.Color = RGB(32, 32, 32)
    WS.Range("B1").Interior.Color = RGB(255, 255, 255)
    WS.Range("B1").EntireRow.RowHeight = 28
    On Error Resume Next
    Dim shp As Shape
    Set shp = WS.Shapes("btnGuardarIncidente")
    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = RGB(0, 120, 215)
        shp.Line.ForeColor.RGB = RGB(0, 84, 153)
        shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        shp.TextFrame.Characters.Font.Bold = True
    End If
    Set shp = WS.Shapes("btnNuevoIncidente")
    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = RGB(0, 153, 51)
        shp.Line.ForeColor.RGB = RGB(0, 102, 34)
        shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        shp.TextFrame.Characters.Font.Bold = True
    End If
    Set shp = WS.Shapes("btnEliminarIncidente")
    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = RGB(220, 53, 69)
        shp.Line.ForeColor.RGB = RGB(176, 42, 55)
        shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        shp.TextFrame.Characters.Font.Bold = True
    End If
    Set shp = Nothing
    WS.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

