Attribute VB_Name = "modIncidenteForm"
Option Explicit

Private Function CtrlText(frm As Object, ctrlName As String) As String
    On Error Resume Next
    CtrlText = CStr(frm.Controls(ctrlName).value)
    On Error GoTo 0
End Function

Private Function CtrlValue(frm As Object, ctrlName As String) As Variant
    On Error Resume Next
    CtrlValue = frm.Controls(ctrlName).value
    On Error GoTo 0
End Function

Private Sub SetCtrlValue(frm As Object, ctrlName As String, ByVal v As Variant)
    On Error Resume Next
    frm.Controls(ctrlName).value = v
    On Error GoTo 0
End Sub

Public Sub InitIncidenteForm(frm As Object)
    On Error Resume Next
    SetupESVWorkbook
    LoadIncidentCatalogs frm
    If LenB(CtrlText(frm, "txtFechaHoraOcurrencia")) = 0 Then SetCtrlValue frm, "txtFechaHoraOcurrencia", Now
    If LenB(CtrlText(frm, "txtFechaHoraReporte")) = 0 Then SetCtrlValue frm, "txtFechaHoraReporte", Now
    On Error GoTo 0
End Sub

Public Function ReadIncidenteFromForm(frm As Object) As clsIncidente
    Dim e As New clsIncidente
    e.id_incidente = CtrlText(frm, "lblIdIncidente")
    e.fecha_hora_ocurrencia = CtrlValue(frm, "txtFechaHoraOcurrencia")
    e.pais = CtrlText(frm, "cmbPais")
    e.provincia = CtrlText(frm, "cmbProvincia")
    e.localidad_zona = CtrlText(frm, "cmbLocalidad")
    e.coordenadas_geograficas = CtrlText(frm, "txtCoordenadas")
    e.lugar_especifico = CtrlText(frm, "txtLugarEspecifico")
    e.uo_incidente = CtrlText(frm, "cmbUOIncidente")
    e.uo_accidentado = CtrlText(frm, "cmbUOAccidentado")
    e.descripcion_esv = CtrlText(frm, "txtDescripcion")
    e.denuncia_policial = CtrlText(frm, "cmbDenuncia")
    e.examen_alcoholemia = CtrlText(frm, "cmbAlcohol")
    e.examen_sustancias = CtrlText(frm, "cmbSustancias")
    e.entrevistas_testigos = CtrlText(frm, "cmbEntrevistas")
    e.accion_inmediata = CtrlText(frm, "txtAccionInmediata")
    e.consecuencias_seguridad = CtrlText(frm, "cmbConsecuencias")
    e.fecha_hora_reporte = CtrlValue(frm, "txtFechaHoraReporte")
    e.cantidad_personas = CtrlValue(frm, "txtCantidadPersonas")
    e.cantidad_vehiculos = CtrlValue(frm, "txtCantidadVehiculos")
    e.clase_evento = CtrlText(frm, "cmbClaseEvento")
    e.tipo_colision = CtrlText(frm, "cmbTipoColision")
    e.nivel_severidad = CtrlText(frm, "cmbNivelSeveridad")
    e.clasificacion_esv = CtrlText(frm, "cmbClasificacion")
    Set ReadIncidenteFromForm = e
End Function

Public Function SaveIncidenteFromForm(frm As Object) As String
    Dim e As clsIncidente
    Set e = ReadIncidenteFromForm(frm)
    SaveIncidenteFromForm = clsIncidenteRepo.SaveEntity(e)
    SetCtrlValue frm, "lblIdIncidente", SaveIncidenteFromForm
    MsgBox "Incidente guardado: " & SaveIncidenteFromForm, vbInformation
End Function

Public Sub WriteIncidenteToForm(frm As Object, e As clsIncidente)
    On Error Resume Next
    SetCtrlValue frm, "lblIdIncidente", e.id_incidente
    SetCtrlValue frm, "txtFechaHoraOcurrencia", e.fecha_hora_ocurrencia
    SetCtrlValue frm, "cmbPais", e.pais
    SetCtrlValue frm, "cmbProvincia", e.provincia
    SetCtrlValue frm, "cmbLocalidad", e.localidad_zona
    SetCtrlValue frm, "txtCoordenadas", e.coordenadas_geograficas
    SetCtrlValue frm, "txtLugarEspecifico", e.lugar_especifico
    SetCtrlValue frm, "cmbUOIncidente", e.uo_incidente
    SetCtrlValue frm, "cmbUOAccidentado", e.uo_accidentado
    SetCtrlValue frm, "txtDescripcion", e.descripcion_esv
    SetCtrlValue frm, "cmbDenuncia", e.denuncia_policial
    SetCtrlValue frm, "cmbAlcohol", e.examen_alcoholemia
    SetCtrlValue frm, "cmbSustancias", e.examen_sustancias
    SetCtrlValue frm, "cmbEntrevistas", e.entrevistas_testigos
    SetCtrlValue frm, "txtAccionInmediata", e.accion_inmediata
    SetCtrlValue frm, "cmbConsecuencias", e.consecuencias_seguridad
    SetCtrlValue frm, "txtFechaHoraReporte", e.fecha_hora_reporte
    SetCtrlValue frm, "txtCantidadPersonas", e.cantidad_personas
    SetCtrlValue frm, "txtCantidadVehiculos", e.cantidad_vehiculos
    SetCtrlValue frm, "cmbClaseEvento", e.clase_evento
    SetCtrlValue frm, "cmbTipoColision", e.tipo_colision
    SetCtrlValue frm, "cmbNivelSeveridad", e.nivel_severidad
    SetCtrlValue frm, "cmbClasificacion", e.clasificacion_esv
    On Error GoTo 0
End Sub

Public Sub LoadIncidenteById(frm As Object, ByVal id As String)
    Dim e As clsIncidente
    Set e = clsIncidenteRepo.FindById(id)
    If e Is Nothing Then
        MsgBox "No se encontró el incidente: " & id, vbExclamation
        Exit Sub
    End If
    LoadIncidentCatalogs frm
    WriteIncidenteToForm frm, e
End Sub
