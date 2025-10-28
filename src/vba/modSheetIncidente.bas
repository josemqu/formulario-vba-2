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

Private Sub EnsureActionButton(WS As Worksheet, ByVal btnName As String, ByVal cellAddr As String, ByVal w As Single, ByVal h As Single, ByVal caption As String, ByVal macroName As String)
    Dim shp As Shape
    On Error Resume Next
    Set shp = WS.Shapes(btnName)
    On Error GoTo 0
    ' Si existe pero no es AutoShape, eliminar y recrear
    If Not shp Is Nothing Then
        If shp.Type <> msoAutoShape Then
            shp.Delete
            Set shp = Nothing
        End If
    End If
    If shp Is Nothing Then
        Set shp = WS.Shapes.AddShape(msoShapeRoundedRectangle, WS.Range(cellAddr).Left, WS.Range(cellAddr).Top, w, h)
        shp.name = btnName
    Else
        shp.Left = WS.Range(cellAddr).Left
        shp.Top = WS.Range(cellAddr).Top
        shp.Width = w
        shp.Height = h
    End If
    With shp
        .TextFrame.Characters.Text = caption
        .OnAction = "modSheetIncidente." & macroName
        .ZOrder msoBringToFront
        .Placement = xlMoveAndSize
        .LockAspectRatio = msoFalse
        On Error Resume Next
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        On Error GoTo 0
    End With
End Sub

Private Sub EnsureGuardarButton(WS As Worksheet)
    EnsureActionButton WS, "btnGuardarIncidente", "E2", 120, 32, "Guardar incidente", "GuardarIncidenteDesdeHoja"
    EnsureActionButton WS, "btnNuevoIncidente", "E4", 120, 32, "Nuevo", "NuevoIncidenteEnHoja"
    EnsureActionButton WS, "btnEliminarIncidente", "E6", 120, 32, "Eliminar", "EliminarIncidenteDesdeHoja"
End Sub

' ==== Sección Personas ====
Private Function PersonasStartRow() As Long
    PersonasStartRow = 26
End Function

Private Sub LayoutPersonasSection(WS As Worksheet)
    Dim r0 As Long: r0 = PersonasStartRow()
    WS.Range("B" & r0 - 1).value = "Personas"
    Dim labels As Variant
    labels = Array( _
        "id_persona", "id_incidente", "nombre_persona", "apellido_persona", "edad_persona", _
        "tipo_persona", "rol_persona", "antiguedad_persona", "tarea_operativa", "turno_operativo", _
        "tipo_danio_persona", "dias_perdidos", "atencion_medica", "in_itinere", _
        "tipo_afectacion", "parte_afectada")
    Dim i As Long
    For i = 0 To UBound(labels)
        WS.Range("B" & (r0 + i)).value = labels(i)
    Next i
    ' Ajustes visuales basicos
    WS.Range("B" & (r0 - 1) & ":B" & (r0 + UBound(labels))).Font.Bold = True
    WS.Range("B" & r0 & ":B" & (r0 + UBound(labels))).Interior.Color = RGB(245, 245, 245)
End Sub

Private Sub EnsureAddPersonaButton(WS As Worksheet)
    EnsureActionButton WS, "btnAddPersona", "E" & (PersonasStartRow() - 1), 140, 26, "Agregar Persona", "AgregarColumnaPersona"
End Sub

Public Sub AgregarColumnaPersona()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    Dim r0 As Long: r0 = PersonasStartRow()
    Dim col As Long: col = NextEntityColumn(WS, r0)
    WS.Cells(r0 - 1, col).value = "Persona " & (col - 2) ' C=3 -> 1
    ' Número para edad y días
    WS.Cells(r0 + 4, col).NumberFormat = "0"
    WS.Cells(r0 + 11, col).NumberFormat = "0"
End Sub

Private Function NextEntityColumn(WS As Worksheet, headerRow As Long) As Long
    ' Empieza en columna C (3) y avanza hasta la primera vacia en la fila de encabezado
    Dim c As Long: c = 3
    Do While LenB(CStr(WS.Cells(headerRow - 1, c).value)) > 0 Or LenB(CStr(WS.Cells(headerRow, c).value)) > 0
        c = c + 1
    Loop
    NextEntityColumn = c
End Function

Private Function ReadAndSavePersonas(WS As Worksheet, ByVal idInc As String) As Long
    On Error GoTo fin
    Dim r0 As Long: r0 = PersonasStartRow()
    Dim col As Long: col = 3 ' C
    Dim countSaved As Long: countSaved = 0
    Do While LenB(CStr(WS.Cells(r0 - 1, col).value)) > 0 Or LenB(CStr(WS.Cells(r0 + 2, col).value)) > 0
        Dim anyValue As Boolean: anyValue = False
        Dim e As New clsPersona
        e.id_persona = CStr(WS.Cells(r0 + 0, col).value)
        e.id_incidente = idInc
        e.nombre_persona = CStr(WS.Cells(r0 + 2, col).value): If LenB(e.nombre_persona) > 0 Then anyValue = True
        e.apellido_persona = CStr(WS.Cells(r0 + 3, col).value)
        e.edad_persona = WS.Cells(r0 + 4, col).value
        e.tipo_persona = CStr(WS.Cells(r0 + 5, col).value)
        e.rol_persona = CStr(WS.Cells(r0 + 6, col).value)
        e.antiguedad_persona = CStr(WS.Cells(r0 + 7, col).value)
        e.tarea_operativa = CStr(WS.Cells(r0 + 8, col).value)
        e.turno_operativo = CStr(WS.Cells(r0 + 9, col).value)
        e.tipo_danio_persona = CStr(WS.Cells(r0 + 10, col).value)
        e.dias_perdidos = WS.Cells(r0 + 11, col).value
        e.atencion_medica = CStr(WS.Cells(r0 + 12, col).value)
        e.in_itinere = CStr(WS.Cells(r0 + 13, col).value)
        e.tipo_afectacion = CStr(WS.Cells(r0 + 14, col).value)
        e.parte_afectada = CStr(WS.Cells(r0 + 15, col).value)
        If anyValue Then
            Dim newId As String
            newId = clsPersonaRepo.SaveEntity(e)
            WS.Cells(r0 + 0, col).value = newId
            WS.Cells(r0 + 1, col).value = idInc
            countSaved = countSaved + 1
        End If
        col = col + 1
    Loop
    ReadAndSavePersonas = countSaved
    Exit Function
fin:
    ReadAndSavePersonas = -1
End Function

' ==== Sección Vehículos ====
Private Function VehiculosStartRow() As Long
    VehiculosStartRow = PersonasStartRow() + 18
End Function

Private Sub LayoutVehiculosSection(WS As Worksheet)
    Dim r0 As Long: r0 = VehiculosStartRow()
    WS.Range("B" & r0 - 1).value = "Vehículos"
    Dim labels As Variant
    labels = Array( _
        "id_vehiculo", "id_incidente", "tipo_vehiculo", "duenio_vehiculo", "uso_vehiculo", _
        "posee_patente", "numero_patente", "anio_fabricacion_vehiculo", "tarea_vehiculo", "tipo_danio_vehiculo", _
        "cinturon_seguridad", "cabina_cuchetas", "airbags", "gestion_flotas", "token_conductor", _
        "marca_dispositivo", "deteccion_fatiga", "camara_trasera", "limitador_velocidad", "camara_delantera", _
        "camara_punto_ciego", "camara_360", "espejo_punto_ciego", "alarma_marcha_atras", "sistema_frenos", _
        "monitoreo_neumaticos", "proteccion_lateral", "proteccion_trasera", "acondicionador_cabina", "calefaccion_cabina", _
        "manos_libres_cabina", "kit_alcoholemia", "kit_emergencia", "epps_vehiculo", _
        "observaciones_vehiculo")
    Dim i As Long
    For i = 0 To UBound(labels)
        WS.Range("B" & (r0 + i)).value = labels(i)
    Next i
    WS.Range("B" & (r0 - 1) & ":B" & (r0 + UBound(labels))).Font.Bold = True
    WS.Range("B" & r0 & ":B" & (r0 + UBound(labels))).Interior.Color = RGB(245, 245, 245)
End Sub

Private Sub EnsureAddVehiculoButton(WS As Worksheet)
    EnsureActionButton WS, "btnAddVehiculo", "G" & (VehiculosStartRow() - 1), 160, 26, "Agregar Vehículo", "AgregarColumnaVehiculo"
End Sub

Public Sub AgregarColumnaVehiculo()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    Dim r0 As Long: r0 = VehiculosStartRow()
    Dim col As Long: col = NextEntityColumn(WS, r0)
    WS.Cells(r0 - 1, col).value = "Vehículo " & (col - 2)
End Sub

Private Function ReadAndSaveVehiculos(WS As Worksheet, ByVal idInc As String) As Long
    On Error GoTo fin
    Dim r0 As Long: r0 = VehiculosStartRow()
    Dim col As Long: col = 3 ' C
    Dim countSaved As Long: countSaved = 0
    Do While LenB(CStr(WS.Cells(r0 - 1, col).value)) > 0 Or LenB(CStr(WS.Cells(r0 + 2, col).value)) > 0
        Dim anyValue As Boolean: anyValue = False
        Dim v As New clsVehiculo
        v.id_vehiculo = CStr(WS.Cells(r0 + 0, col).value)
        v.id_incidente = idInc
        v.tipo_vehiculo = CStr(WS.Cells(r0 + 2, col).value): If LenB(v.tipo_vehiculo) > 0 Then anyValue = True
        v.duenio_vehiculo = CStr(WS.Cells(r0 + 3, col).value)
        v.uso_vehiculo = CStr(WS.Cells(r0 + 4, col).value)
        v.posee_patente = CStr(WS.Cells(r0 + 5, col).value)
        v.numero_patente = CStr(WS.Cells(r0 + 6, col).value)
        v.anio_fabricacion_vehiculo = CStr(WS.Cells(r0 + 7, col).value)
        v.tarea_vehiculo = CStr(WS.Cells(r0 + 8, col).value)
        v.tipo_danio_vehiculo = CStr(WS.Cells(r0 + 9, col).value)
        v.cinturon_seguridad = CStr(WS.Cells(r0 + 10, col).value)
        v.cabina_cuchetas = CStr(WS.Cells(r0 + 11, col).value)
        v.airbags = CStr(WS.Cells(r0 + 12, col).value)
        v.gestion_flotas = CStr(WS.Cells(r0 + 13, col).value)
        v.token_conductor = CStr(WS.Cells(r0 + 14, col).value)
        v.marca_dispositivo = CStr(WS.Cells(r0 + 15, col).value)
        v.deteccion_fatiga = CStr(WS.Cells(r0 + 16, col).value)
        v.camara_trasera = CStr(WS.Cells(r0 + 17, col).value)
        v.limitador_velocidad = CStr(WS.Cells(r0 + 18, col).value)
        v.camara_delantera = CStr(WS.Cells(r0 + 19, col).value)
        v.camara_punto_ciego = CStr(WS.Cells(r0 + 20, col).value)
        v.camara_360 = CStr(WS.Cells(r0 + 21, col).value)
        v.espejo_punto_ciego = CStr(WS.Cells(r0 + 22, col).value)
        v.alarma_marcha_atras = CStr(WS.Cells(r0 + 23, col).value)
        v.sistema_frenos = CStr(WS.Cells(r0 + 24, col).value)
        v.monitoreo_neumaticos = CStr(WS.Cells(r0 + 25, col).value)
        v.proteccion_lateral = CStr(WS.Cells(r0 + 26, col).value)
        v.proteccion_trasera = CStr(WS.Cells(r0 + 27, col).value)
        v.acondicionador_cabina = CStr(WS.Cells(r0 + 28, col).value)
        v.calefaccion_cabina = CStr(WS.Cells(r0 + 29, col).value)
        v.manos_libres_cabina = CStr(WS.Cells(r0 + 30, col).value)
        v.kit_alcoholemia = CStr(WS.Cells(r0 + 31, col).value)
        v.kit_emergencia = CStr(WS.Cells(r0 + 32, col).value)
        v.epps_vehiculo = CStr(WS.Cells(r0 + 33, col).value)
        v.observaciones_vehiculo = CStr(WS.Cells(r0 + 34, col).value)
        If anyValue Then
            Dim newId As String
            newId = clsVehiculoRepo.SaveEntity(v)
            WS.Cells(r0 + 0, col).value = newId
            WS.Cells(r0 + 1, col).value = idInc
            countSaved = countSaved + 1
        End If
        col = col + 1
    Loop
    ReadAndSaveVehiculos = countSaved
    Exit Function
fin:
    ReadAndSaveVehiculos = -1
End Function

Public Sub AbrirFormularioIncidenteEnHoja()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    LayoutForm WS
    ApplyValidations WS
    EnsureGuardarButton WS
    EstilizarFormularioIncidente
    LayoutPersonasSection WS
    LayoutVehiculosSection WS
    EnsureAddPersonaButton WS
    EnsureAddVehiculoButton WS
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
    Dim cantP As Long, cantV As Long
    cantP = ReadAndSavePersonas(WS, id)
    cantV = ReadAndSaveVehiculos(WS, id)
    If cantP >= 0 Then WS.Range("C19").value = cantP
    If cantV >= 0 Then WS.Range("C20").value = cantV
    MsgBox "Incidente guardado: " & id & vbCrLf & _
            "Personas guardadas: " & cantP & vbCrLf & _
            "Vehículos guardados: " & cantV, vbInformation
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

