Attribute VB_Name = "modSetup"
Option Explicit

Private Function EnsureSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.name = sheetName
    End If
End Function

Public Sub SetupESVWorkbook()
    Dim wsI As Worksheet, wsP As Worksheet, wsV As Worksheet, wsC As Worksheet
    Set wsI = EnsureSheet("Incidentes")
    Set wsP = EnsureSheet("Personas")
    Set wsV = EnsureSheet("Vehiculos")
    Set wsC = EnsureSheet("Catalogos")

    Dim hInc, hPer, hVeh

    hInc = Array( _
        "id_incidente", "fecha_hora_ocurrencia", "pais", "provincia", "localidad_zona", "coordenadas_geograficas", _
        "lugar_especifico", "uo_incidente", "uo_accidentado", "descripcion_esv", _
        "denuncia_policial", "lugar_denuncia_policial", "examen_alcoholemia", "examen_sustancias", "entrevistas_testigos", _
        "accion_inmediata", "consecuencias_seguridad", "fecha_hora_reporte", _
        "cantidad_personas", "cantidad_vehiculos", "clase_evento", "tipo_colision", "nivel_severidad", "clasificacion_esv", _
        "tipo_superficie", "posee_banquina", "tipo_ruta", "velocidad_max_permitida_YPF", "densidad_trafico", "condicion_ruta", "iluminacion_ruta", _
        "senalizacion_ruta", "geometria_ruta", "condiciones_climaticas", "rango_temperaturas", _
        "creado_por", "creado_en", "actualizado_por", "actualizado_en")

    hPer = Array( _
        "id_persona", "id_incidente", "nombre_persona", "apellido_persona", "edad_persona", _
        "tipo_persona", "rol_persona", "antiguedad_persona", "tarea_operativa", "turno_operativo", _
        "tipo_danio_persona", "dias_perdidos", "atencion_medica", "in_itinere", _
        "tipo_afectacion", "parte_afectada", "clase_licencia", "entrenamiento", "aptitud_tarea")

    hVeh = Array( _
        "id_vehiculo", "id_incidente", "tipo_vehiculo", "duenio_vehiculo", "uso_vehiculo", _
        "posee_patente", "numero_patente", "anio_fabricacion_vehiculo", "tarea_vehiculo", "tipo_danio_vehiculo", "estado_vehiculo", _
        "cinturon_seguridad", "cabina_cuchetas", "airbags", "gestion_flotas", "token_conductor", _
        "marca_dispositivo", "deteccion_fatiga", "camara_trasera", "limitador_velocidad", "camara_delantera", _
        "camara_punto_ciego", "camara_360", "espejo_punto_ciego", "alarma_marcha_atras", "sistema_frenos", _
        "monitoreo_neumaticos", "proteccion_lateral", "proteccion_trasera", "acondicionador_cabina", "calefaccion_cabina", _
        "manos_libres_cabina", "kit_alcoholemia", "kit_emergencia", "epps_vehiculo", _
        "observaciones_vehiculo", "creado_por", "creado_en", "actualizado_por", "actualizado_en")

    Dim loI As ListObject, loP As ListObject, loV As ListObject
    Set loI = EnsureTable(wsI, "tbIncidente", hInc)
    Set loP = EnsureTable(wsP, "tbPersona", hPer)
    Set loV = EnsureTable(wsV, "tbVehiculo", hVeh)

    SetupCatalogos wsC

    ' MsgBox "Estructura creada/actualizada.", vbInformation
End Sub

Public Sub SetupCatalogos(WS As Worksheet)
    EnsureCatalog WS, "A", "cat_si_no_na"
    EnsureCatalog WS, "B", "cat_tipo_vehiculo"
    EnsureCatalog WS, "C", "cat_duenio_vehiculo"
    EnsureCatalog WS, "D", "cat_uso_vehiculo"
    EnsureCatalog WS, "E", "cat_pais"
    EnsureCatalog WS, "F", "cat_provincia"
    EnsureCatalog WS, "G", "cat_Buenos_Aires"
    EnsureCatalog WS, "H", "cat_CABA"
    EnsureCatalog WS, "I", "cat_Catamarca"
    EnsureCatalog WS, "J", "cat_Chaco"
    EnsureCatalog WS, "K", "cat_Chubut"
    EnsureCatalog WS, "L", "cat_Cordoba"
    EnsureCatalog WS, "M", "cat_Corrientes"
    EnsureCatalog WS, "N", "cat_Entre_Rios"
    EnsureCatalog WS, "O", "cat_Formosa"
    EnsureCatalog WS, "P", "cat_La_Pampa"
    EnsureCatalog WS, "Q", "cat_Mendoza"
    EnsureCatalog WS, "R", "cat_Misiones"
    EnsureCatalog WS, "S", "cat_Neuquen"
    EnsureCatalog WS, "T", "cat_Rio_Negro"
    EnsureCatalog WS, "U", "cat_Salta"
    EnsureCatalog WS, "V", "cat_San_Juan"
    EnsureCatalog WS, "W", "cat_San_Luis"
    EnsureCatalog WS, "X", "cat_Santa_Cruz"
    EnsureCatalog WS, "Y", "cat_Santa_Fe"
    EnsureCatalog WS, "Z", "cat_Santiago"
    EnsureCatalog WS, "AA", "cat_Tierra_del_Fuego"
    EnsureCatalog WS, "AB", "cat_Tucuman"
    EnsureCatalog WS, "AC", "cat_localidad_zona"
    EnsureCatalog WS, "AD", "cat_uo_incidente"
    EnsureCatalog WS, "AE", "cat_uo_accidentado"
    EnsureCatalog WS, "AF", "cat_clase_evento"
    EnsureCatalog WS, "AG", "cat_tipo_colision"
    EnsureCatalog WS, "AH", "cat_nivel_severidad"
    EnsureCatalog WS, "AI", "cat_clasificacion_esv"
    EnsureCatalog WS, "AJ", "cat_tipo_persona"
    EnsureCatalog WS, "AK", "cat_rol_persona"
    EnsureCatalog WS, "AL", "cat_antiguedad_persona"
    EnsureCatalog WS, "AM", "cat_tarea_operativa"
    EnsureCatalog WS, "AN", "cat_turno_operativo"
    EnsureCatalog WS, "AO", "cat_tipo_danio_persona"
    EnsureCatalog WS, "AP", "cat_tipo_afectacion"
    EnsureCatalog WS, "AQ", "cat_parte_afectada"
    EnsureCatalog WS, "AR", "cat_tarea_vehiculo"
    EnsureCatalog WS, "AS", "cat_tipo_danio_vehiculo"
    EnsureCatalog WS, "AT", "cat_tipo_superficie"
    EnsureCatalog WS, "AU", "cat_tipo_ruta"
    EnsureCatalog WS, "AV", "cat_densidad_trafico"
    EnsureCatalog WS, "AW", "cat_condicion_ruta"
    EnsureCatalog WS, "AX", "cat_iluminacion_ruta"
    EnsureCatalog WS, "AY", "cat_senalizacion_ruta"
    EnsureCatalog WS, "AZ", "cat_geometria_ruta"
    EnsureCatalog WS, "BA", "cat_condiciones_climaticas"
    EnsureCatalog WS, "BB", "cat_rango_temperaturas"
    EnsureCatalog WS, "BC", "cat_clase_licencia"
    EnsureCatalog WS, "BD", "cat_entrenamiento"
    EnsureCatalog WS, "BE", "cat_aptitud_tarea"
    EnsureCatalog WS, "BF", "cat_epps_vehiculo"

End Sub

Public Sub AddOrUpdateName(nameText As String, refersToRng As Range)
    Dim wb As Workbook
    Set wb = refersToRng.Parent.Parent ' Workbook propietario del rango
    On Error Resume Next
    Dim nm As Name
    Set nm = wb.Names(nameText)
    On Error GoTo 0
    If nm Is Nothing Then
        wb.Names.Add Name:=nameText, RefersTo:=refersToRng
    Else
        nm.RefersTo = refersToRng
    End If
End Sub

Public Sub EnsureCatalog(WS As Worksheet, colLetter As String, header As String, Optional defaults As Variant)
    Dim hdrCell As Range, firstData As Range, lastCell As Range, dataRng As Range
    Set hdrCell = WS.Range(colLetter & "1")
    hdrCell.value = header

    Set firstData = WS.Range(colLetter & "2")
    ' Buscar ultimo valor en la columna
    Set lastCell = WS.Cells(WS.Rows.Count, hdrCell.Column).End(xlUp)
    If lastCell.Row < 2 Then
        ' Vacio: sembrar defaults solo si se enviaron
        If Not IsMissing(defaults) Then
            firstData.Resize(UBound(defaults) - LBound(defaults) + 1, 1).value = _
                Application.WorksheetFunction.Transpose(defaults)
        End If
        Set dataRng = WS.Range(firstData, WS.Cells(WS.Rows.Count, hdrCell.Column).End(xlUp))
    Else
        ' Ya hay datos: respetar existentes
        Set dataRng = WS.Range(firstData, lastCell)
    End If

    ' Crear/actualizar nombres en minusculas y mayusculas
    AddOrUpdateName header, dataRng
    AddOrUpdateName UCase(header), dataRng
End Sub
