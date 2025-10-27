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
    Dim wsI As Worksheet, wsP As Worksheet, wsV As Worksheet, wsF As Worksheet, wsC As Worksheet
    Set wsI = EnsureSheet("Incidentes")
    Set wsP = EnsureSheet("Personas")
    Set wsV = EnsureSheet("Vehiculos")
    Set wsF = EnsureSheet("Factores")
    Set wsC = EnsureSheet("Catalogos")

    Dim hInc, hPer, hVeh, hFac

    hInc = Array( _
        "id_incidente", "fecha_hora_ocurrencia", "pais", "provincia", "Buenos_Aires", "CABA", _
        "Catamarca", "Chaco", "Chubut", "Córdoba", "Corrientes", "Entre_Ríos", "Formosa", _
        "La_Pampa", "Mendoza", "Misiones", "Neuquen", "Rio_Negro", "Salta", "San_Juan", _
        "San_Luis", "Santa_Cruz", "Santa_Fe", "Santiago", "Tierra_del_Fuego", "Tucuman", _
        "localidad_zona", "coordenadas_geograficas", _
        "lugar_especifico", "uo_incidente", "uo_accidentado", "descripcion_esv", _
        "denuncia_policial", "examen_alcoholemia", "examen_sustancias", "entrevistas_testigos", _
        "accion_inmediata", "consecuencias_seguridad", "fecha_hora_reporte", _
        "cantidad_personas", "cantidad_vehiculos", "clase_evento", "tipo_colision", "nivel_severidad", "clasificacion_esv", _
        "creado_por", "creado_en", "actualizado_por", "actualizado_en")

    hPer = Array( _
        "id_persona", "id_incidente", "nombre_persona", "apellido_persona", "edad_persona", _
        "tipo_persona", "rol_persona", "antiguedad_persona", "tarea_operativa", "turno_operativo", _
        "tipo_danio_persona", "dias_perdidos", "atencion_medica", "in_itinere", _
        "tipo_afectacion", "parte_afectada")

    hVeh = Array( _
        "id_vehiculo", "id_incidente", "tipo_vehiculo", "duenio_vehiculo", "uso_vehiculo", _
        "posee_patente", "numero_patente", "anio_fabricacion_vehiculo", "tarea_vehiculo", "tipo_danio_vehiculo", _
        "cinturon_seguridad", "cabina_cuchetas", "airbags", "gestion_flotas", "token_conductor", _
        "marca_dispositivo", "deteccion_fatiga", "camara_trasera", "limitador_velocidad", "camara_delantera", _
        "camara_punto_ciego", "camara_360", "espejo_punto_ciego", "alarma_marcha_atras", "sistema_frenos", _
        "monitoreo_neumaticos", "proteccion_lateral", "proteccion_trasera", "acondicionador_cabina", "calefaccion_cabina", _
        "manos_libres_cabina", "kit_alcoholemia", "kit_emergencia", "epps_vehiculo", _
        "observaciones_vehiculo", "creado_por", "creado_en", "actualizado_por", "actualizado_en")

    hFac = Array( _
        "id_factor", "id_incidente", "tipo_superficie", "posee_banquina", "tipo_ruta", "densidad_trafico", _
        "condicion_ruta", "iluminacion_ruta", "senalizacion_ruta", "geometria_ruta", "condiciones_climaticas", "rango_temperaturas")

    Dim loI As ListObject, loP As ListObject, loV As ListObject, loF As ListObject
    Set loI = EnsureTable(wsI, "tbIncidente", hInc)
    Set loP = EnsureTable(wsP, "tbPersona", hPer)
    Set loV = EnsureTable(wsV, "tbVehiculo", hVeh)
    Set loF = EnsureTable(wsF, "tbFactores", hFac)

    SetupCatalogos wsC

    ' MsgBox "Estructura creada/actualizada.", vbInformation
End Sub

Public Sub SetupCatalogos(ws As Worksheet)
    ' Comunes
    EnsureCatalog ws, "A", "cat_si_no_na", Array("SI", "NO", "NA")

    ' VehÃ­culos (defaults de ejemplo)
    EnsureCatalog ws, "C", "cat_tipo_vehiculo", Array( _
        "Bicicleta", "Moto", "Ciclomotor", "Autom" & ChrW(243) & "vil", "Pickup", _
        "Cami" & ChrW(243) & "n chasis", "Cami" & ChrW(243) & "n con Cisterna", ChrW(211) & "mnibus")
    EnsureCatalog ws, "E", "cat_duenio_vehiculo", Array("Propio", "Contratista", "Tercero")
    EnsureCatalog ws, "G", "cat_uso_vehiculo", Array("Comercial", "Particular", "Otro", "No se sabe", "NA")

    ' Incidente (placeholders para carga manual)
    EnsureCatalog ws, "I", "cat_pais"
    EnsureCatalog ws, "K", "cat_provincia"
    EnsureCatalog ws, "L", "cat_Buenos_Aires"
    EnsureCatalog ws, "M", "cat_CABA"
    EnsureCatalog ws, "N", "cat_Catamarca"
    EnsureCatalog ws, "O", "cat_Chaco"
    EnsureCatalog ws, "P", "cat_Chubut"
    EnsureCatalog ws, "Q", "cat_Córdoba"
    EnsureCatalog ws, "R", "cat_Corrientes"
    EnsureCatalog ws, "S", "cat_Entre_Ríos"
    EnsureCatalog ws, "T", "cat_Formosa"
    EnsureCatalog ws, "U", "cat_La_Pampa"
    EnsureCatalog ws, "V", "cat_Mendoza"
    EnsureCatalog ws, "W", "cat_Misiones"
    EnsureCatalog ws, "X", "cat_Neuquen"
    EnsureCatalog ws, "Y", "cat_Rio_Negro"
    EnsureCatalog ws, "Z", "cat_Salta"
    EnsureCatalog ws, "AA", "cat_San_Juan"
    EnsureCatalog ws, "AB", "cat_San_Luis"
    EnsureCatalog ws, "AC", "cat_Santa_Cruz"
    EnsureCatalog ws, "AD", "cat_Santa_Fe"
    EnsureCatalog ws, "AE", "cat_Santiago"
    EnsureCatalog ws, "AF", "cat_Tierra_del_Fuego"
    EnsureCatalog ws, "AG", "cat_Tucuman"
    EnsureCatalog ws, "AH", "cat_localidad_zona"
    EnsureCatalog ws, "AI", "cat_uo_incidente"
    EnsureCatalog ws, "AJ", "cat_uo_accidentado"
    EnsureCatalog ws, "AK", "cat_clase_evento"
    EnsureCatalog ws, "AL", "cat_tipo_colision"
    EnsureCatalog ws, "AM", "cat_nivel_severidad"
    EnsureCatalog ws, "AN", "cat_clasificacion_esv"

    ' Personas (placeholders)
    EnsureCatalog ws, "AA", "cat_tipo_persona"
    EnsureCatalog ws, "AC", "cat_rol_persona"
    EnsureCatalog ws, "AE", "cat_antiguedad_persona"
    EnsureCatalog ws, "AG", "cat_tarea_operativa"
    EnsureCatalog ws, "AI", "cat_turno_operativo"
    EnsureCatalog ws, "AJ", "cat_tipo_danio_persona"
    EnsureCatalog ws, "AK", "cat_tipo_afectacion"
    EnsureCatalog ws, "AL", "cat_parte_afectada"

    ' VehÃ­culo adicionales (placeholders)
    EnsureCatalog ws, "AP", "cat_tarea_vehiculo"
    EnsureCatalog ws, "AQ", "cat_tipo_danio_vehiculo"

    ' Factores (placeholders)
    EnsureCatalog ws, "AR", "cat_tipo_superficie"
    EnsureCatalog ws, "AS", "cat_tipo_ruta"
    EnsureCatalog ws, "AY", "cat_densidad_trafico"
    EnsureCatalog ws, "BA", "cat_condicion_ruta"
    EnsureCatalog ws, "BC", "cat_iluminacion_ruta"
    EnsureCatalog ws, "BE", "cat_senalizacion_ruta"
    EnsureCatalog ws, "BG", "cat_geometria_ruta"
    EnsureCatalog ws, "BI", "cat_condiciones_climaticas"
    EnsureCatalog ws, "BK", "cat_rango_temperaturas"
End Sub

Public Sub AddOrUpdateName(nameText As String, refersToRng As Range)
    On Error Resume Next
    Dim nm As name
    Set nm = ThisWorkbook.Names(nameText)
    On Error GoTo 0
    If nm Is Nothing Then
        ThisWorkbook.Names.Add name:=nameText, RefersTo:=refersToRng
    Else
        nm.RefersTo = refersToRng
    End If
End Sub

Public Sub EnsureCatalog(ws As Worksheet, colLetter As String, header As String, Optional defaults As Variant)
    Dim hdrCell As Range, firstData As Range, lastCell As Range, dataRng As Range
    Set hdrCell = ws.Range(colLetter & "1")
    hdrCell.value = header

    Set firstData = ws.Range(colLetter & "2")
    ' Buscar Ãºltimo valor en la columna
    Set lastCell = ws.Cells(ws.Rows.Count, hdrCell.Column).End(xlUp)
    If lastCell.Row < 2 Then
        ' VacÃ­o: sembrar defaults solo si se enviaron
        If Not IsMissing(defaults) Then
            firstData.Resize(UBound(defaults) - LBound(defaults) + 1, 1).value = _
                Application.WorksheetFunction.Transpose(defaults)
        End If
        Set dataRng = ws.Range(firstData, ws.Cells(ws.Rows.Count, hdrCell.Column).End(xlUp))
    Else
        ' Ya hay datos: respetar existentes
        Set dataRng = ws.Range(firstData, lastCell)
    End If

    ' Crear/actualizar nombres en minÃºsculas y mayÃºsculas
    AddOrUpdateName header, dataRng
    AddOrUpdateName UCase(header), dataRng
End Sub
