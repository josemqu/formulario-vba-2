Attribute VB_Name = "clsVehiculoRepo"
Option Explicit

Private Function Headers() As Variant
    Headers = Array( _
        "id_vehiculo", "id_incidente", "tipo_vehiculo", "duenio_vehiculo", "uso_vehiculo", _
        "posee_patente", "numero_patente", "anio_fabricacion_vehiculo", "tarea_vehiculo", "tipo_danio_vehiculo", "estado_vehiculo", _
        "cinturon_seguridad", "cabina_cuchetas", "airbags", "gestion_flotas", "token_conductor", _
        "marca_dispositivo", "deteccion_fatiga", "camara_trasera", "limitador_velocidad", "camara_delantera", _
        "camara_punto_ciego", "camara_360", "espejo_punto_ciego", "alarma_marcha_atras", "sistema_frenos", _
        "monitoreo_neumaticos", "proteccion_lateral", "proteccion_trasera", "acondicionador_cabina", "calefaccion_cabina", _
        "manos_libres_cabina", "kit_alcoholemia", "kit_emergencia", "epps_vehiculo", _
        "observaciones_vehiculo", "creado_por", "creado_en", "actualizado_por", "actualizado_en")
End Function

Private Function WS() As Worksheet
    Set WS = ThisWorkbook.Worksheets("Vehiculos")
End Function

Private Function NextVehiculoId() As String
    Dim lo As ListObject
    Dim maxSeq As Long: maxSeq = 0
    On Error Resume Next
    Set lo = WS.ListObjects("tbVehiculo")
    On Error GoTo 0
    If Not lo Is Nothing Then
        Dim rw As ListRow, s As String, n As Long
        For Each rw In lo.ListRows
            s = CStr(rw.Range.Cells(1, 1).value)
            If LenB(s) > 0 Then
                If Left$(s, 4) = "VEH-" Then
                    n = Val(Mid$(s, 5))
                    If n > maxSeq Then maxSeq = n
                End If
            End If
        Next rw
    End If
    maxSeq = maxSeq + 1
    NextVehiculoId = "VEH-" & Format$(maxSeq, "00000")
End Function

Public Function SaveEntity(ByVal ent As clsVehiculo) As String
    Dim repo As New clsRepository
    Dim d As Object: Set d = ent.ToDict
    If LenB(CStr(ent.id_vehiculo)) = 0 Then
        d("id_vehiculo") = NextVehiculoId()
        d("creado_por") = UserNameOrDefault()
        d("creado_en") = NowIso()
    Else
        d("actualizado_por") = UserNameOrDefault()
        d("actualizado_en") = NowIso()
    End If
    SaveEntity = repo.Save(WS, "tbVehiculo", "id_vehiculo", Headers, d)
End Function

Public Function FindById(id As String) As clsVehiculo
    Dim repo As New clsRepository
    Dim d As Object
    Set d = repo.Find(WS, "tbVehiculo", "id_vehiculo", id)
    If Not d Is Nothing Then
        Dim e As New clsVehiculo
        e.FromDict d
        Set FindById = e
    End If
End Function
