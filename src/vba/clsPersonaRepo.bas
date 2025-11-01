Attribute VB_Name = "clsPersonaRepo"
Option Explicit

Private Function Headers() As Variant
    Headers = Array( _
        "id_persona", "id_incidente", "nombre_persona", "apellido_persona", "edad_persona", _
        "tipo_persona", "rol_persona", "antiguedad_persona", "tarea_operativa", "turno_operativo", _
        "tipo_danio_persona", "dias_perdidos", "atencion_medica", "in_itinere", _
        "tipo_afectacion", "parte_afectada", "clase_licencia", "entrenamiento", "aptitud_tarea")
End Function

Private Function WS() As Worksheet
    Set WS = ThisWorkbook.Worksheets("Personas")
End Function

Private Function NextPersonaId() As String
    Dim lo As ListObject
    Dim maxSeq As Long: maxSeq = 0
    On Error Resume Next
    Set lo = WS.ListObjects("tbPersona")
    On Error GoTo 0
    If Not lo Is Nothing Then
        Dim rw As ListRow, s As String, n As Long
        For Each rw In lo.ListRows
            s = CStr(rw.Range.Cells(1, 1).value)
            If LenB(s) > 0 Then
                If Left$(s, 4) = "PER-" Then
                    n = Val(Mid$(s, 5))
                    If n > maxSeq Then maxSeq = n
                End If
            End If
        Next rw
    End If
    maxSeq = maxSeq + 1
    NextPersonaId = "PER-" & Format$(maxSeq, "00000")
End Function

Public Function SaveEntity(ByVal ent As clsPersona) As String
    Dim repo As New clsRepository
    Dim d As Object: Set d = ent.ToDict
    If LenB(CStr(ent.id_persona)) = 0 Then
        d("id_persona") = NextPersonaId()
    End If
    SaveEntity = repo.Save(WS, "tbPersona", "id_persona", Headers, d)
End Function

Public Function FindById(id As String) As clsPersona
    Dim repo As New clsRepository
    Dim d As Object
    Set d = repo.Find(WS, "tbPersona", "id_persona", id)
    If Not d Is Nothing Then
        Dim e As New clsPersona
        e.FromDict d
        Set FindById = e
    End If
End Function
