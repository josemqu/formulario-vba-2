Attribute VB_Name = "modSheetIncidente"
Option Explicit

Private Function EnsureFormSheet() As Worksheet
    Dim WS As Worksheet
    On Error Resume Next
    Set WS = ThisWorkbook.Worksheets("Form")
    On Error GoTo 0
    If WS Is Nothing Then
        Set WS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        WS.name = "Form"
    End If
    Set EnsureFormSheet = WS
End Function

' Layout & formatting are managed manually in "Form". No auto-formatting here.

' Validations are maintained manually on the "Form" sheet.

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


' ==== Seccion Personas ====
Private Function PersonasStartRow() As Long
    PersonasStartRow = 26
End Function



Public Sub AgregarColumnaPersona()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    Dim firstCol As Long, lastCol As Long
    firstCol = WS.Columns("K").Column ' 11
    lastCol = WS.Columns("T").Column ' 20
    Dim i As Long, visibleCount As Long
    visibleCount = 0
    ' Contar cuantas columnas K:T estan visibles
    For i = firstCol To lastCol
        If WS.Columns(i).Hidden = False Then visibleCount = visibleCount + 1
    Next i
    ' Si ninguna visible, mostrar la primera (K)
    If visibleCount = 0 Then
        WS.Columns(firstCol).Hidden = False
        Exit Sub
    End If
    ' Buscar la proxima columna oculta para mostrar
    For i = firstCol To lastCol
        If WS.Columns(i).Hidden = True Then
            WS.Columns(i).Hidden = False
            Exit Sub
        End If
    Next i
    ' Si llegamos aqui, ya estan visibles las 10 columnas
    MsgBox "Ya alcanzaste el máximo de 10 columnas de Personas (K:T).", vbInformation
End Sub

Public Sub EliminarColumnaPersona()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    Dim firstCol As Long, lastCol As Long
    firstCol = WS.Columns("K").Column
    lastCol = WS.Columns("T").Column
    Dim i As Long, visibleCount As Long, lastVisible As Long
    visibleCount = 0: lastVisible = -1
    For i = firstCol To lastCol
        If WS.Columns(i).Hidden = False Then
            visibleCount = visibleCount + 1
            lastVisible = i
        End If
    Next i
    If visibleCount <= 1 Then
        MsgBox "Debe quedar al menos una columna de Persona visible.", vbInformation
        Exit Sub
    End If
    If lastVisible <> -1 Then WS.Columns(lastVisible).Hidden = True
End Sub

Private Function NextEntityColumn(WS As Worksheet, headerRow As Long) As Long
    ' Empieza en columna C (3) y avanza hasta la primera vacia en la fila de encabezado
    Dim c As Long: c = 3
    Do While LenB(CStr(WS.Cells(headerRow - 1, c).value)) > 0 Or LenB(CStr(WS.Cells(headerRow, c).value)) > 0
        c = c + 1
    Loop
    NextEntityColumn = c
End Function

' ReadAndSavePersonas is temporarily skipped.

' ==== Seccion Vehiculos ====
Private Function VehiculosStartRow() As Long
    VehiculosStartRow = PersonasStartRow() + 18
End Function

' LayoutVehiculosSection is not needed as layout is managed manually.



Public Sub AgregarColumnaVehiculo()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    Dim firstCol As Long, lastCol As Long
    firstCol = WS.Columns("W").Column ' 23
    lastCol = WS.Columns("Z").Column ' 26
    Dim i As Long, visibleCount As Long
    visibleCount = 0
    For i = firstCol To lastCol
        If WS.Columns(i).Hidden = False Then visibleCount = visibleCount + 1
    Next i
    If visibleCount = 0 Then
        WS.Columns(firstCol).Hidden = False
        Exit Sub
    End If
    For i = firstCol To lastCol
        If WS.Columns(i).Hidden = True Then
            WS.Columns(i).Hidden = False
            Exit Sub
        End If
    Next i
    MsgBox "Ya alcanzaste el máximo de columnas de Vehículos (W:Z).", vbInformation
End Sub

Public Sub EliminarColumnaVehiculo()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    Dim firstCol As Long, lastCol As Long
    firstCol = WS.Columns("W").Column
    lastCol = WS.Columns("Z").Column
    Dim i As Long, visibleCount As Long, lastVisible As Long
    visibleCount = 0: lastVisible = -1
    For i = firstCol To lastCol
        If WS.Columns(i).Hidden = False Then
            visibleCount = visibleCount + 1
            lastVisible = i
        End If
    Next i
    If visibleCount <= 1 Then
        MsgBox "Debe quedar al menos una columna de Vehículo visible.", vbInformation
        Exit Sub
    End If
    If lastVisible <> -1 Then WS.Columns(lastVisible).Hidden = True
End Sub

' ReadAndSaveVehiculos is temporarily skipped.

Public Sub AbrirFormularioIncidenteEnHoja()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    WS.Activate
End Sub

Private Function ReadIncidenteFromSheet(WS As Worksheet) As clsIncidente
    Dim e As New clsIncidente
    e.id_incidente = CStr(WS.Range("D5").value)
    Dim f As Variant, h As Variant
    f = WS.Range("D6").value
    h = WS.Range("D7").value
    If IsDate(f) And IsDate(h) Then
        e.fecha_hora_ocurrencia = CDate(f) + TimeValue(h)
    Else
        e.fecha_hora_ocurrencia = Nz(f, Empty)
    End If
    e.pais = CStr(WS.Range("D8").value)
    e.provincia = CStr(WS.Range("D9").value)
    e.localidad_zona = CStr(WS.Range("D10").value)
    e.coordenadas_geograficas = CStr(WS.Range("D11").value)
    e.lugar_especifico = CStr(WS.Range("D12").value)
    e.uo_incidente = CStr(WS.Range("D13").value)
    e.uo_accidentado = CStr(WS.Range("D14").value)
    e.descripcion_esv = CStr(WS.Range("D15").value)
    e.denuncia_policial = CStr(WS.Range("D20").value)
    e.examen_alcoholemia = CStr(WS.Range("D21").value)
    e.examen_sustancias = CStr(WS.Range("D22").value)
    e.entrevistas_testigos = CStr(WS.Range("D23").value)
    e.accion_inmediata = CStr(WS.Range("D24").value)
    e.consecuencias_seguridad = CStr(WS.Range("D25").value)
    e.fecha_hora_reporte = WS.Range("D26").value
    e.cantidad_personas = WS.Range("D27").value
    e.cantidad_vehiculos = WS.Range("D28").value
    e.clase_evento = CStr(WS.Range("D29").value)
    e.tipo_colision = CStr(WS.Range("D30").value)
    e.nivel_severidad = CStr(WS.Range("D31").value)
    e.clasificacion_esv = CStr(WS.Range("D32").value)
    e.tipo_superficie = CStr(WS.Range("AC6").value)
    e.posee_banquina = CStr(WS.Range("AC7").value)
    e.tipo_ruta = CStr(WS.Range("AC8").value)
    e.densidad_trafico = CStr(WS.Range("AC9").value)
    e.condicion_ruta = CStr(WS.Range("AC10").value)
    e.iluminacion_ruta = CStr(WS.Range("AC11").value)
    e.senalizacion_ruta = CStr(WS.Range("AC12").value)
    e.geometria_ruta = CStr(WS.Range("AC13").value)
    e.condiciones_climaticas = CStr(WS.Range("AC14").value)
    e.rango_temperaturas = CStr(WS.Range("AC15").value)
    Set ReadIncidenteFromSheet = e
End Function

Private Sub ClearForm(WS As Worksheet)
    WS.Range("D5").ClearContents
    WS.Range("D6:D7").ClearContents
    WS.Range("D8:D15").ClearContents
    WS.Range("D20:D32").ClearContents
    WS.Range("AC6:AC15").ClearContents
End Sub

Private Function ValidateForm(WS As Worksheet, ByRef messages As Collection) As Boolean
    Dim ok As Boolean: ok = True
    Set messages = New Collection
    If LenB(CStr(WS.Range("D6").value)) = 0 Then ok = False: messages.Add ("Fecha de ocurrencia es requerida.")
    If LenB(CStr(WS.Range("D8").value)) = 0 Then ok = False: messages.Add ("Pais es requerido.")
    If LenB(CStr(WS.Range("D29").value)) = 0 Then ok = False: messages.Add ("Clase de evento es requerida.")
    If LenB(CStr(WS.Range("D27").value)) > 0 Then If Not IsNumeric(WS.Range("D27").value) Then ok = False: messages.Add ("Cantidad personas debe ser numérico.")
    If LenB(CStr(WS.Range("D28").value)) > 0 Then If Not IsNumeric(WS.Range("D28").value) Then ok = False: messages.Add ("Cantidad vehiculos debe ser numérico.")
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
    WS.Range("D5").value = id
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
    Dim id As String: id = CStr(WS.Range("D5").value)
    If LenB(id) = 0 Then
        MsgBox "No hay ID en D5 para eliminar.", vbExclamation
        Exit Sub
    End If
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Eliminar el incidente " & id & " de forma permanente?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar eliminacion")
    If resp <> vbYes Then Exit Sub
    clsIncidenteRepo.DeleteById id
    ClearForm WS
    MsgBox "Incidente eliminado.", vbInformation
End Sub

Public Sub LoadIncidenteEnHojaDesdeIdActual()
    SetupESVWorkbook
    Dim WS As Worksheet
    Set WS = EnsureFormSheet()
    Dim id As String: id = CStr(WS.Range("D5").value)
    If LenB(id) = 0 Then Exit Sub
    Dim e As clsIncidente
    Set e = clsIncidenteRepo.FindById(id)
    If e Is Nothing Then Exit Sub
    WS.Range("D6").value = e.fecha_hora_ocurrencia
    WS.Range("D8").value = e.pais
    WS.Range("D9").value = e.provincia
    WS.Range("D10").value = e.localidad_zona
    WS.Range("D11").value = e.coordenadas_geograficas
    WS.Range("D12").value = e.lugar_especifico
    WS.Range("D13").value = e.uo_incidente
    WS.Range("D14").value = e.uo_accidentado
    WS.Range("D15").value = e.descripcion_esv
    WS.Range("D20").value = e.denuncia_policial
    WS.Range("D21").value = e.examen_alcoholemia
    WS.Range("D22").value = e.examen_sustancias
    WS.Range("D23").value = e.entrevistas_testigos
    WS.Range("D24").value = e.accion_inmediata
    WS.Range("D25").value = e.consecuencias_seguridad
    WS.Range("D26").value = e.fecha_hora_reporte
    WS.Range("D27").value = e.cantidad_personas
    WS.Range("D28").value = e.cantidad_vehiculos
    WS.Range("D29").value = e.clase_evento
    WS.Range("D30").value = e.tipo_colision
    WS.Range("D31").value = e.nivel_severidad
    WS.Range("D32").value = e.clasificacion_esv
    WS.Range("AC6").value = e.tipo_superficie
    WS.Range("AC7").value = e.posee_banquina
    WS.Range("AC8").value = e.tipo_ruta
    WS.Range("AC9").value = e.densidad_trafico
    WS.Range("AC10").value = e.condicion_ruta
    WS.Range("AC11").value = e.iluminacion_ruta
    WS.Range("AC12").value = e.senalizacion_ruta
    WS.Range("AC13").value = e.geometria_ruta
    WS.Range("AC14").value = e.condiciones_climaticas
    WS.Range("AC15").value = e.rango_temperaturas
End Sub

Public Sub OcultarColumnasPersonas()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    WS.Columns("L:T").Hidden = True
End Sub

Public Sub OcultarColumnasVehiculos()
    Dim WS As Worksheet: Set WS = EnsureFormSheet()
    WS.Columns("X:Z").Hidden = True ' deja W visible como base
End Sub

' Styling automation removed as requested.
