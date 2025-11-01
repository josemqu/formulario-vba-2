Attribute VB_Name = "modMultiSelectMenu"
Option Explicit

' Local constants to avoid dependency on Office type library constants
Private Const msoBarPopup As Long = 5
Private Const msoControlButton As Long = 1
Private Const msoControlSeparator As Long = 3
Private Const msoButtonIconAndCaption As Long = 3
Private Const msoButtonUp As Long = 0
Private Const msoButtonDown As Long = 1

Private Const POPUP_NAME As String = "msel_popup"
Private m_targetAddress As String
Private m_targetSheet As String

Public Sub ShowMultiSelectMenu(ByVal tgt As Range)
    On Error GoTo fin
    If tgt Is Nothing Then Exit Sub
    If tgt.Cells.CountLarge <> 1 Then Exit Sub
    ' Guardar contexto del target
    m_targetAddress = tgt.Address(False, False, xlA1, False)
    m_targetSheet = tgt.Parent.Name

    ' Construir lista permitida desde la validación
    Dim items As Collection: Set items = GetValidationItems(tgt)
    If items Is Nothing Or items.Count = 0 Then Exit Sub

    ' Selecciones actuales
    Dim selected As Collection: Set selected = SplitSelection(CStr(tgt.Value))

    ' Destruir popup previo
    On Error Resume Next
    Application.CommandBars(POPUP_NAME).Delete
    On Error GoTo 0

    Dim cb As CommandBar
    Set cb = Application.CommandBars.Add(Name:=POPUP_NAME, Position:=msoBarPopup, Temporary:=True)

    Dim i As Long
    For i = 1 To items.Count
        Dim cap As String: cap = CStr(items(i))
        Dim btn As CommandBarButton
        Set btn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
        btn.Caption = cap
        btn.Style = msoButtonIconAndCaption
        btn.FaceId = 0
        btn.OnAction = "'" & ThisWorkbook.Name & "'!modMultiSelectMenu.ToggleFromPopup"
        If ContainsText(selected, cap) Then
            btn.State = msoButtonDown
        Else
            btn.State = msoButtonUp
        End If
    Next i

    ' Botón limpiar (separado visualmente)
    Dim btnClear As CommandBarButton
    Set btnClear = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btnClear.Caption = "(Limpiar)"
    btnClear.Style = msoButtonIconAndCaption
    btnClear.FaceId = 0
    btnClear.BeginGroup = True
    btnClear.OnAction = "'" & ThisWorkbook.Name & "'!modMultiSelectMenu.ClearFromPopup"

    ' Mostrar popup bajo el cursor
    cb.ShowPopup
    Exit Sub
fin:
    ' En silencio
End Sub

Public Sub ToggleFromPopup()
    On Error GoTo fin
    Dim ctrl As Object ' CommandBarControl
    Set ctrl = Application.CommandBars.ActionControl
    If ctrl Is Nothing Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(m_targetSheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error GoTo fin
    If ws Is Nothing Then Exit Sub

    Dim tgt As Range
    Set tgt = ws.Range(m_targetAddress)
    If tgt Is Nothing Then Exit Sub

    Dim current As String: current = CStr(tgt.Value)
    Dim parts As Collection: Set parts = SplitSelection(current)

    Dim cap As String: cap = CStr(ctrl.Caption)
    If ContainsText(parts, cap) Then
        RemoveText parts, cap
        On Error Resume Next: ctrl.State = msoButtonUp: On Error GoTo fin
    Else
        parts.Add cap
        On Error Resume Next: ctrl.State = msoButtonDown: On Error GoTo fin
    End If

    Application.EnableEvents = False
    tgt.Value = JoinCollection(parts, ", ")
    Application.EnableEvents = True
    ' Cerrar popup actual y reabrir limpio
    On Error Resume Next
    Application.CommandBars(POPUP_NAME).Visible = False
    Application.CommandBars(POPUP_NAME).Delete
    On Error GoTo fin
    DoEvents
    ' Re-mostrar el popup para permitir seleccionar varias opciones sin cerrarlo
    ShowMultiSelectMenu tgt
    Exit Sub
fin:
    Application.EnableEvents = True
End Sub

Public Sub ClearFromPopup()
    On Error GoTo fin
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(m_targetSheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error GoTo fin
    If ws Is Nothing Then Exit Sub
    Application.EnableEvents = False
    ws.Range(m_targetAddress).ClearContents
    Application.EnableEvents = True
    ' Cerrar popup actual y reabrir limpio
    On Error Resume Next
    Application.CommandBars(POPUP_NAME).Visible = False
    Application.CommandBars(POPUP_NAME).Delete
    On Error GoTo fin
    DoEvents
    ' Re-mostrar el popup para continuar seleccionando si se desea
    ShowMultiSelectMenu ws.Range(m_targetAddress)
    Exit Sub
fin:
    Application.EnableEvents = True
End Sub

Private Function GetValidationItems(ByVal tgt As Range) As Collection
    On Error GoTo fin
    If tgt.Validation.Type <> xlValidateList Then Exit Function
    Dim src As String: src = tgt.Validation.Formula1
    Dim items As New Collection

    If LenB(src) = 0 Then Set GetValidationItems = items: Exit Function

    If Left$(Trim$(src), 1) = "=" Then
        Dim ref As String: ref = Mid$(src, 2)
        Dim rng As Range
        On Error Resume Next
        Set rng = tgt.Parent.Range(ref)
        If rng Is Nothing Then Set rng = ThisWorkbook.Names(ref).RefersToRange
        On Error GoTo 0
        If Not rng Is Nothing Then
            Dim c As Range
            For Each c In rng.Cells
                If LenB(CStr(c.Value)) > 0 Then items.Add CStr(c.Value)
            Next c
        End If
    Else
        Dim arr() As String
        Dim i As Long
        arr = Split(src, ",")
        For i = LBound(arr) To UBound(arr)
            items.Add Trim$(arr(i))
        Next i
    End If

    Set GetValidationItems = items
    Exit Function
fin:
    Set GetValidationItems = Nothing
End Function

Private Function SplitSelection(ByVal s As String) As Collection
    Dim col As New Collection
    Dim arr() As String
    Dim i As Long, t As String
    If LenB(Trim$(s)) = 0 Then Set SplitSelection = col: Exit Function
    arr = Split(s, ",")
    For i = LBound(arr) To UBound(arr)
        t = Trim$(arr(i))
        If LenB(t) > 0 Then col.Add t
    Next i
    Set SplitSelection = col
End Function

Private Function ContainsText(ByVal col As Collection, ByVal txt As String) As Boolean
    On Error GoTo fin
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(CStr(col(i)), txt, vbTextCompare) = 0 Then ContainsText = True: Exit Function
    Next i
fin:
End Function

Private Sub RemoveText(ByRef col As Collection, ByVal txt As String)
    On Error GoTo fin
    Dim i As Long
    For i = col.Count To 1 Step -1
        If StrComp(CStr(col(i)), txt, vbTextCompare) = 0 Then
            col.Remove i
        End If
    Next i
fin:
End Sub

Private Function JoinCollection(ByVal col As Collection, ByVal sep As String) As String
    Dim i As Long
    Dim s As String
    For i = 1 To col.Count
        If LenB(s) > 0 Then s = s & sep
        s = s & CStr(col(i))
    Next i
    JoinCollection = s
End Function
