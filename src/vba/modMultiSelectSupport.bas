Attribute VB_Name = "modMultiSelectSupport"
Option Explicit

Public Function ShowMultiSelectForCell(ByVal tgt As Range) As Boolean
    On Error GoTo fin
    If tgt Is Nothing Then Exit Function
    If tgt.Validation.Type <> xlValidateList Then Exit Function

    Dim items As Collection
    Set items = GetValidationItems(tgt)
    If items Is Nothing Or items.Count = 0 Then Exit Function

    Dim existing As Collection
    Set existing = SplitSelection(CStr(tgt.Value))

    Dim frm As Object ' frmMultiSelect
    Set frm = VBA.UserForms.Add("frmMultiSelect")

    With frm
        Set .TargetRange = tgt
        .LoadItems items, existing
        .Show
        If .ResultAccepted Then
            Application.EnableEvents = False
            tgt.Value = JoinCollection(.SelectedItems, ", ")
            Application.EnableEvents = True
            ShowMultiSelectForCell = True
        End If
        Unload frm
    End With
    Exit Function
fin:
    On Error Resume Next
    Application.EnableEvents = True
End Function

Public Function GetValidationItems(ByVal tgt As Range) As Collection
    On Error GoTo fin
    Dim items As New Collection
    Dim src As String
    src = tgt.Validation.Formula1
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

Public Function SplitSelection(ByVal s As String) As Collection
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

Public Function JoinCollection(ByVal col As Collection, ByVal sep As String) As String
    Dim i As Long
    Dim s As String
    For i = 1 To col.Count
        If LenB(s) > 0 Then s = s & sep
        s = s & CStr(col(i))
    Next i
    JoinCollection = s
End Function
