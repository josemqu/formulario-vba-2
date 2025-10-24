Attribute VB_Name = "modUUID"
Option Explicit

Public Function NewUUID() As String
    Dim guid As String
    guid = CreateObject("Scriptlet.TypeLib").guid
    NewUUID = Mid$(guid, 2, Len(guid) - 2)
End Function

Public Function NowIso() As String
    NowIso = Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
End Function

Public Function UserNameOrDefault() As String
    On Error Resume Next
    UserNameOrDefault = Environ$("Username")
    If Len(UserNameOrDefault) = 0 Then UserNameOrDefault = "usuario"
End Function

Public Function NewShortId() As String
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Incidentes")
    If Not ws Is Nothing Then Set lo = ws.ListObjects("tbIncidente")
    On Error GoTo 0

    Dim prefix As String
    prefix = "INC-" & Format$(Date, "yyyymmdd") & "-"

    Dim maxN As Long: maxN = 0
    If Not lo Is Nothing Then
        Dim col As ListColumn
        On Error Resume Next
        Set col = lo.ListColumns("id_incidente")
        On Error GoTo 0
        If Not col Is Nothing Then
            Dim vals As Variant
            If Not col.DataBodyRange Is Nothing Then
                vals = col.DataBodyRange.Value
                Dim one As String, n As Long, r As Long
                If IsArray(vals) Then
                    For r = 1 To UBound(vals, 1)
                        one = CStr(vals(r, 1))
                        If Left$(one, Len(prefix)) = prefix Then
                            n = Val(Mid$(one, Len(prefix) + 1))
                            If n > maxN Then maxN = n
                        End If
                    Next r
                Else
                    one = CStr(vals)
                    If Left$(one, Len(prefix)) = prefix Then
                        n = Val(Mid$(one, Len(prefix) + 1))
                        If n > maxN Then maxN = n
                    End If
                End If
            End If
        End If
    End If

    NewShortId = prefix & Format$(maxN + 1, "000")
End Function
