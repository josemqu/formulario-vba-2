VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents App As Application
Public Sub Init()
    Set App = Application
End Sub
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo fin
    If Sh Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If Sh.name <> "Catalogos" Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Dim cols As Collection
    Set cols = New Collection

    Dim c As Range
    For Each c In Target.Columns
        cols.Add c.Column
    Next c

    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")

    Application.EnableEvents = False
    Dim colNum As Variant
    For Each colNum In cols
        If Not processed.Exists(colNum) Then
            processed(colNum) = True
            Dim header As String
            header = CStr(Sh.Cells(1, colNum).value)
            If LenB(header) > 0 Then
                Dim colLetter As String
                colLetter = Split(Sh.Cells(1, colNum).Address(False, False), CStr(Sh.Cells(1, colNum).Row))(0)
                EnsureCatalog Sh, colLetter, header
            End If
        End If
    Next colNum

fin:
    On Error Resume Next
    Application.EnableEvents = True
End Sub
