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

Private WithEvents mApp As Application

Public Property Set App(ByVal v As Application)
    Set mApp = v
End Property

Private Sub mApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Sh.Name = "Formulario" Then
        If Not Application.Intersect(Target, Sh.Range("C2")) Is Nothing Then
            LoadIncidenteEnHojaDesdeIdActual
        End If
    End If
End Sub
