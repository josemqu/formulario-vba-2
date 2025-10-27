Attribute VB_Name = "modAppEventsInit"
Option Explicit

Public gAppEvents As clsAppEvents

Public Sub Auto_Open()
    InitAppEvents
End Sub

Public Sub InitAppEvents()
    On Error Resume Next
    If gAppEvents Is Nothing Then
        Set gAppEvents = New clsAppEvents
    End If
    gAppEvents.Init
End Sub
