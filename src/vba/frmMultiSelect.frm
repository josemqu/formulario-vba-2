VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMultiSelect 
   Caption         =   "Seleccionar opciones"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lbItems 
      Height          =   3600
      Left            =   120
      Top             =   120
      Width           =   6960
      MultiSelect     =   1 ' fmMultiSelectMulti
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   4440
      Top             =   4080
      Width           =   1320
      Default         =   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   5880
      Top             =   4080
      Width           =   1200
      Cancel          =   -1  'True
   End
End
Attribute VB_Name = "frmMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ResultAccepted As Boolean
Public SelectedItems As Collection
Public TargetRange As Range

Public Sub LoadItems(ByVal items As Collection, ByVal preselected As Collection)
    Dim i As Long
    Me.lbItems.Clear
    For i = 1 To items.Count
        Me.lbItems.AddItem CStr(items(i))
    Next i
    ' Preseleccionar
    If Not preselected Is Nothing Then
        Dim j As Long, t As String
        For j = 0 To Me.lbItems.ListCount - 1
            t = CStr(Me.lbItems.List(j))
            If ContainsText(preselected, t) Then
                Me.lbItems.Selected(j) = True
            End If
        Next j
    End If
End Sub

Private Sub cmdOK_Click()
    Set SelectedItems = New Collection
    Dim i As Long
    For i = 0 To Me.lbItems.ListCount - 1
        If Me.lbItems.Selected(i) Then SelectedItems.Add CStr(Me.lbItems.List(i))
    Next i
    ResultAccepted = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    ResultAccepted = False
    Me.Hide
End Sub

Private Function ContainsText(ByVal col As Collection, ByVal txt As String) As Boolean
    On Error GoTo fin
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(CStr(col(i)), txt, vbTextCompare) = 0 Then ContainsText = True: Exit Function
    Next i
fin:
End Function
