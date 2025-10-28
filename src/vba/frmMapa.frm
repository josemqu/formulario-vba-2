VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMapa 
   Caption         =   "Mapa"
   ClientHeight    =   6090
   ClientLeft      =   -380
   ClientTop       =   -1800
   ClientWidth     =   9150.001
   OleObjectBlob   =   "frmMapa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim htmlPath As String
    htmlPath = "C:\web\formulario-vba-2\mapa.html"
    Me.wbMapa.Navigate htmlPath
End Sub

Private Sub wbMapa_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next

    Dim coords As String
    ' Lee "lat,lng" desde Form!celdaCoordenadas
    coords = Trim$(ThisWorkbook.Sheets("Form").Range("celdaCoordenadas").Text)
    If Len(coords) = 0 Then Exit Sub

    ' Escapar comilla simple para el JS
    coords = Replace(coords, "'", "\'")

    ' Llama a la función del HTML para ubicar el marcador
    Dim js As String
    js = "setMarkerFromHostText('" & coords & "');"
    Me.wbMapa.Object.Document.parentWindow.execScript js, "JavaScript"
End Sub

Private Sub wbMapa_TitleChange(ByVal Text As String)
    Dim s As String
    s = Trim$(Text)
    If Left$(s, 11) <> "coords_str:" Then Exit Sub

    s = Mid$(s, 12) ' "lat,lng" como texto
    With ThisWorkbook.Sheets("Form")
        .Range("celdaCoordenadas").value = s
    End With
End Sub
