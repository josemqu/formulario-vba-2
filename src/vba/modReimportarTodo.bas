Attribute VB_Name = "modReimportarTodo"

Option Explicit

' Requiere referencia: Microsoft Visual Basic for Applications Extensibility 5.3
' Trust Center: habilitar "Trust access to the VBA project object model"

' Tipos de componente (evita depender de constantes de VBIDE)
Private Const CT_STD_MODULE As Long = 1      ' vbext_ct_StdModule
Private Const CT_CLASS_MODULE As Long = 2    ' vbext_ct_ClassModule
Private Const CT_MSFORM As Long = 3          ' vbext_ct_MSForm
Private Const CT_DOCUMENT As Long = 100      ' vbext_ct_Document

Public Sub ReimportarTodoElCodigo()
    Dim vbProj As Object
    Dim comp As Object
    Dim i As Long
    Dim base As String
    Dim modPath As String
    Dim clsPath As String
    Dim f As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    base = ThisWorkbook.Path & "\src\vba"
    modPath = base & "\"
    clsPath = base & "\"
    
    ' Paso 1: Eliminar todos los m?dulos y clases (excepto este m?dulo y hojas/formularios)
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        
        ' Solo eliminar m?dulos est?ndar y clases
        ' NO eliminar hojas (Document) ni formularios (MSForm)
        If comp.Type = CT_STD_MODULE Or comp.Type = CT_CLASS_MODULE Then
            ' No eliminar este m?dulo mientras se ejecuta
            If comp.name <> "modReimportarTodo" Then
                On Error Resume Next
                vbProj.VBComponents.Remove comp
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' Paso 2: Importar m?dulos est?ndar (.bas)
    f = Dir(modPath & "*.bas")
    Do While Len(f) > 0
        vbProj.VBComponents.Import modPath & f
        f = Dir
    Loop
    
    ' Paso 3: Importar clases (.cls) con encabezado para que VBIDE las reconozca como Class Modules
    f = Dir(clsPath & "*.cls")
    Do While Len(f) > 0
        Dim src As String
        src = clsPath & f
        vbProj.VBComponents.Import src
        f = Dir
    Loop

    ' Paso 4: Corregir si alguna clase se importó como módulo estándar por error
    FixMisImportedClasses vbProj

    ' Paso 5: Limpiar encabezados inválidos en clases (si los hubiera)
    For i = 1 To vbProj.VBComponents.Count
        Set comp = vbProj.VBComponents(i)
        If comp.Type = CT_CLASS_MODULE Then
            CleanClassHeaders comp.CodeModule
        End If
    Next i
    
    ' MsgBox "Reimportación completada." & vbCrLf & _
    '        "Módulos y clases eliminados y reimportados desde:" & vbCrLf & _
    '        "- " & modPath & vbCrLf & _
    '        "- " & clsPath, vbInformation, "Reimportación"
End Sub


Public Sub ForzarClasePorNombre(ByVal compName As String)
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    Dim comp As Object
    On Error Resume Next
    Set comp = vbProj.VBComponents(compName)
    On Error GoTo 0
    If comp Is Nothing Then Exit Sub
    If comp.Type = CT_CLASS_MODULE Then
        CleanClassHeaders comp.CodeModule
        Exit Sub
    End If
    If comp.Type = CT_STD_MODULE Then
        Dim cm As Object: Set cm = comp.CodeModule
        Dim txt As String
        If cm.CountOfLines > 0 Then txt = cm.lines(1, cm.CountOfLines)
        Dim newComp As Object
        Set newComp = vbProj.VBComponents.Add(CT_CLASS_MODULE)
        On Error Resume Next
        newComp.name = compName & "_tmp_cls"
        On Error GoTo 0
        Dim cmNew As Object: Set cmNew = newComp.CodeModule
        If cmNew.CountOfLines > 0 Then cmNew.DeleteLines 1, cmNew.CountOfLines
        If LenB(txt) > 0 Then cmNew.AddFromString txt
        On Error Resume Next
        vbProj.VBComponents.Remove comp
        newComp.name = compName
        On Error GoTo 0
        CleanClassHeaders cmNew
    End If
End Sub

Private Sub FixMisImportedClasses(vbProj As Object)
    Dim i As Long
    For i = vbProj.VBComponents.Count To 1 Step -1
        Dim comp As Object
        Set comp = vbProj.VBComponents(i)
        If comp.Type = CT_STD_MODULE Then
            Dim cm As Object
            Set cm = comp.CodeModule
            Dim txt As String
            If cm.CountOfLines > 0 Then
                txt = cm.lines(1, cm.CountOfLines)
            Else
                txt = ""
            End If
            ' Heurística: si contiene WithEvents o atributos típicos de clases, debe ser clase
            If InStr(1, txt, "WithEvents", vbTextCompare) > 0 _
               Or InStr(1, txt, "Attribute VB_PredeclaredId", vbTextCompare) > 0 _
               Or InStr(1, txt, "VERSION 1.0 CLASS", vbTextCompare) > 0 Then
                Dim oldName As String: oldName = comp.name
                Dim newComp As Object
                Set newComp = vbProj.VBComponents.Add(CT_CLASS_MODULE)
                On Error Resume Next
                newComp.name = oldName & "_tmp_cls"
                On Error GoTo 0
                Dim cmNew As Object
                Set cmNew = newComp.CodeModule
                If cmNew.CountOfLines > 0 Then cmNew.DeleteLines 1, cmNew.CountOfLines
                If LenB(txt) > 0 Then cmNew.AddFromString txt
                ' Eliminar el módulo antiguo y renombrar el nuevo con el nombre original
                On Error Resume Next
                vbProj.VBComponents.Remove comp
                newComp.name = oldName
                On Error GoTo 0
                ' Limpieza de encabezados por si quedaron
                CleanClassHeaders cmNew
            End If
        End If
    Next i
End Sub

Public Function PreprocessClassFile(srcPath As String) As String
    Dim fso As Object, ts As Object, content As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(srcPath, 1)
    content = ts.ReadAll
    ts.Close
    content = NormalizeLineEndings(content)
    content = StripClassHeader(content)
    Dim tmpPath As String
    tmpPath = srcPath & ".tmp"
    Set ts = fso.OpenTextFile(tmpPath, 2, True)
    ts.Write content
    ts.Close
    PreprocessClassFile = tmpPath
End Function

Public Function StripClassHeader(ByVal content As String) As String
    Dim lines() As String
    lines = Split(content, vbCrLf)
    Dim i As Long, outLines() As String, outCount As Long
    ReDim outLines(0 To UBound(lines))
    Dim inHeader As Boolean
    inHeader = False
    If UBound(lines) >= 0 Then
        If Left$(Trim$(lines(0)), 7) = "VERSION" Then inHeader = True
    End If
    Dim ended As Boolean
    ended = Not inHeader
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = lines(i)
        If Not ended Then
            If Trim$(ln) = "End" Or Trim$(ln) = "END" Then
                ended = True
            End If
        Else
            outLines(outCount) = ln
            outCount = outCount + 1
        End If
    Next i
    If outCount = 0 Then
        StripClassHeader = content
    Else
        ReDim Preserve outLines(0 To outCount - 1)
        StripClassHeader = Join(outLines, vbCrLf)
    End If
End Function

Private Function NormalizeLineEndings(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, vbCrLf)
    NormalizeLineEndings = s
End Function

Public Sub CleanClassHeaders(cm As Object)
    ' Elimina, si existen, líneas de encabezado no válidas dentro del editor
    ' como: VERSION 1.0 CLASS / BEGIN / END al inicio del módulo
    Dim maxCheck As Long: maxCheck = Application.WorksheetFunction.Min(10, cm.CountOfLines)
    Dim i As Long
    Dim removed As Boolean
    ' Repetir hasta que ya no encuentre encabezados en las primeras líneas
    Do
        removed = False
        If cm.CountOfLines = 0 Then Exit Do
        ' Revisar primeras líneas posibles del encabezado exportado
        For i = 1 To maxCheck
            Dim ln As String
            ln = Trim$(cm.lines(i, 1))
            Dim t As String
            t = UCase$(ln)
            If t Like "VERSION * CLASS" Or t = "BEGIN" Or t = "END" Or Left$(t, 8) = "MULTIUSE" Then
                cm.DeleteLines i, 1
                removed = True
                Exit For
            ElseIf Left$(ln, 9) = "Attribute" Then
                ' Atributos son v?lidos, no tocarlos
            ElseIf Len(ln) = 0 Then
                cm.DeleteLines i, 1
                removed = True
                Exit For
            End If
        Next i
    Loop While removed
End Sub

Public Sub EliminarModReimportarTodo()
    ' Este procedimiento elimina el m?dulo modReimportarTodo despu?s de usarlo
    ' Ejecutar SOLO despu?s de haber ejecutado ReimportarTodoElCodigo
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("modReimportarTodo")
    On Error GoTo 0
    
    ' MsgBox "Modulo modReimportarTodo eliminado.", vbInformation
End Sub

Public Sub LimpiarEncabezadosDeTodasLasClases()
    Dim vbProj As Object
    Dim comp As Object
    Set vbProj = Application.VBE.ActiveVBProject
    For Each comp In vbProj.VBComponents
        If comp.Type = CT_CLASS_MODULE Then
            CleanClassHeaders comp.CodeModule
        End If
    Next comp
End Sub



