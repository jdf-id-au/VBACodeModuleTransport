Attribute VB_Name = "Core"
' Facilitate version control of VBA modules within workbooks.
'
' Build:
' 1. Open a new workbook in Excel, open VB Editor, import VBACodeModuleTransport.Core.bas
' 2. Check Tools > References... > Microsoft Visual Basic for Applications Extensibility 5.3
' 3. Rename the project to 'VBACodeModuleTransport' in Properties inspector.
' 4. Set ThisWorkbook 'IsAddin' property to True.
' 5. File > Save as "VBACodeModuleTransport.xlam"
'
' Install:
' 1. Open target workbook in Excel (ensure file type xlsm).
' 2. Check Tools > Add-Ins > Vbacodemoduletransport.Xlam
' 3. Open VB Editor and select VBAProject corresponding to target workbook.
' 4. Check Tools > References... > VBACodeModuleTransport
'
' Call: (manually or in response to events)
'   VBACodeModuleTransport.Export Application.ThisWorkbook
'   VBACodeModuleTransport.Import Application.ThisWorkbook
'
' Jeremy Field 2015-12-13
'
' https://github.com/hilkoc/vbaDeveloper/blob/master/src/vbaDeveloper.xlam/Build.bas
' http://stackoverflow.com/a/2003792/780743
' http://stackoverflow.com/a/30127696/780743
' http://www.cpearson.com/excel/vbe.aspx
' (unfortunately vbaDeveloper only works on Windows because of Microsoft Scripting Runtime dependency)
'
' vbext_ct_ types are from VBA Extensibility Type Library (see Tools > References...)
' FIXME this module itself is imported/exported manually to VBACoreModuleTransport.Core.bas
' TODO should store exports in directory named after the source workbook

Option Explicit

Private Function Path(wb As Workbook, module_name As String, module_type As Integer) As String
    Dim ext As String
    Select Case module_type
        Case vbext_ct_StdModule
            ext = ".bas"
        Case vbext_ct_ClassModule, vbext_ct_Document
            ext = ".cls"
    End Select
    Path = wb.Path & Application.PathSeparator & module_name & ext
End Function

Public Sub Export(wb As Workbook)
    ' Export all non-empty Standard, Class and Document Modules from Workbook wb to its parent directory.
    Dim i As Integer
    Dim c As VBIDE.VBComponent
    With wb.VBProject
        For i = 1 To .VBComponents.Count
            Set c = .VBComponents(i)
            Select Case c.Type
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
                    If c.CodeModule.CountOfLines > 0 Then
                        c.Export Path(wb, c.CodeModule.Name, c.Type)
                    End If
            End Select
        Next i
    End With
End Sub

Public Sub Import(wb As Workbook)
    ' Replace existing Standard Modules in Workbook wb with corresponding files from its parent directory.
    Dim i As Integer
    Dim c As VBIDE.VBComponent
    Dim original_name As String
    Dim original_type As Integer
    Dim p As String
    Dim readline As String
    Dim readlines As String
    Dim line_count As Integer
    With wb.VBProject
        For i = .VBComponents.Count To 1 Step -1
            Set c = .VBComponents(i)
            Select Case c.Type
                ' Encountered impenetrable 50053 error importing [non-MSO] class module.
                ' Workaround by deleting and inserting lines for all class modules.
                Case vbext_ct_StdModule
                    original_name = c.Name
                    original_type = c.Type
                    ' Prevent name change when new module imported
                    c.Name = c.Name & "_remove"
                    p = Path(wb, original_name, original_type)
                    On Error GoTo ErrFileMissing
                    .VBComponents.Import p
                    On Error GoTo 0 ' Deregister this procedure's error handler(s).
                    .VBComponents.Remove c
SkipModule:
                    On Error GoTo 0
                Case vbext_ct_ClassModule, vbext_ct_Document
                    line_count = 0
                    readlines = ""
                    p = Path(wb, c.Name, c.Type)
                    On Error GoTo ErrClassFileMissing
                    Open p For Input As #1
                    Do Until EOF(1)
                        Line Input #1, readline
                        If line_count > 8 Then ' skip header
                            readlines = readlines & readline & vbNewLine
                        End If
                        line_count = line_count + 1
                    Loop
                    Close #1
                    On Error GoTo 0
                    c.CodeModule.DeleteLines 1, c.CodeModule.CountOfLines ' FIXME is this adding a line with each cycle?
                    c.CodeModule.InsertLines c.CodeModule.CountOfLines + 1, readlines ' c.CodeModule.CountOfLines should be 0
SkipClass:
                    On Error GoTo 0
            End Select
        Next i
    End With
    Exit Sub
ErrFileMissing:
    If Err.Number = 57 Or Err.Number = 53 Then ' Device I/O Error, File Not Found Error
        Resume SkipModule
    Else
        ' FIXME this sucks. How to re-throw? Err.Raise errno doesn't work.
        MsgBox "Unhandled error " & Err.Number & " during Import (within add-in): " & p
    End If
    Exit Sub
ErrClassFileMissing:
    If Err.Number = 75 Or Err.Number = 53 Then ' Path/File access error, File Not Found Error
        Resume SkipClass
    Else
        MsgBox "Unhandled error " & Err.Number & " during class Import (within add-in): " & p
    End If
End Sub
