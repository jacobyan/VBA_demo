Attribute VB_Name = "生成Excel"
Sub 生成Excel()
    ActiveWorkbook.Save
    
    Dim targetdir As String
    Dim origindir As String
    targetdir = ActiveWorkbook.Path & "\" & "Excel文件夹\"
    origindir = ActiveWorkbook.Path
    
    ChDir targetdir
    
    ActiveWorkbook.SaveAs Filename:= _
       targetdir & Range("B2") & "-" & Range("B3") & "-" & Range("B4") & "-" & Range("B6") & "-" & "Excel原始文件.xlsm", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        
'
'
'    Workbooks.Open Filename:=OriginalFile
'
    
    
End Sub




