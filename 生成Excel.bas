Attribute VB_Name = "����Excel"
Sub ����Excel()
    ActiveWorkbook.Save
    
    Dim targetdir As String
    Dim origindir As String
    targetdir = ActiveWorkbook.Path & "\" & "Excel�ļ���\"
    origindir = ActiveWorkbook.Path
    
    ChDir targetdir
    
    ActiveWorkbook.SaveAs Filename:= _
       targetdir & Range("B2") & "-" & Range("B3") & "-" & Range("B4") & "-" & Range("B6") & "-" & "Excelԭʼ�ļ�.xlsm", FileFormat:= _
        xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        
'
'
'    Workbooks.Open Filename:=OriginalFile
'
    
    
End Sub




