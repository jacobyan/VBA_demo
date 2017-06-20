Attribute VB_Name = "生成Word"
Sub Gene_Teach_Doc()
     ' to test this code, paste it into an Excel module
     ' add a reference to the Word-library
     ' create a new folder named C:\Foldername or edit the filnames in the code
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Dim wrdPic As Word.InlineShape
    imageFolderName = Sheet2.Cells(1, 5).Value
    myPicPath = imageFolderName
    Dim cwd As String
    Dim i As Integer
    Dim curSheet As Object
    Dim curWorkbook As Workbook
    Dim wordsToParse As String
    Set curSheet = Sheet2
    Set curWorkbook = ThisWorkbook
    Set wrdApp = CreateObject("Word.Application")
    Dim basic_classTitle As String
    Dim basic_grade As String
    Dim basic_semister As String
    Dim basic_sequence As String
    Dim basic_studentsInAGroup As String
    Dim basic_bookName As String
    Dim basic_classDuriation As String
    Dim curRow As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''
    basic_classTitle = curSheet.Cells(8, 2).Value
    basic_grade = curSheet.Cells(2, 2).Value
    basic_semister = curSheet.Cells(3, 2).Value
    basic_sequence = curSheet.Cells(4, 2).Value
    basic_studentsInAGroup = curSheet.Cells(5, 2).Value
    basic_bookName = curSheet.Cells(6, 2).Value
    basic_classDuriation = curSheet.Cells(7, 2).Value
    
    wrdApp.Visible = True
    wordFileName = basic_grade & basic_semister & "-" & basic_sequence & "-" & basic_classTitle & "-" & "教师用书.docx"
    cwd = ActiveWorkbook.Path
    cword = cwd & "\Word文件夹\" & wordFileName
    '获取文档模板 getParentFolder(cwd) & "\文档模板.dotx"
    Set wrdDoc = wrdApp.Documents.Add(Template:= _
        cwd & "\文档模板.dotx", NewTemplate:=False, DocumentType:=0)
     ' or
     'Set wrdDoc = wrdApp.Documents.Open("C:\Foldername\Filename.doc")
     ' sample word operations
    With wrdDoc
                With .content
                .InsertAfter basic_classTitle
                .Paragraphs(.Paragraphs.Count).style = wrdDoc.Styles("标题")
                .InsertParagraphAfter
                .InsertParagraphAfter
                End With
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''' 写入基本信息栏
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                With .content
                Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=3, NumColumns:= _
                        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    Set startLine = wrdDoc.Range(Start:=tblNew.Cell(1, 1).Range.Start, _
                            End:=tblNew.Cell(1, 4).Range.End)
                    With tblNew
                        .Cell(1, 1).Range.InsertAfter "教学基本信息"    '=== 第1排 ===
                        .Cell(2, 1).Range.InsertAfter "授课年级"        '=== 第2排 ===
                        .Cell(2, 2).Range.InsertAfter basic_grade
                        .Cell(2, 3).Range.InsertAfter "教学分组"        '=== 第2排 ===
                        .Cell(2, 4).Range.InsertAfter basic_studentsInAGroup
                        .Cell(3, 1).Range.InsertAfter "参考教材"        '=== 第3排 ===
                        .Cell(3, 2).Range.InsertAfter basic_bookName
                        .Cell(3, 3).Range.InsertAfter "设置课时"        '=== 第3排 ===
                        .Cell(3, 4).Range.InsertAfter basic_classDuriation
                    
                    .Columns.AutoFit
                    End With
                    startLine.Cells.Merge
                    tblNew.Cell(1, 1).Range.style = wrdDoc.Styles("增强")
                tblNew.AutoFitBehavior (wdAutoFitWindow)
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                
                 .InsertParagraphAfter
                 .InsertParagraphAfter
                End With
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''' \  写入概述信息 /
                '''' 表格行数-> 9至44行
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                curRow = 9
                With .content
                        Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=1, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                        
                        With tblNew.Cell(1, 1)

                            '============================
                            .Range.InsertAfter "教学目标"
                            .Range.Paragraphs(.Range.Paragraphs.Count).style = wrdDoc.Styles("标题 1")
                            .Range.InsertParagraphAfter
                            '-------------------
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="知识与技能目标", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 2"))
                            
                            '-------------------
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="情感态度与价值观目标", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 2"))
                            
                            
                            '=============================
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="教学重点", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            
                            '=============================
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="教学难点", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            
                            '=============================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="学科知识", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            
                            '===========================================
                            .Range.InsertAfter "教学准备"
                            .Range.Paragraphs(.Range.Paragraphs.Count).style = wrdDoc.Styles("标题 1")
                            .Range.InsertParagraphAfter
                            '--------------------
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="材料", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 2"))
                            '--------------------
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="工具", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 2"))
                            '--------------------
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="媒体资源", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 2"))
                            '--------------------
                            curRow = miniBlock(curBlockRows:=2, curRow:=curRow, miniTitle:="其他", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 2"))

                        End With
                        .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "课程导入"    '=== 第1排 ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="引入", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "教学流程"    '=== 第1排 ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="制作目标", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="认识材料", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="认识工具", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="准备制作", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=49, curRow:=curRow, miniTitle:="开始制作", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="发散思维", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="总结分享", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "课后整理"    '=== 第1排 ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="整理目标", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "教学反思"    '=== 第1排 ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="课堂反思表格", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("标题 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
        
        End With 'content
        If Dir(cword) <> "" Then
            Kill cword
        End If
        .SaveAs (cword)
        '.Close ' close the document
     End With 'wordApplication
'    wrdApp.Quit ' close the Word application
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
End Sub

Function miniBlock(curBlockRows As Integer, curRow As Integer, miniTitle As String, _
mysheet As Excel.Worksheet, curRange As Word.Range, docStyle As Object)
    Dim wordsToParse As String
    'curBlockRows = 3
    curRange.InsertAfter miniTitle
    curRange.Paragraphs(curRange.Paragraphs.Count).style = docStyle 'wrdDoc.Styles("标题 2")
    curRange.InsertParagraphAfter
    For i = curRow To curRow + curBlockRows - 1
        wordsToParse = mysheet.Cells(i, 2).Value
        If Len(wordsToParse) > 0 Then
            curRange.InsertParagraphAfter
            Call addContent(wordsToParse, curRange)
        End If
    Next i
    miniBlock = curRow + curBlockRows
End Function


Sub addContent(wordsToParse As String, wrdRange As Word.Range)
    Dim parseResult() As String '数组
    Dim curContent As String
    Dim curPicPath As String
    Dim ub
    Dim lb
    myPicPath = imageFolderName
    curContent = wordsToParse
    Debug.Print "curContent: " & curContent
    parseResult = parseContent(curContent)
    wrdRange.InsertAfter parseResult(0) '图文中的文字内容
    wrdRange.InsertParagraphAfter
    ub = UBound(parseResult)
    lb = LBound(parseResult)
    'If (ub < 5) And (ub > 1) Then
    If (ub < 10) And (ub > 1) Then
        For i = 1 To UBound(parseResult) / 2
            picName = parseResult(2 * i - 1)
            Debug.Print "curPicPath: " & curPicPath
            If Len(Dir(myPicPath & "\" & picName)) Then
               curPicPath = myPicPath & "\" & picName
               'wrdRange.InlineShapes.AddPicture().Range.InsertCaption ("")
                Set wrdPic = wrdRange.InlineShapes.AddPicture(curPicPath, False, True, wrdRange.Paragraphs(wrdRange.Paragraphs.Count).Range)
                            
                            'wrdPic.Title = parseResult(2 * i)  'picName
                            wrdPic.Range.InsertCaption Label:="Figure", title:=" :" & parseResult(2 * i), Position:=wdCaptionPositionBelow
                            wrdPic.AlternativeText = picName
'                            wrdPic.Height = (200# / wrdPic.Width) * wrdPic.Height
'                            wrdPic.Width = 200 '(200# / wrdPic.Height) * wrdPic.Width
                            
                            wrdPic.Width = (200# / wrdPic.Height) * wrdPic.Width
                            wrdPic.Height = 200
                            
            End If
        Next i
    ElseIf ub > 10 Then
        MsgBox ("格式有问题！")
    End If
End Sub

                            '=============================
    '                        curBlockRows = 3
    '                        .Range.InsertAfter "知识与技能目标"
    '                        .Range.Paragraphs(.Range.Paragraphs.Count).Style = wrdDoc.Styles("标题 2")
    '                        .Range.InsertParagraphAfter
    '
    '                        For i = curRow To curRow + curBlockRows - 1
    '                            wordsToParse = curSheet.Cells(i, 2).Value
    '                            If Len(wordsToParse) > 0 Then
    '                                .Range.InsertParagraphAfter
    '                                Call addContent(wordsToParse, .Range)
    '                            End If
    '                        Next i
    '                        curRow = curRow + curBlockRows
