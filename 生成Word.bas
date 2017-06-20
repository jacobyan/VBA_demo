Attribute VB_Name = "����Word"
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
    wordFileName = basic_grade & basic_semister & "-" & basic_sequence & "-" & basic_classTitle & "-" & "��ʦ����.docx"
    cwd = ActiveWorkbook.Path
    cword = cwd & "\Word�ļ���\" & wordFileName
    '��ȡ�ĵ�ģ�� getParentFolder(cwd) & "\�ĵ�ģ��.dotx"
    Set wrdDoc = wrdApp.Documents.Add(Template:= _
        cwd & "\�ĵ�ģ��.dotx", NewTemplate:=False, DocumentType:=0)
     ' or
     'Set wrdDoc = wrdApp.Documents.Open("C:\Foldername\Filename.doc")
     ' sample word operations
    With wrdDoc
                With .content
                .InsertAfter basic_classTitle
                .Paragraphs(.Paragraphs.Count).style = wrdDoc.Styles("����")
                .InsertParagraphAfter
                .InsertParagraphAfter
                End With
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''' д�������Ϣ��
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                With .content
                Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=3, NumColumns:= _
                        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    Set startLine = wrdDoc.Range(Start:=tblNew.Cell(1, 1).Range.Start, _
                            End:=tblNew.Cell(1, 4).Range.End)
                    With tblNew
                        .Cell(1, 1).Range.InsertAfter "��ѧ������Ϣ"    '=== ��1�� ===
                        .Cell(2, 1).Range.InsertAfter "�ڿ��꼶"        '=== ��2�� ===
                        .Cell(2, 2).Range.InsertAfter basic_grade
                        .Cell(2, 3).Range.InsertAfter "��ѧ����"        '=== ��2�� ===
                        .Cell(2, 4).Range.InsertAfter basic_studentsInAGroup
                        .Cell(3, 1).Range.InsertAfter "�ο��̲�"        '=== ��3�� ===
                        .Cell(3, 2).Range.InsertAfter basic_bookName
                        .Cell(3, 3).Range.InsertAfter "���ÿ�ʱ"        '=== ��3�� ===
                        .Cell(3, 4).Range.InsertAfter basic_classDuriation
                    
                    .Columns.AutoFit
                    End With
                    startLine.Cells.Merge
                    tblNew.Cell(1, 1).Range.style = wrdDoc.Styles("��ǿ")
                tblNew.AutoFitBehavior (wdAutoFitWindow)
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                
                 .InsertParagraphAfter
                 .InsertParagraphAfter
                End With
                
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '''' \  д�������Ϣ /
                '''' �������-> 9��44��
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                curRow = 9
                With .content
                        Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=1, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                        
                        With tblNew.Cell(1, 1)

                            '============================
                            .Range.InsertAfter "��ѧĿ��"
                            .Range.Paragraphs(.Range.Paragraphs.Count).style = wrdDoc.Styles("���� 1")
                            .Range.InsertParagraphAfter
                            '-------------------
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="֪ʶ�뼼��Ŀ��", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 2"))
                            
                            '-------------------
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="���̬�����ֵ��Ŀ��", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 2"))
                            
                            
                            '=============================
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="��ѧ�ص�", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            
                            '=============================
                            curRow = miniBlock(curBlockRows:=3, curRow:=curRow, miniTitle:="��ѧ�ѵ�", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            
                            '=============================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="ѧ��֪ʶ", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            
                            '===========================================
                            .Range.InsertAfter "��ѧ׼��"
                            .Range.Paragraphs(.Range.Paragraphs.Count).style = wrdDoc.Styles("���� 1")
                            .Range.InsertParagraphAfter
                            '--------------------
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 2"))
                            '--------------------
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 2"))
                            '--------------------
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="ý����Դ", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 2"))
                            '--------------------
                            curRow = miniBlock(curBlockRows:=2, curRow:=curRow, miniTitle:="����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 2"))

                        End With
                        .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "�γ̵���"    '=== ��1�� ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "��ѧ����"    '=== ��1�� ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="����Ŀ��", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="��ʶ����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="��ʶ����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="׼������", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=49, curRow:=curRow, miniTitle:="��ʼ����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=7, curRow:=curRow, miniTitle:="��ɢ˼ά", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="�ܽ����", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "�κ�����"    '=== ��1�� ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="����Ŀ��", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
                        End With
                    End With 'tblNew
                    .InsertParagraphAfter
                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Set tblNew = .Tables.Add(Range:=.Paragraphs(.Paragraphs.Count).Range, NumRows:=2, NumColumns:= _
                        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitWindow)
                    With tblNew
                        With .Cell(1, 1)
                            .Range.InsertAfter "��ѧ��˼"    '=== ��1�� ===
                        End With
                        .Columns.AutoFit
                        .AutoFitBehavior (wdAutoFitWindow)
                        With .Cell(2, 1)
                            '===========================================
                            curRow = miniBlock(curBlockRows:=1, curRow:=curRow, miniTitle:="���÷�˼���", _
                            mysheet:=curSheet, curRange:=.Range, docStyle:=wrdDoc.Styles("���� 1"))
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
    curRange.Paragraphs(curRange.Paragraphs.Count).style = docStyle 'wrdDoc.Styles("���� 2")
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
    Dim parseResult() As String '����
    Dim curContent As String
    Dim curPicPath As String
    Dim ub
    Dim lb
    myPicPath = imageFolderName
    curContent = wordsToParse
    Debug.Print "curContent: " & curContent
    parseResult = parseContent(curContent)
    wrdRange.InsertAfter parseResult(0) 'ͼ���е���������
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
        MsgBox ("��ʽ�����⣡")
    End If
End Sub

                            '=============================
    '                        curBlockRows = 3
    '                        .Range.InsertAfter "֪ʶ�뼼��Ŀ��"
    '                        .Range.Paragraphs(.Range.Paragraphs.Count).Style = wrdDoc.Styles("���� 2")
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
