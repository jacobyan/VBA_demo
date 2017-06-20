Attribute VB_Name = "ģ��������ǰ"
    Public PPApp As PowerPoint.Application
    Public PPPres As PowerPoint.Presentation
    Public TotalSlide As Object
    Public PPSlide As PowerPoint.Slide
    
    
    
    
Sub ExcelToNewPowerPoint()

    
    Dim startpage As Byte
    Dim slide_max_num As Byte
    startpage = 0
    slide_max_num = 5
    
    
    Dim ImageFilePath As String
    
    ImageFilePath = Range("E1")
    
    ' Create instance of PowerPoint
    Set PPApp = CreateObject("Powerpoint.Application")
    ' For automation to work, PowerPoint must be visible
    ' (alternatively, other extraordinary measures must be taken)
    PPApp.Visible = True

'    ' Create a presentation
'    Set PPPres = PPApp.Presentations.Add

    'ʹ��ģ���½�һ��PPT
    PPApp.Presentations.Open Filename:="C:\Users\Jacob\Desktop\�γ�\666�γ̿���ģ��깤��\pptģ���ļ���\0����ҳ.potx", Untitled:=msoTrue
'    ���½���PPT��ֵ��PPPres�����PPT
    Set PPPres = PPApp.ActivePresentation

'��ӱ���
    
    Set TotalSlide = PPApp.ActivePresentation.Slides(1)
    TotalSlide.Shapes(1).TextFrame.TextRange.Text = Range("B4") & ". " & Range("B8")
    

'��ӿγ̵���

    Dim slide_num As Byte
    
    slide_num = 2
    Do While Range("B" & (43 + slide_num)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "�γ̵���"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (43 + slide_num)), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (43 + slide_num)), "ContentText")
        End With
        
        Dim oPPtShp As PowerPoint.Shape
        For Each oPPtShp In TotalSlide.Shapes
            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                With oPPtShp
                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (43 + slide_num)), "ImagePath"), msoFalse, msoTrue, _
                                    .Left, .Top, .Width, .Height

                    DoEvents
                End With
            End If
        Next
        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6

    slide_num = slide_num + 1
    Loop

    Dim slide_const_num As Byte
    Dim slide_detect_start As Byte
    slide_const_num = 6
    For slide_detect_start = 1 To (slide_const_num - slide_num + 1) Step 1
    PPApp.ActivePresentation.Slides(slide_num).Delete
    Next slide_detect_start
    


'����Ŀ��

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "����Ŀ��"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B52"), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B52"), "ContentText")
        '    .Shapes(3).Fill.UserPicture ("C:\Users\Jacob\Desktop\�γ�\�γ�ͼƬ��\arduino�ܱ�.jpg")
        End With
        
'        Dim oPPtShp As PowerPoint.Shape
        For Each oPPtShp In TotalSlide.Shapes
            '~~> You only need to work on Picture place holders
            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                With oPPtShp
                    '~~> Now add the Picture
                    '~~> For this example, picture path is in Cell A1
                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B52"), "ImagePath"), msoFalse, msoTrue, _
                                    .Left, .Top, .Width, .Height
                    
                    '~~> Insert DoEvents here specially for big files, or network files
                    '~~> DoEvents halts macro momentarily until the
                    '~~> system finishes what it's doing which is loading the picture file
                    DoEvents
                End With
            End If
        Next
        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
    
    
    

    
    
'��ʶ����
'    startRange = 53
'    slide_max_num = 5
'    addContent
'

    
    Dim slide_num_begin As Byte
    slide_num_begin = slide_num
    
    Do While Range("B" & (53 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "��ʶ����"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (53 + startpage)), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (53 + startpage)), "ContentText")
        End With
        
'        Dim oPPtShp As PowerPoint.Shape
        For Each oPPtShp In TotalSlide.Shapes
            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                With oPPtShp
                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (53 + startpage)), "ImagePath"), msoFalse, msoTrue, _
                                    .Left, .Top, .Width, .Height

                    DoEvents
                End With
            End If
        Next
        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6

    startpage = startpage + 1
    slide_num = slide_num + 1
    Loop

'MsgBox (slide_num)

'    Dim slide_const_num As Byte
'    Dim slide_detect_start As Byte
'    slide_const_num = slide_max_num
'
    For slide_detect_start = 0 To (slide_const_num - startpage) Step 1
    PPApp.ActivePresentation.Slides(slide_num + 1).Delete
    Next slide_detect_start
    
    
    
    

    
    
'��ʶ����


    startpage = 0
    
    Do While Range("B" & (60 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "��ʶ����"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (60 + startpage)), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (60 + startpage)), "ContentText")
        End With
        
'        Dim oPPtShp As PowerPoint.Shape
        For Each oPPtShp In TotalSlide.Shapes
            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                With oPPtShp
                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (60 + startpage)), "ImagePath"), msoFalse, msoTrue, _
                                    .Left, .Top, .Width, .Height

                    DoEvents
                End With
            End If
        Next
        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6

    startpage = startpage + 1
    slide_num = slide_num + 1
    Loop

'MsgBox (slide_num)

'    Dim slide_const_num As Byte
'    Dim slide_detect_start As Byte
'    slide_const_num = slide_max_num
'
    For slide_detect_start = 0 To (5 - startpage) Step 1
    PPApp.ActivePresentation.Slides(slide_num + 1).Delete
    Next slide_detect_start
    




'����׼��

    startpage = 0
    
    Do While Range("B" & (67 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "����׼��"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (67 + startpage)), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (67 + startpage)), "ContentText")
        End With
        
'        Dim oPPtShp As PowerPoint.Shape
        For Each oPPtShp In TotalSlide.Shapes
            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                With oPPtShp
                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (67 + startpage)), "ImagePath"), msoFalse, msoTrue, _
                                    .Left, .Top, .Width, .Height

                    DoEvents
                End With
            End If
        Next
        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6

    startpage = startpage + 1
    slide_num = slide_num + 1
    Loop

'MsgBox (slide_num)

'    Dim slide_const_num As Byte
'    Dim slide_detect_start As Byte
'    slide_const_num = slide_max_num
'
    For slide_detect_start = 0 To (5 - startpage) Step 1
    PPApp.ActivePresentation.Slides(slide_num + 1).Delete
    Next slide_detect_start



'��ʼ����

'    startpage = 0
'
'    Do While Range("B" & (74 + startpage)) > " "
'
'        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
'        With TotalSlide
'            .Shapes(1).TextFrame.TextRange.Text = "��ʼ����"
'            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (74 + startpage)), "QuestionTitle")
'            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (74 + startpage)), "ContentText")
'        End With
'
''        Dim oPPtShp As PowerPoint.Shape
'        For Each oPPtShp In TotalSlide.Shapes
'            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
'                With oPPtShp
'                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (74 + startpage)), "ImagePath"), msoFalse, msoTrue, _
'                                    .Left, .Top, .Width, .Height
'
'                    DoEvents
'                End With
'            End If
'        Next
'        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
'        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
'
'    startpage = startpage + 1
'    slide_num = slide_num + 1
'    Loop
'
''MsgBox (slide_num)
'
''    Dim slide_const_num As Byte
''    Dim slide_detect_start As Byte
''    slide_const_num = slide_max_num
''
'    For slide_detect_start = 0 To (20 - startpage) Step 1
'    PPApp.ActivePresentation.Slides(slide_num + 1).Delete
'    Next slide_detect_start


'��ɢ˼ά












''   PPApp.ActivePresentation.Slides.Add PPApp.ActivePresentation.Slides.Count + 1, ppLayoutText
    PPApp.ActiveWindow.ViewType = ppViewNormal
    With PPPres
        .SaveAs "C:\Users\Jacob\Desktop\������ʱ\MyPreso2332.pptx"
        
'        .Close
    End With

    ' Quit PowerPoint
'    PPApp.Quit

    ' Clean up
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set PPApp = Nothing

    
    
End Sub


Function GetString(TargetRange As Range, Types As String) As String



    Dim mRegExp As RegExp
    Dim mMatches As MatchCollection      'ƥ���ַ������϶���
    Dim mMatch As Match        'ƥ���ַ���

    Set mRegExp = New RegExp
    With mRegExp
        .Global = True                              'True��ʾƥ������, False��ʾ��ƥ���һ��������
        .IgnoreCase = True                          'True��ʾ�����ִ�Сд, False��ʾ���ִ�Сд

        Select Case Types
            Case "QuestionTitle"
                    .Pattern = "^([\u4e00-\u9fa5]|[a-zA-Z])*(��|\n|\?|\r)"   'ƥ���ַ�ģʽ
            Case "ContentText"
                    .Pattern = "[^(��|\n|\?|\r)]*��+?"   'ƥ���ַ�ģʽ
            Case "ImagePath"
                    .Pattern = "[^@]*jpg|\.jpeg|\.gif"   'ƥ���ַ�ģʽ
            Case Else
                MsgBox ("û��ƥ����")

        End Select

        Set mMatches = .Execute(TargetRange.Text)   'ִ��������ң���������ƥ�����ļ��ϣ���δ�ҵ�����Ϊ��
        For Each mMatch In mMatches
            GetString = SumValueInText + (mMatch.Value)
        Next
    End With

    Set mRegExp = Nothing
    Set mMatches = Nothing
End Function



Sub addContent()
    
    Do While Range("B" & (53 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "��ʶ����"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (53 + startpage)), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (53 + startpage)), "ContentText")
        End With
        
'        Dim oPPtShp As PowerPoint.Shape
        For Each oPPtShp In TotalSlide.Shapes
            If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                With oPPtShp
                    TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (53 + startpage)), "ImagePath"), msoFalse, msoTrue, _
                                    .Left, .Top, .Width, .Height

                    DoEvents
                End With
            End If
        Next
        TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
        TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6

    startpage = startpage + 1
    slide_num = slide_num + 1
    Loop

'MsgBox (slide_num)

'    Dim slide_const_num As Byte
'    Dim slide_detect_start As Byte
    slide_const_num = slide_max_num
    
    For slide_detect_start = 1 To (slide_const_num - startpage) Step 1
    PPApp.ActivePresentation.Slides(slide_num).Delete
    Next slide_detect_start
    
'    If startRange <> 53 Then
'        slide_num = slide_num - 1
'
'
'    End If

End Sub



