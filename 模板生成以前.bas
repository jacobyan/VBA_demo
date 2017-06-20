Attribute VB_Name = "模板生成以前"
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

    '使用模板新建一个PPT
    PPApp.Presentations.Open Filename:="C:\Users\Jacob\Desktop\课程\666课程开发模板宏工具\ppt模板文件夹\0标题页.potx", Untitled:=msoTrue
'    将新建的PPT赋值给PPPres，这个PPT
    Set PPPres = PPApp.ActivePresentation

'添加标题
    
    Set TotalSlide = PPApp.ActivePresentation.Slides(1)
    TotalSlide.Shapes(1).TextFrame.TextRange.Text = Range("B4") & ". " & Range("B8")
    

'添加课程导入

    Dim slide_num As Byte
    
    slide_num = 2
    Do While Range("B" & (43 + slide_num)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "课程导入"
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
    


'制作目标

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "制作目标"
            .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B52"), "QuestionTitle")
            .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B52"), "ContentText")
        '    .Shapes(3).Fill.UserPicture ("C:\Users\Jacob\Desktop\课程\课程图片库\arduino跑表.jpg")
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
    
    
    

    
    
'认识材料
'    startRange = 53
'    slide_max_num = 5
'    addContent
'

    
    Dim slide_num_begin As Byte
    slide_num_begin = slide_num
    
    Do While Range("B" & (53 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "认识材料"
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
    
    
    
    

    
    
'认识工具


    startpage = 0
    
    Do While Range("B" & (60 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "认识工具"
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
    




'制作准备

    startpage = 0
    
    Do While Range("B" & (67 + startpage)) > " "

        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
        With TotalSlide
            .Shapes(1).TextFrame.TextRange.Text = "制作准备"
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



'开始制作

'    startpage = 0
'
'    Do While Range("B" & (74 + startpage)) > " "
'
'        Set TotalSlide = PPApp.ActivePresentation.Slides(slide_num + 1)
'        With TotalSlide
'            .Shapes(1).TextFrame.TextRange.Text = "开始制作"
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


'发散思维












''   PPApp.ActivePresentation.Slides.Add PPApp.ActivePresentation.Slides.Count + 1, ppLayoutText
    PPApp.ActiveWindow.ViewType = ppViewNormal
    With PPPres
        .SaveAs "C:\Users\Jacob\Desktop\测试临时\MyPreso2332.pptx"
        
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
    Dim mMatches As MatchCollection      '匹配字符串集合对象
    Dim mMatch As Match        '匹配字符串

    Set mRegExp = New RegExp
    With mRegExp
        .Global = True                              'True表示匹配所有, False表示仅匹配第一个符合项
        .IgnoreCase = True                          'True表示不区分大小写, False表示区分大小写

        Select Case Types
            Case "QuestionTitle"
                    .Pattern = "^([\u4e00-\u9fa5]|[a-zA-Z])*(？|\n|\?|\r)"   '匹配字符模式
            Case "ContentText"
                    .Pattern = "[^(？|\n|\?|\r)]*。+?"   '匹配字符模式
            Case "ImagePath"
                    .Pattern = "[^@]*jpg|\.jpeg|\.gif"   '匹配字符模式
            Case Else
                MsgBox ("没有匹配项")

        End Select

        Set mMatches = .Execute(TargetRange.Text)   '执行正则查找，返回所有匹配结果的集合，若未找到，则为空
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
            .Shapes(1).TextFrame.TextRange.Text = "认识材料"
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



