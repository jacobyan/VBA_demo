Attribute VB_Name = "生成PPT"
    Public ImageFilePath As String
    Public PPApp As PowerPoint.Application
    Public PPPres As PowerPoint.Presentation
    Public TotalSlide As Object
    Public PPSlide As PowerPoint.Slide
    
    
    
    
    
Sub ExcelToNewPowerPoint()
    
    ImageFilePath = Range("E1")


    ' Create instance of PowerPoint
    Set PPApp = CreateObject("Powerpoint.Application")
    ' For automation to work, PowerPoint must be visible
    ' (alternatively, other extraordinary measures must be taken)
    PPApp.Visible = True

'    ' Create a presentation
'    Set PPPres = PPApp.Presentations.Add

    '使用模板新建一个PPT
    PPApp.Presentations.Open Filename:=ActiveWorkbook.Path & "\PPT生成模板.potx", Untitled:=msoTrue
'    将新建的PPT赋值给PPPres，这个PPT
    Set PPPres = PPApp.ActivePresentation

'添加标题
    
    Set TotalSlide = PPApp.ActivePresentation.Slides(1)
    TotalSlide.Shapes(1).TextFrame.TextRange.Text = Range("B4") & ". " & Range("B8")
    
    
    Dim startRange As Byte
    startRange = 0
    Dim Num_of_Image As Byte
    Num_of_Image = 0
    Dim cusLayout As CustomLayout
    
For startRange = 0 To 95 Step 1
    If Range("B" & (45 + startRange)) > " " Then

        Num_of_Image = GetString(Range("B" & (45 + startRange)), "ImagePath", 0)  '0表示返回图片数量, 先计算有多少图片
'        MsgBox (GetString(Range("B" & (45 + startRange)), "ImagePath", 0))
        
        
        Select Case Num_of_Image
            Case Is = 1
                If GetString(Range("B" & (45 + startRange)), "QuestionTitle", 1) > " " Then
                
                    Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(2)  '课程导入,单张图片+标题
                    PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout
                    Call addContentToSlide("oneImage_title", 45 + startRange)
                Else
                    Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(3)  '单张图片 无标题
                    PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout
                    Call addContentToSlide("oneImage", 45 + startRange)
                End If
                
            Case Is = 2
                    Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(4)  '2张图片
                    PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout
                    Call addContentToSlide("twoImage", 45 + startRange)
            
            Case Is = 3
                    Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(5)  '3张图片
                    PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout
                    Call addContentToSlide("threeImage", 45 + startRange)
            
            Case Is = 4
                    Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(6)  '4张图片
                    PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout
                    Call addContentToSlide("fourImage", 45 + startRange)
                    
            Case Is = 0
                        Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(8)  '无图片
                        PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout
                        Call addContentToSlide("noImage", 45 + startRange)
            Case Else
            
        End Select
    Else
    
        
        
    End If
    Next startRange
    
                        Set cusLayout = PPApp.ActivePresentation.SlideMaster.CustomLayouts(7)  '结尾
                        PPApp.ActivePresentation.Slides.AddSlide PPApp.ActivePresentation.Slides.Count + 1, cusLayout





''   PPApp.ActivePresentation.Slides.Add PPApp.ActivePresentation.Slides.Count + 1, ppLayoutText
    PPApp.ActiveWindow.ViewType = ppViewNormal
    With PPPres
        .SaveAs ActiveWorkbook.Path & "\" & "PPT文件夹\" & Range("B2") & "-" & Range("B3") & "-" & Range("B4") & "-" & Range("B6") & "-" & "教师指导PPT.pptx"
        
'        .Close
    End With

    ' Quit PowerPoint
'    PPApp.Quit

    ' Clean up
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set PPApp = Nothing

'    MsgBox GetString(Range("B45"), "QuestionTitle")
    
End Sub

Function GetString(TargetRange As Range, Types As String, ImageReturn As Byte) As String


    Dim mRegExp As RegExp
    Dim mMatches As MatchCollection      '匹配字符串集合对象
    Dim mMatch As Match        '匹配字符串
    
    Dim ImageNum As Byte
    ImageNum = 0

    Set mRegExp = New RegExp
    With mRegExp
        .Global = True                              'True表示匹配所有, False表示仅匹配第一个符合项
        .IgnoreCase = True                          'True表示不区分大小写, False表示区分大小写
        
        Select Case Types
            Case "QuestionTitle"
                    .Pattern = "^[^？]+？"   '匹配字符模式 ok
                    Set mMatches = .Execute(TargetRange.Text)
                    For Each mMatch In mMatches
                        GetString = SumValueInText + (mMatch.Value)
                    Next
                    
            Case "AnswerText"
                    .Pattern = "？([\s\S]*?)@"   '匹配字符模式 ok
                    Set mMatches = .Execute(TargetRange.Text)
                     For Each mMatch In mMatches
                        GetString = SumValueInText + (mMatch.Value)
                     Next
                     
                     
                     If GetString > " " Then
                        GetString = mMatches.Item(0).SubMatches(0)  ' 提取捕获内容
                    Else
                        GetString = "缺内容"
                    End If
                    
                    
            Case "InstructionText"
                    .Pattern = "^[^@]+"   '匹配字符模式 ok
                    Set mMatches = .Execute(TargetRange.Text)
                    For Each mMatch In mMatches
                        GetString = SumValueInText + (mMatch.Value)
                     Next
                    
              Case "TeachProcess"
                    .Pattern = "[\u4e00-\u9fa5]{4,5}\b"   '匹配字符模式 ok
                    Set mMatches = .Execute(TargetRange.Text)
                    GetString = mMatches.Item(0).Value
                    
                    
            Case "ImagePath"
                    .Pattern = "[^@]*(jpg|gif|png|jpeg)"   '匹配字符模式 ok
                    Set mMatches = .Execute(TargetRange.Text)
'                    if mMatches.Item(0).value <>
                    
                    If mMatches.Count > 1 Then
                        ImageNum = (mMatches.Count) / 2
                    Else
                        ImageNum = mMatches.Count
                    End If
                    
                    
                    Select Case ImageReturn
                        Case Is = 0
                            GetString = ImageNum
                        Case Is = 1
                            GetString = mMatches.Item(0).Value
                         Case Is = 2
                            GetString = mMatches.Item(2).Value
                        Case Is = 3
                            GetString = mMatches.Item(4).Value
                         Case Is = 4
                            GetString = mMatches.Item(6).Value
                    Case Else
                    End Select
                    
                    
                    
            Case Else
                
        End Select
                 

        

    End With
    
    Set mRegExp = Nothing
    Set mMatches = Nothing
End Function

Function addContentToSlide(typeOfslide As String, RangeNum As Byte)

    Select Case typeOfslide
    
         Case Is = "noImage"
            Set TotalSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
            With TotalSlide
                .Shapes(1).TextFrame.TextRange.Text = GetString(Range("A" & (RangeNum)), "TeachProcess", 0)
                .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "InstructionText", 0)
'                .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "AnswerText", 0)
            End With
    
    
        Case Is = "oneImage_title"
            Set TotalSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
            With TotalSlide
                .Shapes(1).TextFrame.TextRange.Text = GetString(Range("A" & (RangeNum)), "TeachProcess", 0)
                .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "QuestionTitle", 0)
                .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "AnswerText", 0)
            End With
            Dim oPPtShp As PowerPoint.Shape
            For Each oPPtShp In TotalSlide.Shapes
                If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                    With oPPtShp
                        TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 1), msoFalse, msoTrue, _
                                        .Left, .Top, .Width, .Height
                        DoEvents
                    End With
                End If
            Next
            TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
            TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
          
    Case Is = "oneImage"
           Set TotalSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
            With TotalSlide
                .Shapes(1).TextFrame.TextRange.Text = GetString(Range("A" & (RangeNum)), "TeachProcess", 0)
                .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "InstructionText", 0)
'                .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "AnswerText", 0)
            End With
'            Dim oPPtShp As PowerPoint.Shape
            For Each oPPtShp In TotalSlide.Shapes
                If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                    With oPPtShp
                        TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 1), msoFalse, msoTrue, _
                                        .Left, .Top, .Width, .Height
                        DoEvents
                    End With
                End If
            Next
            TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
            TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
            
    Case Is = "twoImage"
            Set TotalSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
            With TotalSlide
                .Shapes(1).TextFrame.TextRange.Text = GetString(Range("A" & (RangeNum)), "TeachProcess", 0)
                .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "InstructionText", 0)
'                .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "AnswerText", 0)
            End With
'            Dim oPPtShp As PowerPoint.Shape
            For Each oPPtShp In TotalSlide.Shapes
                If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                
                    Select Case oPPtShp.Name
                        Case "Picture Placeholder 3"
                            With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 1), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
                            End With
                        Case "Picture Placeholder 4"
                             With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 2), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(4).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(4).PictureFormat.Contrast = 0.6
                            End With
                        Case Else
                    End Select
                End If
            Next
        
        
        
        
    Case Is = "threeImage"
            Set TotalSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
            With TotalSlide
                .Shapes(1).TextFrame.TextRange.Text = GetString(Range("A" & (RangeNum)), "TeachProcess", 0)
                .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "InstructionText", 0)
'                .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "AnswerText", 0)
            End With
'            Dim oPPtShp As PowerPoint.Shape
            For Each oPPtShp In TotalSlide.Shapes
                If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                
                    Select Case oPPtShp.Name
                        Case "Picture Placeholder 3"
                            With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 1), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
                            End With
                        Case "Picture Placeholder 4"
                             With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 2), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(4).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(4).PictureFormat.Contrast = 0.6
                            End With
                        Case "Picture Placeholder 5"
                             With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 3), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(5).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(5).PictureFormat.Contrast = 0.6
                            End With
                        Case Else
                    End Select
                End If
            Next


    Case Is = "fourImage"
            Set TotalSlide = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
            With TotalSlide
                .Shapes(1).TextFrame.TextRange.Text = GetString(Range("A" & (RangeNum)), "TeachProcess", 0)
                .Shapes(2).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "InstructionText", 0)
'                .Shapes(4).TextFrame.TextRange.Text = GetString(Range("B" & (RangeNum)), "AnswerText", 0)
            End With
'            Dim oPPtShp As PowerPoint.Shape
            For Each oPPtShp In TotalSlide.Shapes
                If oPPtShp.PlaceholderFormat.Type = ppPlaceholderPicture Then
                
                    Select Case oPPtShp.Name
                        Case "Picture Placeholder 3"
                            With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 1), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(3).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(3).PictureFormat.Contrast = 0.6
                            End With
                        Case "Picture Placeholder 4"
                             With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 2), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(4).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(4).PictureFormat.Contrast = 0.6
                            End With
                        Case "Picture Placeholder 5"
                             With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 3), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(5).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(5).PictureFormat.Contrast = 0.6
                            End With
                         Case "Picture Placeholder 6"
                             With oPPtShp
                                TotalSlide.Shapes.AddPicture ImageFilePath & GetString(Range("B" & (RangeNum)), "ImagePath", 4), msoFalse, msoTrue, _
                                                .Left, .Top, .Width, .Height
                                DoEvents
                                TotalSlide.Shapes(6).PictureFormat.Brightness = 0.6
                                TotalSlide.Shapes(6).PictureFormat.Contrast = 0.6
                            End With
                            
                            
                            
                        Case Else
                    End Select
                End If
            Next





        Case Else
    End Select
    
    


End Function
'


