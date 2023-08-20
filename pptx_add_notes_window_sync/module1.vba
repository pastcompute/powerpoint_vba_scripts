
Dim X As New Class1
Sub InitializeApp()
    ' Debug.Print "InitializeApp"
    Set X.PPTApp = Application
End Sub

Sub PrintWindowSizes()
    Dim w As DocumentWindow

    Let c = ActivePresentation.Windows.Count
    ' Debug.Print "PrintWindowSizes Windows.Count==" & c

    For Each w In Application.Windows
        Debug.Print w.WindowState & "," & w.Left & "+" & w.Top & "+" & w.Width & "x" & w.Height
    Next
End Sub

Sub PositionWindows()
    ' Debug.Print "PositionWindows"
    Call PrintWindowSizes
    Dim a As DocumentWindow
    Dim w As DocumentWindow
    Set a = ActivePresentation.Windows.Item(1)
    
    Let c = ActivePresentation.Windows.Count
    Debug.Print "PositionWindows Windows.Count==" & c
    
    If ActivePresentation.Windows.Count = 1 Then
        Set w = ActivePresentation.NewWindow
    Else
        Set w = ActivePresentation.Windows.Item(2)
    End If
    a.Left = 137
    a.Top = 41
    a.Width = 1276
    a.Height = 933
    a.ViewType = ppViewNormal
    
    w.Left = 1453
    w.Top = 41
    w.Width = 470
    w.Height = 933
    w.ViewType = ppViewNotesPage

End Sub

Sub OnSlideChange()
    ' Debug.Print "OnSlideChange(). ViewType==" & ActiveWindow.ViewType & " ppViewNotesPage==" & ppViewNotesPage
    
    Dim currentSlide As slide
    Dim currentWindow As DocumentWindow
    Dim targetWindow As DocumentWindow
    Dim slideIndex As Long
    
    ' Check if in normal view
    If ActiveWindow.ViewType = ppViewNormal Then
        Set currentWindow = ActiveWindow
        Set currentSlide = currentWindow.View.slide
        Debug.Print "currentSlide=='" & currentSlide.Name & "'"
        
        ' Loop through all open windows
        For Each targetWindow In Application.Windows
            ' Debug.Print "targetWindow.ViewType==" & targetWindow.ViewType
            If targetWindow.ViewType = ppViewNotesPage And Not targetWindow Is currentWindow Then
                slideIndex = currentSlide.slideIndex
                ' Debug.Print "Current slide index: " & slideIndex
                
                If slideIndex <= currentSlide.Parent.Slides.Count And slideIndex > 0 Then
                    targetWindow.View.GotoSlide slideIndex
                    ' Debug.Print "Synchronized slide in target window."
                End If
            End If
        Next targetWindow
    End If
    CountNotHiddenSlides1
End Sub

' TODO fixme make the other one a function that we call, and remove this duplicated code urgh
Sub CountNotHiddenSlides()
    Dim pptPresentation As Presentation
    Dim slide As slide
    Dim notHiddenCount As Integer
    
    Set pptPresentation = ActivePresentation
    
    notHiddenCount = 0    
    For Each slide In pptPresentation.Slides
        If Not slide.SlideShowTransition.Hidden Then
            notHiddenCount = notHiddenCount + 1
        End If
    Next slide
    MsgBox "Slides=" & pptPresentation.Slides.Count & " Visible=" & notHiddenCount
End Sub

Sub CountNotHiddenSlides1()
    Dim pptPresentation As Presentation
    Dim slide As slide
    Dim notHiddenCount As Integer
    
    Set pptPresentation = ActivePresentation
    
    notHiddenCount = 0    
    For Each slide In pptPresentation.Slides
        If Not slide.SlideShowTransition.Hidden Then
            notHiddenCount = notHiddenCount + 1
        End If
    Next slide
    
    Debug.Print "Slides=" & pptPresentation.Slides.Count & " Visible=" & notHiddenCount
End Sub



