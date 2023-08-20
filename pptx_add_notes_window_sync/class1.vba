Public WithEvents PPTApp As Application

Private Sub PPTApp_SlideSelectionChanged(ByVal SldRange As SlideRange)
    Debug.Print "PPTApp_SlideSelectionChanged"
    
    Module1.OnSlideChange
    

End Sub
