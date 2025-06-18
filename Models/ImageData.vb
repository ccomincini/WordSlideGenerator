Namespace WordSlideGenerator
    Public Class ImageData
        Public Property SlideIndex As Integer
        Public Property Description As String
        Public Property PosLeft As Integer
        Public Property PosTop As Integer
        Public Property PosWidth As Integer
        Public Property PosHeight As Integer

        Public Sub New()
        End Sub

        Public Sub New(slideIndex As Integer, description As String, left As Integer, top As Integer, width As Integer, height As Integer)
            Me.SlideIndex = slideIndex
            Me.Description = description
            Me.PosLeft = left
            Me.PosTop = top
            Me.PosWidth = width
            Me.PosHeight = height
        End Sub
    End Class
End Namespace