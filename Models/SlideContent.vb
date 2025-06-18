Namespace WordSlideGenerator
    Public Class SlideContent
        Public Property Title As String = ""
        Public Property Text As String = ""
        Public Property SpeakerNotes As String = ""
        Public Property ImageDescription As String = ""
        Public Property Notes As String = ""
        Public Property SlideType As SlideContentType

        Public Sub New()
        End Sub

        Public Sub New(title As String, slideType As SlideContentType)
            Me.Title = title
            Me.SlideType = slideType
        End Sub

        Public Function HasImage() As Boolean
            Return Not String.IsNullOrWhiteSpace(ImageDescription)
        End Function

        Public Function GetCompleteNotes() As String
            Dim completeNotes As String = ""

            If Not String.IsNullOrWhiteSpace(Notes) Then
                completeNotes = "APPUNTI:" & vbCrLf & Notes & vbCrLf & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(SpeakerNotes) Then
                completeNotes &= SpeakerNotes
            End If

            If HasImage() Then
                If completeNotes <> "" Then completeNotes &= vbCrLf & vbCrLf
                completeNotes &= "IMMAGINE SUGGERITA:" & vbCrLf & ImageDescription
            End If

            Return completeNotes
        End Function
    End Class

    Public Enum SlideContentType
        CourseModule
        Lesson
        Content
    End Enum
End Namespace