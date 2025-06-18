Namespace WordSlideGenerator
    Public Class ImageManager
        Private _images As New List(Of ImageData) ' Cambiato il nome del campo privato
        Private logger As Logger

        Public ReadOnly Property ImageCount As Integer
            Get
                Return _images.Count
            End Get
        End Property

        Public ReadOnly Property Images As List(Of ImageData) ' Proprietà pubblica
            Get
                Return _images ' Restituisce il campo privato
            End Get
        End Property

        Public Sub New(logger As Logger)
            Me.logger = logger
        End Sub

        Public Sub RegisterImage(slideIndex As Integer, description As String, left As Integer, top As Integer, width As Integer, height As Integer)
            If _images.Count >= AppConstants.MAX_IMAGES Then
                logger.LogWarning($"Limite massimo immagini raggiunto ({AppConstants.MAX_IMAGES})")
                Exit Sub
            End If

            Dim imageData As New ImageData(slideIndex, description, left, top, width, height)
            _images.Add(imageData)

            logger.LogInfo($"Immagine registrata per slide {slideIndex}")
        End Sub

        Public Sub Clear()
            _images.Clear()
        End Sub

        Public Function HasImages() As Boolean
            Return _images.Count > 0
        End Function

        Public Sub ShowImageGenerationMessage()
            Dim message As String = $"Generazione di {ImageCount} immagini..." & vbCrLf & AppConstants.IMAGES_FUTURE_FEATURE
            MessageBox.Show(message, "Generazione Immagini", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
    End Class
End Namespace