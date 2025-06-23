Namespace WordSlideGenerator
    Public Class ImageManager
        Private logger As Logger
        Private images As New List(Of ImageData)

        Public Sub New(logger As Logger)
            Me.logger = logger
        End Sub

        ' METODO ORIGINALE (CORRETTO)
        Public Sub RegisterImage(slideIndex As Integer, description As String, left As Integer, top As Integer, width As Integer, height As Integer)
            ' Implementazione esistente... CORRETTO: Usa PosLeft, PosTop, PosWidth, PosHeight
            Dim imageData As New ImageData With {
                .SlideIndex = slideIndex,
                .Description = description,
                .PosLeft = left,
                .PosTop = top,
                .PosWidth = width,
                .PosHeight = height
            }
            images.Add(imageData)
            logger.LogInfo($"🖼️ Immagine registrata: {description}")
        End Sub

        ' NUOVO OVERLOAD - per DocumentProcessor
        ''' <summary>
        ''' Registra un'immagine con parametri predefiniti per la fase di processing del documento
        ''' </summary>
        ''' <param name="description">Descrizione dell'immagine</param>
        Public Sub RegisterImage(description As String)
            ' Chiama il metodo completo con valori predefiniti
            ' slideIndex = 0 (verrà aggiornato durante la generazione delle slide)
            ' Posizione e dimensioni standard per placeholder
            RegisterImage(0, description, 1, 1, 4, 3)
        End Sub

        ' Metodo per aggiornare lo slideIndex quando viene creata la slide effettiva
        Public Sub UpdateSlideIndex(description As String, newSlideIndex As Integer)
            Dim imageData = images.FirstOrDefault(Function(img) img.Description = description AndAlso img.SlideIndex = 0)
            If imageData IsNot Nothing Then
                imageData.SlideIndex = newSlideIndex
                logger.LogInfo($"🔄 Aggiornato slideIndex per immagine: {description} -> Slide {newSlideIndex}")
            End If
        End Sub

        ' Altri metodi esistenti...
        Public ReadOnly Property ImageCount As Integer
            Get
                Return images.Count
            End Get
        End Property

        Public Function HasImages() As Boolean
            Return images.Count > 0
        End Function

        Public Sub Clear()
            images.Clear()
            logger.LogInfo("🗑️ Cache immagini pulita")
        End Sub

        Public Sub ShowImageGenerationMessage()
            MessageBox.Show(AppConstants.IMAGES_FUTURE_FEATURE, "Generazione Immagini", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub
    End Class
End Namespace