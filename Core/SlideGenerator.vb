Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports PPTShape = Microsoft.Office.Interop.PowerPoint.Shape

Namespace WordSlideGenerator
    Public Class SlideGenerator
        Private presentation As Presentation
        Private logger As Logger
        Private imageManager As ImageManager
        Private sectionGenerator As SectionGenerator
        Private slideCounter As Integer = 0

        Public Sub New(presentation As Presentation, logger As Logger, imageManager As ImageManager, sectionGenerator As SectionGenerator)
            Me.presentation = presentation
            Me.logger = logger
            Me.imageManager = imageManager
            Me.sectionGenerator = sectionGenerator
        End Sub

        ''' <summary>
        ''' Genera tutte le slide dalla lista di contenuti elaborati
        ''' </summary>
        Public Sub GenerateSlides(slideContents As List(Of SlideContent))
            logger.LogProcess("Inizio generazione slide...")

            Try
                For Each content As SlideContent In slideContents
                    slideCounter += 1

                    Select Case content.ContentType
                        Case SlideContentType.CourseModule
                            CreateModuleSlide(content, slideCounter)

                        Case SlideContentType.Lesson
                            CreateLessonSlide(content, slideCounter)

                        Case SlideContentType.Content
                            CreateContentSlide(content, slideCounter)
                    End Select

                    ' Aggiorna slideIndex nell'ImageManager se la slide ha un'immagine
                    If Not String.IsNullOrWhiteSpace(content.ImageDescription) Then
                        imageManager.UpdateSlideIndex(content.ImageDescription, slideCounter)
                    End If
                Next

                logger.LogSuccess($"Generazione slide completata: {slideCounter} slide create")

            Catch ex As Exception
                logger.LogError("Errore durante la generazione slide", ex)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Crea slide separatrice per modulo didattico
        ''' </summary>
        Private Sub CreateModuleSlide(content As SlideContent, slideIndex As Integer)
            Try
                Dim slide As Slide = presentation.Slides.Add(slideIndex, PpSlideLayout.ppLayoutTitleOnly)

                ' Configura titolo del modulo
                Dim titleShape As PPTShape = slide.Shapes.Title
                With titleShape.TextFrame.TextRange
                    .Text = content.Title
                    .Font.Name = AppConstants.DEFAULT_FONT_NAME
                    .Font.Size = AppConstants.TITLE_FONT_SIZE
                    .Font.Color.RGB = AppConstants.TITLE_COLOR
                    .Font.Bold = MsoTriState.msoTrue
                    .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter
                End With

                ' Applica sfondo al modulo
                ApplyModuleBackground(slide)

                ' Crea sezione PowerPoint per il modulo
                sectionGenerator.CreateModuleSection(content.Title, slideIndex)

                logger.LogInfo($"📁 Slide modulo creata: {content.Title}")

            Catch ex As Exception
                logger.LogError($"Errore creazione slide modulo: {content.Title}", ex)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Crea slide di apertura per lezione
        ''' </summary>
        Private Sub CreateLessonSlide(content As SlideContent, slideIndex As Integer)
            Try
                Dim slide As Slide = presentation.Slides.Add(slideIndex, PpSlideLayout.ppLayoutText)

                ' Configura titolo della lezione
                Dim titleShape As PPTShape = slide.Shapes.Title
                With titleShape.TextFrame.TextRange
                    .Text = content.Title
                    .Font.Name = AppConstants.DEFAULT_FONT_NAME
                    .Font.Size = AppConstants.TITLE_FONT_SIZE
                    .Font.Color.RGB = AppConstants.TITLE_COLOR
                    .Font.Bold = MsoTriState.msoTrue
                End With

                ' Aggiungi contenuto se presente
                If Not String.IsNullOrWhiteSpace(content.Text) Then
                    Dim contentShape As PPTShape = slide.Shapes.Placeholders(2)
                    With contentShape.TextFrame.TextRange
                        .Text = content.Text
                        .Font.Name = AppConstants.DEFAULT_FONT_NAME
                        .Font.Size = AppConstants.CONTENT_FONT_SIZE
                        .Font.Color.RGB = AppConstants.CONTENT_COLOR
                    End With
                End If

                ' Applica sfondo alla lezione
                ApplyLessonBackground(slide)

                ' Aggiungi note del relatore se presenti
                AddSpeakerNotes(slide, content)

                ' Crea sezione PowerPoint per la lezione
                sectionGenerator.CreateLessonSection(content.Title, slideIndex)

                logger.LogInfo($"📖 Slide lezione creata: {content.Title}")

            Catch ex As Exception
                logger.LogError($"Errore creazione slide lezione: {content.Title}", ex)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Crea slide di contenuto normale
        ''' </summary>
        Private Sub CreateContentSlide(content As SlideContent, slideIndex As Integer)
            Try
                Dim layout As PpSlideLayout
                Dim hasImage As Boolean = Not String.IsNullOrWhiteSpace(content.ImageDescription)

                ' Scegli layout in base alla presenza di immagini
                If hasImage Then
                    layout = PpSlideLayout.ppLayoutTwoObjects
                Else
                    layout = PpSlideLayout.ppLayoutText
                End If

                Dim slide As Slide = presentation.Slides.Add(slideIndex, layout)

                ' Configura titolo
                ConfigureSlideTitle(slide, content.Title)

                ' Configura contenuto testuale
                ConfigureSlideContent(slide, content.Text, hasImage)

                ' Crea placeholder immagine se necessario
                If hasImage Then
                    CreateImagePlaceholder(slide, content.ImageDescription, slideIndex)
                End If

                ' Aggiungi note del relatore
                AddSpeakerNotes(slide, content)

                logger.LogInfo($"📄 Slide contenuto creata: {content.Title}")

            Catch ex As Exception
                logger.LogError($"Errore creazione slide contenuto: {content.Title}", ex)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Configura il titolo della slide
        ''' </summary>
        Private Sub ConfigureSlideTitle(slide As Slide, title As String)
            Dim titleShape As PPTShape = slide.Shapes.Title
            With titleShape.TextFrame.TextRange
                .Text = title
                .Font.Name = AppConstants.DEFAULT_FONT_NAME
                .Font.Size = AppConstants.TITLE_FONT_SIZE
                .Font.Color.RGB = AppConstants.TITLE_COLOR
                .Font.Bold = MsoTriState.msoTrue
            End With
        End Sub

        ''' <summary>
        ''' Configura il contenuto testuale della slide
        ''' </summary>
        Private Sub ConfigureSlideContent(slide As Slide, text As String, hasImage As Boolean)
            If String.IsNullOrWhiteSpace(text) Then Return

            Dim contentShape As PPTShape = slide.Shapes.Placeholders(2)
            Dim fontSize As Integer = If(hasImage, AppConstants.CONTENT_FONT_SIZE_TWO_COLUMN, AppConstants.CONTENT_FONT_SIZE)

            With contentShape.TextFrame.TextRange
                .Text = text
                .Font.Name = AppConstants.DEFAULT_FONT_NAME
                .Font.Size = fontSize
                .Font.Color.RGB = AppConstants.CONTENT_COLOR
            End With

            ' Configura il TextFrame per gestire correttamente il wrapping
            With contentShape.TextFrame
                .WordWrap = MsoTriState.msoTrue
                .AutoSize = PpAutoSize.ppAutoSizeShapeToFitText
                .MarginBottom = 10
                .MarginTop = 10
                .MarginLeft = 10
                .MarginRight = 10
            End With
        End Sub

        ''' <summary>
        ''' Crea placeholder descrittivo per l'immagine
        ''' </summary>
        Private Sub CreateImagePlaceholder(slide As Slide, imageDescription As String, slideIndex As Integer)
            Try
                ' Ottieni il placeholder per le immagini (terzo placeholder nel layout TwoObjects)
                Dim imagePlaceholder As PPTShape = slide.Shapes.Placeholders(3)

                ' Crea testo descrittivo per il placeholder
                Dim placeholderText As String = $"[IMMAGINE]" & vbCrLf & vbCrLf &
                                              $"Descrizione:" & vbCrLf &
                                              imageDescription & vbCrLf & vbCrLf &
                                              $"Slide: {slideIndex}"

                ' Configura il placeholder come testo
                With imagePlaceholder.TextFrame.TextRange
                    .Text = placeholderText
                    .Font.Name = AppConstants.DEFAULT_FONT_NAME
                    .Font.Size = AppConstants.IMAGE_PLACEHOLDER_FONT_SIZE
                    .Font.Color.RGB = AppConstants.IMAGE_PLACEHOLDER_COLOR
                    .Font.Italic = MsoTriState.msoTrue
                    .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter
                End With

                ' Configura il TextFrame
                With imagePlaceholder.TextFrame
                    .WordWrap = MsoTriState.msoTrue
                    .AutoSize = PpAutoSize.ppAutoSizeShapeToFitText
                    .VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle
                End With

                logger.LogInfo($"🖼️ Placeholder immagine creato per slide {slideIndex}")

            Catch ex As Exception
                logger.LogError($"Errore creazione placeholder immagine slide {slideIndex}", ex)
            End Try
        End Sub

        ''' <summary>
        ''' Aggiunge note del relatore alla slide
        ''' </summary>
        Private Sub AddSpeakerNotes(slide As Slide, content As SlideContent)
            Dim notes As String = ""

            ' Combina note del relatore e appunti
            If Not String.IsNullOrWhiteSpace(content.SpeakerNotes) Then
                notes = "VOCE NARRANTE:" & vbCrLf & content.SpeakerNotes
            End If

            If Not String.IsNullOrWhiteSpace(content.Notes) Then
                If notes <> "" Then notes &= vbCrLf & vbCrLf
                notes &= "APPUNTI:" & vbCrLf & content.Notes
            End If

            ' Aggiungi descrizione immagine alle note se presente
            If Not String.IsNullOrWhiteSpace(content.ImageDescription) Then
                If notes <> "" Then notes &= vbCrLf & vbCrLf
                notes &= "IMMAGINE SUGGERITA:" & vbCrLf & content.ImageDescription
            End If

            ' Applica le note alla slide
            If notes <> "" Then
                slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes
                logger.LogInfo($"📝 Note aggiunte alla slide: {content.Title}")
            End If
        End Sub

        ''' <summary>
        ''' Applica sfondo specifico per slide modulo
        ''' </summary>
        Private Sub ApplyModuleBackground(slide As Slide)
            Try
                With slide.Background.Fill
                    .Visible = MsoTriState.msoTrue
                    .ForeColor.RGB = AppConstants.MODULE_BACKGROUND_COLOR
                    .Transparency = 0.1
                End With
            Catch ex As Exception
                logger.LogWarning($"Impossibile applicare sfondo modulo: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Applica sfondo specifico per slide lezione
        ''' </summary>
        Private Sub ApplyLessonBackground(slide As Slide)
            Try
                With slide.Background.Fill
                    .Visible = MsoTriState.msoTrue
                    .ForeColor.RGB = AppConstants.LESSON_BACKGROUND_COLOR
                    .Transparency = 0.05
                End With
            Catch ex As Exception
                logger.LogWarning($"Impossibile applicare sfondo lezione: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace