Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Namespace WordSlideGenerator

    Public Class SlideGenerator
        Private presentation As Presentation
        Private logger As Logger
        Private imageManager As ImageManager
        Private sectionGenerator As SectionGenerator

        Public Sub New(presentation As Presentation, logger As Logger, imageManager As ImageManager, sectionGenerator As SectionGenerator)
            Me.presentation = presentation
            Me.logger = logger
            Me.imageManager = imageManager
            Me.sectionGenerator = sectionGenerator
        End Sub

        Public Sub GenerateSlides(slideContents As List(Of SlideContent))
            Dim slideCount As Integer = 0

            For Each content As SlideContent In slideContents
                Select Case content.SlideType
                    Case SlideContentType.CourseModule
                        CreateModuleSlide(content)
                        slideCount += 1

                    Case SlideContentType.Lesson
                        CreateLessonSlide(content)
                        slideCount += 1

                    Case SlideContentType.Content
                        CreateContentSlide(content)
                        slideCount += 1
                End Select
            Next


            logger.LogSuccess($"Generazione completata: {slideCount} slide create")
            sectionGenerator.CreateSections(presentation, slideContents)
        End Sub

        Private Sub CreateModuleSlide(content As SlideContent)
            Try
                ' Layout "Solo titolo" per slide modulo
                Dim newSlide As Slide = presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutTitleOnly)

                ' Imposta titolo del modulo
                If newSlide.Shapes.HasTitle Then
                    newSlide.Shapes.Title.TextFrame.TextRange.Text = "MODULO: " & content.Title

                    With newSlide.Shapes.Title.TextFrame.TextRange
                        .Font.Name = AppConstants.DEFAULT_FONT_NAME
                        .Font.Size = AppConstants.TITLE_FONT_SIZE
                        .Font.Bold = True
                        .Font.Color.RGB = AppConstants.TITLE_COLOR
                        .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter
                    End With
                End If

                ' Sfondo colorato
                newSlide.Background.Fill.Visible = MsoTriState.msoTrue
                newSlide.Background.Fill.ForeColor.RGB = AppConstants.MODULE_BACKGROUND_COLOR

                logger.LogInfo($"Slide separatrice modulo creata: {content.Title}")

            Catch ex As Exception
                logger.LogError($"Errore creazione slide modulo: {content.Title}", ex)
            End Try
        End Sub


        Private Sub CreateLessonSlide(content As SlideContent)
            Try
                Dim newSlide As Slide = presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutText)

                ' Estrai solo il titolo della lezione (senza "Lezione X:")
                Dim lessonTitleOnly As String = TextRecognizer.EstraiSoloTitoloLezione(content.Title)

                ' Imposta titolo lezione
                If newSlide.Shapes.HasTitle Then
                    newSlide.Shapes.Title.TextFrame.TextRange.Text = lessonTitleOnly

                    With newSlide.Shapes.Title.TextFrame.TextRange
                        .Font.Name = AppConstants.DEFAULT_FONT_NAME
                        .Font.Size = 44
                        .Font.Bold = True
                        .Font.Color.RGB = AppConstants.TITLE_COLOR
                        .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter
                    End With
                End If

                ' Contenuto standard panoramica
                Try
                    With newSlide.Shapes.Placeholders(2).TextFrame.TextRange
                        .Text = "Panoramica della Lezione" & vbCrLf & vbCrLf &
                                "Obiettivi di apprendimento" & vbCrLf &
                                "Argomenti principali" & vbCrLf &
                                "Attività pratiche"
                        .Font.Name = AppConstants.DEFAULT_FONT_NAME
                        .Font.Size = AppConstants.CONTENT_FONT_SIZE
                        .Font.Color.RGB = AppConstants.CONTENT_COLOR
                        .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft
                        .ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse
                    End With
                Catch
                    ' Layout alternativo se placeholder non disponibile
                End Try

                ' Sfondo leggero
                newSlide.Background.Fill.Visible = MsoTriState.msoTrue
                newSlide.Background.Fill.ForeColor.RGB = AppConstants.LESSON_BACKGROUND_COLOR

                ' Note per il docente
                SetSlideNotes(newSlide, $"Slide di apertura per: {content.Title}")

                logger.LogInfo($"Slide apertura lezione creata: {lessonTitleOnly}")

            Catch ex As Exception
                logger.LogError($"Errore creazione slide apertura: {content.Title}", ex)
            End Try
        End Sub

        Private Sub CreateContentSlide(content As SlideContent)
            Try
                Dim newSlide As Slide
                Dim cleanText As String = TextCleaner.PulisciTestoCompleto(content.Text)

                logger.LogInfo($"Creazione slide: {content.Title}")

                If content.HasImage() Then
                    ' Layout "Due contenuti" per slide con immagine
                    newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutTwoObjects)
                    SetupTwoColumnSlide(newSlide, content, cleanText)

                    ' Registra per generazione immagini
                    imageManager.RegisterImage(presentation.Slides.Count, content.ImageDescription, 480, 150, 250, 350)
                Else
                    ' Layout "Titolo e testo" per slide senza immagine
                    newSlide = presentation.Slides.Add(presentation.Slides.Count + 1, PpSlideLayout.ppLayoutText)
                    SetupSingleColumnSlide(newSlide, content, cleanText)
                End If

                ' Imposta note complete nella slide
                Dim completeNotes As String = content.GetCompleteNotes()
                If completeNotes <> "" Then
                    SetSlideNotes(newSlide, TextCleaner.PulisciTestoCompleto(completeNotes))
                End If

                logger.LogSuccess($"Slide creata: {content.Title}")

            Catch ex As Exception
                logger.LogError($"Errore creazione slide: {content.Title}", ex)
            End Try
        End Sub

        Private Sub SetupTwoColumnSlide(slide As Slide, content As SlideContent, cleanText As String)
            ' Imposta titolo
            If content.Title <> "" Then
                slide.Shapes.Title.TextFrame.TextRange.Text = content.Title
                FormatTitle(slide.Shapes.Title.TextFrame.TextRange)
            End If

            ' Contenuto testuale nel primo placeholder (sinistra)
            Try
                With slide.Shapes.Placeholders(2).TextFrame.TextRange
                    .Text = cleanText
                    .Font.Name = AppConstants.DEFAULT_FONT_NAME
                    .Font.Size = AppConstants.CONTENT_FONT_SIZE_TWO_COLUMN
                    .Font.Color.RGB = AppConstants.CONTENT_COLOR
                    .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft
                    .ParagraphFormat.SpaceAfter = 6
                    .ParagraphFormat.SpaceBefore = 0
                    .ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse
                End With

                ' Configura il textframe per il word wrap
                With slide.Shapes.Placeholders(2).TextFrame
                    .WordWrap = MsoTriState.msoTrue
                    .AutoSize = PpAutoSize.ppAutoSizeShapeToFitText
                    .MarginLeft = 20
                    .MarginRight = 20
                    .MarginTop = 20
                    .MarginBottom = 20
                End With
            Catch ex As Exception
                logger.LogError("Errore impostazione contenuto principale", ex)
            End Try

            ' Placeholder immagine nel secondo placeholder (destra)
            Try
                With slide.Shapes.Placeholders(3).TextFrame.TextRange
                    .Text = "📷 IMMAGINE SUGGERITA:" & vbCrLf & vbCrLf & TextCleaner.PulisciTestoCompleto(content.ImageDescription)
                    .Font.Name = AppConstants.DEFAULT_FONT_NAME
                    .Font.Size = AppConstants.IMAGE_PLACEHOLDER_FONT_SIZE
                    .Font.Italic = True
                    .Font.Color.RGB = AppConstants.IMAGE_PLACEHOLDER_COLOR
                    .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft
                    .ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse
                End With

                ' Configura il textframe per il placeholder immagine
                With slide.Shapes.Placeholders(3).TextFrame
                    .WordWrap = MsoTriState.msoTrue
                    .AutoSize = PpAutoSize.ppAutoSizeShapeToFitText
                    .MarginLeft = 15
                    .MarginRight = 15
                    .MarginTop = 15
                    .MarginBottom = 15
                End With

                ' Stile visivo placeholder
                Try
                    With slide.Shapes.Placeholders(3).Line
                        .Visible = MsoTriState.msoTrue
                        .ForeColor.RGB = RGB(200, 200, 200)
                        .Weight = 1
                    End With

                    With slide.Shapes.Placeholders(3).Fill
                        .Visible = MsoTriState.msoTrue
                        .ForeColor.RGB = RGB(248, 248, 248)
                        .Transparency = 0
                    End With
                Catch
                    ' Ignora errori di formattazione se non supportati
                End Try
            Catch ex As Exception
                logger.LogError("Errore impostazione placeholder immagine", ex)
            End Try
        End Sub

        Private Sub SetupSingleColumnSlide(slide As Slide, content As SlideContent, cleanText As String)
            ' Imposta titolo
            If content.Title <> "" Then
                slide.Shapes.Title.TextFrame.TextRange.Text = content.Title
                FormatTitle(slide.Shapes.Title.TextFrame.TextRange)
            End If

            ' Contenuto nel placeholder del testo
            Try
                With slide.Shapes.Placeholders(2).TextFrame.TextRange
                    .Text = cleanText
                    .Font.Name = AppConstants.DEFAULT_FONT_NAME
                    .Font.Size = AppConstants.CONTENT_FONT_SIZE
                    .Font.Color.RGB = AppConstants.CONTENT_COLOR
                    .ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft
                    .ParagraphFormat.SpaceAfter = 8
                    .ParagraphFormat.SpaceBefore = 0
                    .ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse
                End With

                ' Configura il textframe
                With slide.Shapes.Placeholders(2).TextFrame
                    .WordWrap = MsoTriState.msoTrue
                    .AutoSize = PpAutoSize.ppAutoSizeShapeToFitText
                    .MarginLeft = 30
                    .MarginRight = 30
                    .MarginTop = 30
                    .MarginBottom = 30
                End With
            Catch ex As Exception
                logger.LogError("Errore impostazione contenuto slide", ex)
            End Try
        End Sub

        Private Sub FormatTitle(titleRange As TextRange)
            With titleRange
                .Font.Name = AppConstants.DEFAULT_FONT_NAME
                .Font.Size = AppConstants.TITLE_FONT_SIZE
                .Font.Bold = True
                .Font.Color.RGB = AppConstants.TITLE_COLOR
            End With
        End Sub

        Private Sub SetSlideNotes(slide As Slide, notes As String)
            If notes <> "" Then
                Try
                    slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes
                Catch
                    Try
                        slide.NotesPage.Shapes(2).TextFrame.TextRange.Text = notes
                    Catch
                        ' Ignora errori note se non disponibili
                    End Try
                End Try
            End If
        End Sub
    End Class
End Namespace
