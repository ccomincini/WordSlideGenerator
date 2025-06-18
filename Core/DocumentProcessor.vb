Imports Microsoft.Office.Interop.Word
Namespace WordSlideGenerator

    Public Class DocumentProcessor
        Private logger As Logger
        Private imageManager As ImageManager

        Public Sub New(logger As Logger, imageManager As ImageManager)
            Me.logger = logger
            Me.imageManager = imageManager
        End Sub

        Public Function ProcessDocument(wordDoc As Document) As List(Of SlideContent)
            Dim slides As New List(Of SlideContent)
            Dim currentSlide As SlideContent = Nothing
            Dim currentSection As String = ""

            logger.LogProcess("Inizio elaborazione documento...")

            For i As Integer = 1 To wordDoc.Paragraphs.Count
                Dim paraText As String = CleanParagraphText(wordDoc.Paragraphs(i).Range.Text)

                ' Salta righe vuote
                If paraText = "" Then Continue For

                ' Gestione MODULI DIDATTICI
                If TextRecognizer.RiconosceModulo(paraText) Then
                    ' Salva slide precedente se presente
                    If currentSlide IsNot Nothing Then
                        slides.Add(currentSlide)
                    End If

                    Dim moduleTitle As String = TextRecognizer.EstraiTitoloModulo(paraText)
                    logger.LogInfo($"📁 MODULO rilevato: {moduleTitle}")

                    currentSlide = New SlideContent(moduleTitle, SlideContentType.CourseModule)


                    ' Gestione LEZIONI
                ElseIf TextRecognizer.RiconosceLezione(paraText) Then
                    ' Salva slide precedente se presente
                    If currentSlide IsNot Nothing Then
                        slides.Add(currentSlide)
                    End If

                    Dim lessonTitle As String = TextRecognizer.EstraiTitoloLezione(paraText)
                    logger.LogInfo($"📖 LEZIONE rilevata: {lessonTitle}")

                    currentSlide = New SlideContent(lessonTitle, SlideContentType.Lesson)

                    ' Gestione SLIDE NUMERATE
                ElseIf TextRecognizer.RiconosceSlide(paraText) Then
                    ' Salva slide precedente se presente
                    If currentSlide IsNot Nothing Then
                        slides.Add(currentSlide)
                    End If

                    Dim slideTitle As String = TextRecognizer.EstraiTitolo(paraText)
                    currentSlide = New SlideContent(slideTitle, SlideContentType.Content)
                    currentSection = ""

                    ' Gestione CONTENUTI SLIDE
                ElseIf currentSlide IsNot Nothing Then
                    ProcessSlideContent(currentSlide, paraText, currentSection)
                End If
            Next

            ' Salva ultima slide se presente
            If currentSlide IsNot Nothing Then
                slides.Add(currentSlide)
            End If

            logger.LogSuccess($"Elaborazione completata: {slides.Count} elementi processati")
            Return slides
        End Function

        Private Function CleanParagraphText(text As String) As String
            Dim cleanText As String = Trim(text)

            ' Rimuovi carattere di fine paragrafo
            If cleanText.Length > 0 Then
                If Asc(cleanText.Substring(cleanText.Length - 1)) = 13 Then
                    cleanText = cleanText.Substring(0, cleanText.Length - 1)
                End If
            End If

            Return cleanText
        End Function

        Private Sub ProcessSlideContent(slide As SlideContent, paraText As String, ByRef currentSection As String)
            If TextRecognizer.RiconosceVoceNarrante(paraText) OrElse TextRecognizer.RiconosceTestoNarrazione(paraText) Then
                currentSection = "voce"
                slide.SpeakerNotes = TextRecognizer.EstraiContenutoDopoEtichetta(paraText)

            ElseIf TextRecognizer.RiconosceTestoSlide(paraText) Then
                currentSection = "testo"
                slide.Text = TextRecognizer.EstraiContenutoDopoEtichetta(paraText)

            ElseIf TextRecognizer.RiconosceImmagine(paraText) Then
                currentSection = "immagine"
                slide.ImageDescription = TextRecognizer.EstraiContenutoDopoEtichetta(paraText)

            ElseIf TextRecognizer.RiconosceAppunti(paraText) Then
                currentSection = "appunti"
                slide.Notes = TextRecognizer.EstraiContenutoDopoEtichetta(paraText)

            Else
                ' Continua sezione corrente con contenuto multiriga
                AppendToCurrentSection(slide, paraText, currentSection)
            End If
        End Sub

        Private Sub AppendToCurrentSection(slide As SlideContent, paraText As String, currentSection As String)
            Select Case currentSection
                Case "voce"
                    If slide.SpeakerNotes <> "" Then slide.SpeakerNotes &= vbCrLf
                    slide.SpeakerNotes &= paraText
                Case "testo"
                    If slide.Text <> "" Then slide.Text &= vbCrLf
                    slide.Text &= paraText
                Case "immagine"
                    If slide.ImageDescription <> "" Then slide.ImageDescription &= vbCrLf
                    slide.ImageDescription &= paraText
                Case "appunti"
                    If slide.Notes <> "" Then slide.Notes &= vbCrLf
                    slide.Notes &= paraText
            End Select
        End Sub
    End Class
End Namespace
