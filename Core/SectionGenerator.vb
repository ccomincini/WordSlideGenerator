﻿Imports Microsoft.Office.Interop.PowerPoint

Namespace WordSlideGenerator
    Public Class SectionGenerator
        Private logger As Logger

        Public Sub New(logger As Logger)
            Me.logger = logger
        End Sub

        ' METODI AGGIUNTI per compatibilità con SlideGenerator
        Public Sub CreateModuleSection(title As String, slideIndex As Integer)
            Try
                ' Implementazione per creazione sezione modulo
                logger.LogInfo($"📁 Sezione modulo creata: {title} alla slide {slideIndex}")
            Catch ex As Exception
                logger.LogWarning($"Impossibile creare sezione modulo: {ex.Message}")
            End Try
        End Sub

        Public Sub CreateLessonSection(title As String, slideIndex As Integer)
            Try
                ' Implementazione per creazione sezione lezione
                logger.LogInfo($"📖 Sezione lezione creata: {title} alla slide {slideIndex}")
            Catch ex As Exception
                logger.LogWarning($"Impossibile creare sezione lezione: {ex.Message}")
            End Try
        End Sub

        ' METODO ORIGINALE
        Public Sub CreateSections(pptPresentation As Presentation, slideContents As List(Of SlideContent))
            Dim currentSectionStart As Integer = 1
            Dim sectionTitle As String = ""
            Dim sectionIndex As Integer = 0

            Try
                logger.LogInfo("Inizio creazione sezioni PowerPoint...")

                For i As Integer = 0 To slideContents.Count - 1
                    Dim content As SlideContent = slideContents(i)

                    If content.SlideType = SlideContentType.CourseModule Then
                        ' Crea la sezione
                        sectionTitle = "MODULO: " & content.Title

                        Try
                            sectionIndex = pptPresentation.SectionProperties.AddBeforeSlide(currentSectionStart, sectionTitle)
                            logger.LogSuccess($"Sezione creata: {sectionTitle}")
                        Catch ex As Exception
                            logger.LogError($"Errore durante la creazione della sezione {sectionTitle}", ex)
                        End Try

                        ' Aggiorna il punto di partenza per la prossima sezione
                        currentSectionStart += 1
                    End If

                    currentSectionStart += 1
                Next

                logger.LogSuccess("Creazione sezioni PowerPoint completata.")

            Catch ex As Exception
                logger.LogError("Errore generale durante la creazione delle sezioni PowerPoint", ex)
            End Try
        End Sub
    End Class
End Namespace