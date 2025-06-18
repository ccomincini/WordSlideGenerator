Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Namespace WordSlideGenerator

    Public Class OfficeManager
        Private wordApp As Microsoft.Office.Interop.Word.Application
        Private pptApp As Microsoft.Office.Interop.PowerPoint.Application
        Private wordDoc As Document
        Private pptPresentation As Presentation
        Private logger As Logger

        Public ReadOnly Property WordDocument As Document
            Get
                Return wordDoc
            End Get
        End Property

        Public ReadOnly Property PowerPointPresentation As Presentation
            Get
                Return pptPresentation
            End Get
        End Property

        Public Sub New(logger As Logger)
            Me.logger = logger
        End Sub

        Public Function InitializeApplications() As Boolean
            Try
                logger.LogInfo("Inizializzazione applicazioni Office...")

                ' Inizializza Word e PowerPoint
                Try
                    wordApp = New Microsoft.Office.Interop.Word.Application()
                    wordApp.Visible = False
                    logger.LogInfo("Word Inizializzato")
                Catch ex As Exception
                    logger.LogError("Errore inizializzazione Word", ex)
                    Return False
                End Try

                Try
                    pptApp = New Microsoft.Office.Interop.PowerPoint.Application()
                    pptApp.Visible = True
                    logger.LogInfo("PowerPoint Inizializzato")
                Catch ex As Exception
                    logger.LogError("Errore inizializzazione PowerPoint", ex)
                    Return False
                End Try

                logger.LogSuccess("Applicazioni Office inizializzate")
                Return True

            Catch ex As Exception
                logger.LogError("Errore inizializzazione Office", ex)
                Return False
            End Try
        End Function


        Public Function OpenWordDocument(filePath As String) As Boolean
            Try
                logger.LogInfo("Apertura documento Word...")
                wordDoc = wordApp.Documents.Open(filePath, [ReadOnly]:=True)
                logger.LogSuccess($"Documento Word aperto: {IO.Path.GetFileName(filePath)}")
                Return True

            Catch ex As Exception
                logger.LogError("Errore apertura documento Word", ex)
                Return False
            End Try
        End Function

        Public Function CreatePowerPointPresentation() As Boolean
            Try
                logger.LogInfo("Creazione nuova presentazione PowerPoint...")

                If pptApp Is Nothing Then
                    logger.LogError("pptApp è Nothing!", New NullReferenceException())
                    Return False
                End If

                logger.LogInfo("pptApp.Presentations.Add() in corso...") ' Nuovo log
                pptPresentation = pptApp.Presentations.Add()

                If pptPresentation Is Nothing Then
                    logger.LogError("pptPresentation è Nothing dopo Presentations.Add()!", New NullReferenceException())
                    Return False
                End If

                ' Rimuovi slide vuota iniziale
                If pptPresentation.Slides.Count > 0 Then
                    pptPresentation.Slides(1).Delete()
                End If

                logger.LogSuccess("Presentazione PowerPoint creata")
                Return True

            Catch ex As Exception
                logger.LogError("Errore creazione presentazione PowerPoint", ex)
                Return False
            End Try
        End Function


        Public Sub CloseWord()
            Try
                If wordDoc IsNot Nothing Then
                    logger.LogInfo("Chiusura documento Word...")
                    wordDoc.Close(False)
                    Marshal.ReleaseComObject(wordDoc)
                    wordDoc = Nothing
                End If

                If wordApp IsNot Nothing Then
                    logger.LogInfo("Chiusura applicazione Word...")
                    wordApp.Quit(False)
                    Marshal.ReleaseComObject(wordApp)
                    wordApp = Nothing
                End If

                logger.LogSuccess("Word chiuso correttamente")

            Catch ex As Exception
                logger.LogError("Errore durante chiusura Word", ex)
            End Try
        End Sub

        Public Sub ReleaseAllResources()
            Try
                If wordDoc IsNot Nothing Then
                    logger.LogInfo("Liberazione risorse Word...")
                    wordDoc.Close(False)
                    Marshal.ReleaseComObject(wordDoc)
                    wordDoc = Nothing
                End If

                If wordApp IsNot Nothing Then
                    wordApp.Quit(False)
                    Marshal.ReleaseComObject(wordApp)
                    wordApp = Nothing
                End If

                If pptPresentation IsNot Nothing Then
                    logger.LogInfo("Liberazione risorse PowerPoint...")
                    Marshal.ReleaseComObject(pptPresentation)
                    pptPresentation = Nothing
                End If

                If pptApp IsNot Nothing Then
                    Marshal.ReleaseComObject(pptApp)
                    pptApp = Nothing
                End If

                ' Forza garbage collection
                GC.Collect()
                GC.WaitForPendingFinalizers()

                logger.LogSuccess("Tutte le risorse Office liberate")

            Catch ex As Exception
                logger.LogError("Errore durante liberazione risorse", ex)
            End Try
        End Sub

        Public Function GetSlideCount() As Integer
            If pptPresentation IsNot Nothing Then
                Return pptPresentation.Slides.Count
            End If
            Return 0
        End Function
    End Class
End Namespace