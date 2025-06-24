Imports Microsoft.Office.Interop.Word
Imports WordSlideGenerator.WordSlideGenerator

Public Class DocumentProcessor
    Private ReadOnly _logger As Logger
    Private ReadOnly _imageManager As ImageManager

    ' Costruttore originale compatibile con Form1
    Public Sub New(logger As Logger, imageManager As ImageManager)
        _logger = logger
        _imageManager = imageManager
    End Sub

    ''' <summary>
    ''' Elabora un documento Word e estrae i contenuti delle slide - METODO PRINCIPALE
    ''' </summary>
    Public Function ProcessDocument(documento As Document) As List(Of SlideContent)
        Dim slideContents As New List(Of SlideContent)
        Dim slideCorrente As SlideContent = Nothing
        Dim sezioneCorrente As String = "Generale"

        ' Inizializza helper per pattern recognition avanzato
        Dim textRecognizer As New TextRecognizer(_logger)
        Dim textCleaner As New TextCleaner(_logger)

        Try
            _logger.LogInfo("[DocumentProcessor] Inizio elaborazione documento")
            _logger.LogInfo($"[DocumentProcessor] Numero paragrafi: {documento.Paragraphs.Count}")

            For Each paragrafo As Paragraph In documento.Paragraphs
                Dim testoRiga As String = paragrafo.Range.Text
                If String.IsNullOrWhiteSpace(testoRiga) Then Continue For

                ' Debug logging - testo originale
                _logger.LogInfo($"[DocumentProcessor] Elaborando riga: '{testoRiga.Trim()}'")

                ' FASE 1: Riconoscimento pattern e classificazione contenuto
                Dim tipoContenuto As String = ClassificaContenuto(testoRiga, textRecognizer)
                _logger.LogInfo($"[DocumentProcessor] Tipo contenuto riconosciuto: {tipoContenuto}")

                ' FASE 2: Elaborazione basata sul tipo di contenuto
                Select Case tipoContenuto
                    Case "MODULO"
                        ' Nuovo modulo didattico
                        Dim titoloModulo As String = textRecognizer.EstraiTitoloModulo(testoRiga)
                        titoloModulo = textCleaner.PulisciTestoCompleto(titoloModulo)

                        slideCorrente = CreaSlideModulo(titoloModulo)
                        slideContents.Add(slideCorrente)
                        sezioneCorrente = titoloModulo
                        _logger.LogSuccess($"[DocumentProcessor] Creato modulo: '{titoloModulo}'")

                    Case "LEZIONE"
                        ' Nuova lezione
                        Dim titoloLezione As String = textRecognizer.EstraiTitoloLezione(testoRiga)
                        titoloLezione = textCleaner.PulisciTestoCompleto(titoloLezione)

                        slideCorrente = CreaSlideLezione(titoloLezione, sezioneCorrente)
                        slideContents.Add(slideCorrente)
                        _logger.LogSuccess($"[DocumentProcessor] Creata lezione: '{titoloLezione}'")

                    Case "SLIDE"
                        ' Nuova slide di contenuto
                        Dim titoloSlide As String = textRecognizer.EstraiTitoloSlide(testoRiga)
                        titoloSlide = textCleaner.PulisciTestoCompleto(titoloSlide)

                        slideCorrente = CreaSlideContenuto(titoloSlide, sezioneCorrente)
                        slideContents.Add(slideCorrente)
                        _logger.LogSuccess($"[DocumentProcessor] Creata slide: '{titoloSlide}'")

                    Case "CONTENUTO_SLIDE"
                        ' 🎯 CONTENUTO PRINCIPALE DELLA SLIDE - FIX PER "Contenuto della slide:"
                        If slideCorrente IsNot Nothing Then
                            Dim contenutoOriginale As String = textRecognizer.EstraiContenutoSlide(testoRiga)
                            _logger.LogInfo($"[DocumentProcessor] Contenuto PRIMA pulizia: '{contenutoOriginale}'")

                            Dim contenutoPulito As String = textCleaner.PulisciTestoCompleto(contenutoOriginale)
                            _logger.LogSuccess($"[DocumentProcessor] Contenuto DOPO pulizia: '{contenutoPulito}'")

                            ' Validazione pulizia
                            If Not textCleaner.ValidaPuliziaTesto(contenutoPulito) Then
                                _logger.LogWarning($"[DocumentProcessor] Validazione pulizia fallita per: '{contenutoPulito}'")
                            End If

                            slideCorrente.Text = AggiungiTestoSlide(slideCorrente.Text, contenutoPulito)
                        Else
                            _logger.LogWarning($"[DocumentProcessor] Contenuto slide trovato senza slide corrente: '{testoRiga}'")
                        End If

                    Case "NOTE_SPEAKER"
                        ' Note per il relatore
                        If slideCorrente IsNot Nothing Then
                            Dim noteOriginali As String = textRecognizer.EstraiNoteRelatore(testoRiga)
                            Dim notePulite As String = textCleaner.PulisciTestoCompleto(noteOriginali)

                            slideCorrente.SpeakerNotes = AggiungiTestoSlide(slideCorrente.SpeakerNotes, notePulite)
                            _logger.LogInfo($"[DocumentProcessor] Aggiunte note speaker: '{notePulite}'")
                        End If

                    Case "IMMAGINE"
                        ' Descrizione immagine
                        If slideCorrente IsNot Nothing Then
                            Dim descrizioneOriginale As String = textRecognizer.EstraiDescrizioneImmagine(testoRiga)
                            Dim descrizionePulita As String = textCleaner.PulisciTestoPerPlaceholder(descrizioneOriginale)

                            slideCorrente.ImageDescription = descrizionePulita
                            _logger.LogInfo($"[DocumentProcessor] Immagine: '{descrizionePulita}'")

                            ' Registra immagine nell'ImageManager
                            _imageManager.RegisterImage(descrizionePulita)
                        End If

                    Case "NOTE_AGGIUNTIVE"
                        ' Note aggiuntive generiche
                        If slideCorrente IsNot Nothing Then
                            Dim noteOriginali As String = textRecognizer.EstraiNoteAggiuntive(testoRiga)
                            Dim notePulite As String = textCleaner.PulisciTestoCompleto(noteOriginali)

                            slideCorrente.Notes = AggiungiTestoSlide(slideCorrente.Notes, notePulite)
                            _logger.LogInfo($"[DocumentProcessor] Note aggiuntive: '{notePulite}'")
                        End If

                    Case Else
                        ' Contenuto non classificato - aggiunge alla slide corrente se esiste
                        If slideCorrente IsNot Nothing Then
                            Dim testoGenerico As String = textCleaner.PulisciTestoCompleto(testoRiga)

                            If Not String.IsNullOrWhiteSpace(testoGenerico) Then
                                slideCorrente.Text = AggiungiTestoSlide(slideCorrente.Text, testoGenerico)
                                _logger.LogInfo($"[DocumentProcessor] Aggiunto contenuto generico: '{testoGenerico}'")
                            End If
                        Else
                            _logger.LogWarning($"[DocumentProcessor] Contenuto non classificato ignorato: '{testoRiga}'")
                        End If
                End Select
            Next

            _logger.LogSuccess($"[DocumentProcessor] Elaborazione completata. {slideContents.Count} slide create")
            Return slideContents

        Catch ex As Exception
            _logger.LogError($"[DocumentProcessor] Errore durante elaborazione: {ex.Message}", ex)
            Return slideContents
        End Try
    End Function

    ''' <summary>
    ''' Classifica il tipo di contenuto di una riga di testo
    ''' </summary>
    Private Function ClassificaContenuto(testoRiga As String, textRecognizer As TextRecognizer) As String
        Try
            ' Test in ordine di priorità (più specifico al meno specifico)

            If textRecognizer.RiconosceModulo(testoRiga) Then
                Return "MODULO"
            End If

            If textRecognizer.RiconosceLezione(testoRiga) Then
                Return "LEZIONE"
            End If

            If textRecognizer.RiconosceSlide(testoRiga) Then
                Return "SLIDE"
            End If

            ' 🎯 PRIORITÀ ALTA per "Contenuto della slide:"
            If textRecognizer.RiconosceTestoSlide(testoRiga) Then
                Return "CONTENUTO_SLIDE"
            End If

            If textRecognizer.RiconosceNoteRelatore(testoRiga) Then
                Return "NOTE_SPEAKER"
            End If

            If textRecognizer.RiconosceImmagine(testoRiga) Then
                Return "IMMAGINE"
            End If

            If textRecognizer.RiconosceNoteAggiuntive(testoRiga) Then
                Return "NOTE_AGGIUNTIVE"
            End If

            Return "GENERICO"

        Catch ex As Exception
            _logger.LogError($"[DocumentProcessor] Errore classificazione contenuto: {ex.Message}", ex)
            Return "GENERICO"
        End Try
    End Function

    ''' <summary>
    ''' Crea una nuova slide per un modulo didattico
    ''' </summary>
    Private Function CreaSlideModulo(titolo As String) As SlideContent
        Return New SlideContent With {
            .Title = titolo,
            .Text = "",
            .SpeakerNotes = $"Introduzione al modulo: {titolo}",
            .ImageDescription = "",
            .Notes = "",
            .SlideType = SlideContentType.CourseModule,
            .ContentType = SlideContentType.CourseModule
        }
    End Function

    ''' <summary>
    ''' Crea una nuova slide per una lezione
    ''' </summary>
    Private Function CreaSlideLezione(titolo As String, modulo As String) As SlideContent
        Return New SlideContent With {
            .Title = titolo,
            .Text = "",
            .SpeakerNotes = $"Apertura lezione: {titolo} (Modulo: {modulo})",
            .ImageDescription = "",
            .Notes = "",
            .SlideType = SlideContentType.Lesson,
            .ContentType = SlideContentType.Lesson
        }
    End Function

    ''' <summary>
    ''' Crea una nuova slide di contenuto
    ''' </summary>
    Private Function CreaSlideContenuto(titolo As String, sezione As String) As SlideContent
        Return New SlideContent With {
            .Title = titolo,
            .Text = "",
            .SpeakerNotes = "",
            .ImageDescription = "",
            .Notes = "",
            .SlideType = SlideContentType.Content,
            .ContentType = SlideContentType.Content
        }
    End Function

    ''' <summary>
    ''' Aggiunge testo a un campo esistente gestendo correttamente le interruzioni di riga
    ''' </summary>
    Private Function AggiungiTestoSlide(testoEsistente As String, nuovoTesto As String) As String
        If String.IsNullOrWhiteSpace(nuovoTesto) Then
            Return testoEsistente
        End If

        If String.IsNullOrWhiteSpace(testoEsistente) Then
            Return nuovoTesto
        End If

        ' Aggiungi con interruzione di riga appropriata
        Return testoEsistente.TrimEnd() & vbCrLf & nuovoTesto.TrimStart()
    End Function

End Class