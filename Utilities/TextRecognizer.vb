Imports System.Text.RegularExpressions
Imports WordSlideGenerator.WordSlideGenerator

Public Class TextRecognizer
    Private ReadOnly _logger As Logger

    Public Sub New(logger As Logger)
        _logger = logger
    End Sub

    ''' <summary>
    ''' Riconosce un modulo didattico con pattern robusti multilingue
    ''' </summary>
    Public Function RiconosceModulo(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        ' Pattern per moduli didattici (italiano + inglese)
        Dim moduloPatterns As String() = {
            "^[\s]*#[\s]*modulo[\s]*didattico[\s]*:?",
            "^[\s]*#[\s]*modulo[\s]*:?",
            "^[\s]*modulo[\s]*didattico[\s]*:?",
            "^[\s]*modulo[\s]*\d*[\s]*:?",
            "^[\s]*#[\s]*module[\s]*:?",
            "^[\s]*module[\s]*\d*[\s]*:?",
            "^[\s]*#[\s]*unit[\s]*\d*[\s]*:?"
        }

        For Each pattern As String In moduloPatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                _logger.LogInfo($"[TextRecognizer] Modulo riconosciuto con pattern: '{pattern}'")
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Riconosce una lezione con pattern robusti multilingue
    ''' </summary>
    Public Function RiconosceLezione(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        ' Pattern per lezioni (italiano + inglese)
        Dim lezionePatterns As String() = {
            "^[\s]*##[\s]*lezione[\s]*\d*",
            "^[\s]*lezione[\s]*\d*[\s]*:?",
            "^[\s]*##[\s]*lesson[\s]*\d*",
            "^[\s]*lesson[\s]*\d*[\s]*:?",
            "^[\s]*##[\s]*lecture[\s]*\d*",
            "^[\s]*lecture[\s]*\d*[\s]*:?"
        }

        For Each pattern As String In lezionePatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                _logger.LogInfo($"[TextRecognizer] Lezione riconosciuta con pattern: '{pattern}'")
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Riconosce una slide con parsing numerico robusto - FIX PRINCIPALE
    ''' </summary>
    Public Function RiconosceSlide(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        Try
            ' Pattern principali per slide numerate
            Dim slidePatterns As String() = {
                "^[\s]*slide[\s]*\d+[\s]*:?",
                "^[\s]*slide[\s]+\d+[\s]*:?",
                "^[\s]*slide[\s]*#\d+[\s]*:?",
                "^[\s]*slide[\s]*n\.?\d+[\s]*:?",
                "^[\s]*slide[\s]*numero[\s]*\d+[\s]*:?",
                "^[\s]*\d+[\.]?[\s]*slide[\s]*:?",
                "^[\s]*slide[\s]*\(\d+\)[\s]*:?"
            }

            For Each pattern As String In slidePatterns
                If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                    _logger.LogSuccess($"[TextRecognizer] Slide riconosciuta con pattern: '{pattern}' per testo: '{testo.Trim()}'")
                    Return True
                End If
            Next

            ' Pattern aggiuntivi per varianti meno comuni
            Dim slidePatternExtra As String() = {
                "^[\s]*diapositiva[\s]*\d+[\s]*:?",
                "^[\s]*pagina[\s]*\d+[\s]*:?",
                "^[\s]*schermata[\s]*\d+[\s]*:?"
            }

            For Each pattern As String In slidePatternExtra
                If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                    _logger.LogSuccess($"[TextRecognizer] Slide riconosciuta (extra) con pattern: '{pattern}'")
                    Return True
                End If
            Next

            _logger.LogInfo($"[TextRecognizer] Slide NON riconosciuta per: '{testo.Trim()}'")
            Return False

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore riconoscimento slide: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Riconosce contenuto di slide - FIX PRINCIPALE per "Contenuto della slide:"
    ''' </summary>
    Public Function RiconosceTestoSlide(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        ' Pattern specifici per contenuto slide (ORDINE IMPORTANTE)
        Dim contenutoPatterns As String() = {
            "^[\s]*contenuto[\s]*della[\s]*slide[\s]*:?",
            "^[\s]*testo[\s]*della[\s]*slide[\s]*:?",
            "^[\s]*testo[\s]*slide[\s]*:?",
            "^[\s]*contenuto[\s]*slide[\s]*:?",
            "^[\s]*bullet[\s]*point[\s]*s?[\s]*:?",
            "^[\s]*punti[\s]*elenco[\s]*:?",
            "^[\s]*slide[\s]*content[\s]*:?",
            "^[\s]*content[\s]*:?",
            "^[\s]*main[\s]*points[\s]*:?",
            "^[\s]*key[\s]*points[\s]*:?",
            "^[\s]*points[\s]*:?",
            "^[\s]*argomenti[\s]*:?",
            "^[\s]*temi[\s]*:?",
            "^[\s]*elementi[\s]*:?"
        }

        For Each pattern As String In contenutoPatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                _logger.LogSuccess($"[TextRecognizer] Contenuto slide riconosciuto con pattern: '{pattern}' per testo: '{testo.Trim()}'")
                Return True
            End If
        Next

        _logger.LogInfo($"[TextRecognizer] Contenuto slide NON riconosciuto per: '{testo.Trim()}'")
        Return False
    End Function

    ''' <summary>
    ''' Riconosce note del relatore/speaker
    ''' </summary>
    Public Function RiconosceNoteRelatore(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        Dim notePatterns As String() = {
            "^[\s]*voce[\s]*narrante[\s]*:?",
            "^[\s]*testo[\s]*voce[\s]*narrante[\s]*:?",
            "^[\s]*speaker[\s]*notes?[\s]*:?",
            "^[\s]*note[\s]*relatore[\s]*:?",
            "^[\s]*notes?[\s]*:?",
            "^[\s]*commentary[\s]*:?",
            "^[\s]*narrazione[\s]*:?",
            "^[\s]*commento[\s]*:?"
        }

        For Each pattern As String In notePatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                _logger.LogInfo($"[TextRecognizer] Note relatore riconosciute con pattern: '{pattern}'")
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Riconosce descrizioni di immagini
    ''' </summary>
    Public Function RiconosceImmagine(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        Dim immaginePatterns As String() = {
            "^[\s]*immagine[\s]*consigliata[\s]*:?",
            "^[\s]*immagine[\s]*suggerita[\s]*:?",
            "^[\s]*immagine[\s]*:?",
            "^[\s]*image[\s]*:?",
            "^[\s]*visual[\s]*:?",
            "^[\s]*grafica[\s]*:?",
            "^[\s]*foto[\s]*:?",
            "^[\s]*figure[\s]*:?",
            "^[\s]*illustration[\s]*:?",
            "^[\s]*picture[\s]*:?",
            "^[\s]*graphic[\s]*:?"
        }

        For Each pattern As String In immaginePatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                _logger.LogInfo($"[TextRecognizer] Immagine riconosciuta con pattern: '{pattern}'")
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Riconosce note aggiuntive generiche
    ''' </summary>
    Public Function RiconosceNoteAggiuntive(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then Return False

        Dim noteAggiuntivePatterns As String() = {
            "^[\s]*appunti[\s]*:?",
            "^[\s]*note[\s]*aggiuntive[\s]*:?",
            "^[\s]*note[\s]*extra[\s]*:?",
            "^[\s]*promemoria[\s]*:?",
            "^[\s]*reminder[\s]*:?",
            "^[\s]*annotazioni[\s]*:?",
            "^[\s]*osservazioni[\s]*:?",
            "^[\s]*additional[\s]*notes?[\s]*:?"
        }

        For Each pattern As String In noteAggiuntivePatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.IgnoreCase) Then
                _logger.LogInfo($"[TextRecognizer] Note aggiuntive riconosciute con pattern: '{pattern}'")
                Return True
            End If
        Next

        Return False
    End Function

    ''' <summary>
    ''' Estrae il titolo di un modulo dal testo
    ''' </summary>
    Public Function EstraiTitoloModulo(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            ' Rimuovi marcatori comuni di modulo
            Dim titoloPattern As String = "^[\s]*#*[\s]*modulo[\s]*didattico[\s]*:?[\s]*|^[\s]*#*[\s]*modulo[\s]*\d*[\s]*:?[\s]*|^[\s]*#*[\s]*module[\s]*\d*[\s]*:?[\s]*"
            Dim titolo As String = Regex.Replace(testo, titoloPattern, "", RegexOptions.IgnoreCase).Trim()

            Return If(String.IsNullOrWhiteSpace(titolo), "Modulo Senza Titolo", titolo)

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione titolo modulo: {ex.Message}")
            Return "Modulo Senza Titolo"
        End Try
    End Function

    ''' <summary>
    ''' Estrae il titolo di una lezione dal testo
    ''' </summary>
    Public Function EstraiTitoloLezione(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            ' Rimuovi marcatori comuni di lezione
            Dim titoloPattern As String = "^[\s]*##*[\s]*lezione[\s]*\d*[\s]*:?[\s]*|^[\s]*##*[\s]*lesson[\s]*\d*[\s]*:?[\s]*|^[\s]*##*[\s]*lecture[\s]*\d*[\s]*:?[\s]*"
            Dim titolo As String = Regex.Replace(testo, titoloPattern, "", RegexOptions.IgnoreCase).Trim()

            Return If(String.IsNullOrWhiteSpace(titolo), "Lezione Senza Titolo", titolo)

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione titolo lezione: {ex.Message}")
            Return "Lezione Senza Titolo"
        End Try
    End Function

    ''' <summary>
    ''' Estrae il titolo di una slide dal testo - CON PARSING NUMERICO ROBUSTO
    ''' </summary>
    Public Function EstraiTitoloSlide(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            ' Pattern per rimuovere marcatori di slide e estrarre il titolo
            Dim patterns As String() = {
                "^[\s]*slide[\s]*\d+[\s]*:[\s]*",
                "^[\s]*slide[\s]*\d+[\s]*",
                "^[\s]*slide[\s]*#\d+[\s]*:?[\s]*",
                "^[\s]*slide[\s]*n\.?\d+[\s]*:?[\s]*",
                "^[\s]*\d+[\.]?[\s]*slide[\s]*:?[\s]*",
                "^[\s]*diapositiva[\s]*\d+[\s]*:?[\s]*",
                "^[\s]*pagina[\s]*\d+[\s]*:?[\s]*"
            }

            For Each pattern As String In patterns
                Dim titolo As String = Regex.Replace(testo, pattern, "", RegexOptions.IgnoreCase).Trim()
                If Not String.IsNullOrWhiteSpace(titolo) And titolo <> testo.Trim() Then
                    _logger.LogSuccess($"[TextRecognizer] Titolo slide estratto: '{titolo}' dal pattern: '{pattern}'")
                    Return titolo
                End If
            Next

            ' Se nessun pattern funziona, ritorna il testo pulito
            Dim titoloGenerico As String = testo.Trim()
            _logger.LogWarning($"[TextRecognizer] Nessun pattern slide riconosciuto, uso testo generico: '{titoloGenerico}'")
            Return If(String.IsNullOrWhiteSpace(titoloGenerico), "Slide Senza Titolo", titoloGenerico)

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione titolo slide: {ex.Message}")
            Return "Slide Senza Titolo"
        End Try
    End Function

    ''' <summary>
    ''' Estrae il contenuto di una slide dal testo - FIX PER "Contenuto della slide:"
    ''' </summary>
    Public Function EstraiContenutoSlide(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            ' Pattern specifici per rimuovere marcatori di contenuto slide
            Dim patterns As String() = {
                "^[\s]*contenuto[\s]*della[\s]*slide[\s]*:[\s]*",
                "^[\s]*testo[\s]*della[\s]*slide[\s]*:[\s]*",
                "^[\s]*testo[\s]*slide[\s]*:[\s]*",
                "^[\s]*contenuto[\s]*slide[\s]*:[\s]*",
                "^[\s]*bullet[\s]*point[\s]*s?[\s]*:[\s]*",
                "^[\s]*punti[\s]*elenco[\s]*:[\s]*",
                "^[\s]*slide[\s]*content[\s]*:[\s]*",
                "^[\s]*content[\s]*:[\s]*",
                "^[\s]*main[\s]*points[\s]*:[\s]*",
                "^[\s]*argomenti[\s]*:[\s]*",
                "^[\s]*temi[\s]*:[\s]*"
            }

            For Each pattern As String In patterns
                Dim contenuto As String = Regex.Replace(testo, pattern, "", RegexOptions.IgnoreCase).Trim()
                If Not String.IsNullOrWhiteSpace(contenuto) And contenuto <> testo.Trim() Then
                    _logger.LogSuccess($"[TextRecognizer] Contenuto estratto: '{contenuto.Substring(0, Math.Min(50, contenuto.Length))}...' dal pattern: '{pattern}'")
                    Return contenuto
                End If
            Next

            ' Se nessun pattern funziona, ritorna il testo originale
            _logger.LogInfo($"[TextRecognizer] Nessun pattern contenuto riconosciuto, uso testo originale")
            Return testo.Trim()

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione contenuto slide: {ex.Message}")
            Return testo.Trim()
        End Try
    End Function

    ''' <summary>
    ''' Estrae note del relatore dal testo
    ''' </summary>
    Public Function EstraiNoteRelatore(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            Dim patterns As String() = {
                "^[\s]*voce[\s]*narrante[\s]*:[\s]*",
                "^[\s]*testo[\s]*voce[\s]*narrante[\s]*:[\s]*",
                "^[\s]*speaker[\s]*notes?[\s]*:[\s]*",
                "^[\s]*note[\s]*relatore[\s]*:[\s]*",
                "^[\s]*notes?[\s]*:[\s]*",
                "^[\s]*commentary[\s]*:[\s]*"
            }

            For Each pattern As String In patterns
                Dim note As String = Regex.Replace(testo, pattern, "", RegexOptions.IgnoreCase).Trim()
                If Not String.IsNullOrWhiteSpace(note) And note <> testo.Trim() Then
                    Return note
                End If
            Next

            Return testo.Trim()

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione note relatore: {ex.Message}")
            Return testo.Trim()
        End Try
    End Function

    ''' <summary>
    ''' Estrae descrizione immagine dal testo
    ''' </summary>
    Public Function EstraiDescrizioneImmagine(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            Dim patterns As String() = {
                "^[\s]*immagine[\s]*consigliata[\s]*:[\s]*",
                "^[\s]*immagine[\s]*suggerita[\s]*:[\s]*",
                "^[\s]*immagine[\s]*:[\s]*",
                "^[\s]*image[\s]*:[\s]*",
                "^[\s]*visual[\s]*:[\s]*",
                "^[\s]*grafica[\s]*:[\s]*",
                "^[\s]*foto[\s]*:[\s]*"
            }

            For Each pattern As String In patterns
                Dim descrizione As String = Regex.Replace(testo, pattern, "", RegexOptions.IgnoreCase).Trim()
                If Not String.IsNullOrWhiteSpace(descrizione) And descrizione <> testo.Trim() Then
                    Return descrizione
                End If
            Next

            Return testo.Trim()

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione descrizione immagine: {ex.Message}")
            Return testo.Trim()
        End Try
    End Function

    ''' <summary>
    ''' Estrae note aggiuntive dal testo
    ''' </summary>
    Public Function EstraiNoteAggiuntive(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""

        Try
            Dim patterns As String() = {
                "^[\s]*appunti[\s]*:[\s]*",
                "^[\s]*note[\s]*aggiuntive[\s]*:[\s]*",
                "^[\s]*note[\s]*extra[\s]*:[\s]*",
                "^[\s]*promemoria[\s]*:[\s]*",
                "^[\s]*annotazioni[\s]*:[\s]*"
            }

            For Each pattern As String In patterns
                Dim note As String = Regex.Replace(testo, pattern, "", RegexOptions.IgnoreCase).Trim()
                If Not String.IsNullOrWhiteSpace(note) And note <> testo.Trim() Then
                    Return note
                End If
            Next

            Return testo.Trim()

        Catch ex As Exception
            _logger.LogError($"[TextRecognizer] Errore estrazione note aggiuntive: {ex.Message}")
            Return testo.Trim()
        End Try
    End Function

End Class