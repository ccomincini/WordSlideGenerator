Imports System.Text.RegularExpressions
Imports WordSlideGenerator.WordSlideGenerator

Public Class TextCleaner
    Private ReadOnly _logger As Logger

    Public Sub New(logger As Logger)
        _logger = logger
    End Sub

    ''' <summary>
    ''' Pulisce completamente il testo rimuovendo bullet points, numerazioni e normalizzando la formattazione
    ''' </summary>
    Public Function PulisciTestoCompleto(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then
            Return String.Empty
        End If

        Try
            _logger.LogInfo($"[TextCleaner] Input originale: {testo.Substring(0, Math.Min(100, testo.Length))}...")

            Dim testoPulito As String = testo

            ' FASE 1: Rimozione iterativa dei bullet points
            testoPulito = RimuoviBulletPointsAvanzati(testoPulito)

            ' FASE 2: Rimozione numerazioni (1., a), i., etc.)
            testoPulito = RimuoviNumerazioni(testoPulito)

            ' FASE 3: Pulizia spazi e caratteri speciali
            testoPulito = NormalizzaSpazi(testoPulito)

            ' FASE 4: Gestione interruzioni di riga
            testoPulito = NormalizzaInterruzioniRiga(testoPulito)

            ' FASE 5: Rimozione righe vuote multiple
            testoPulito = RimuoviRigheVuoteMultiple(testoPulito)

            _logger.LogSuccess($"[TextCleaner] Output pulito: {testoPulito.Substring(0, Math.Min(100, testoPulito.Length))}...")

            Return testoPulito.Trim()

        Catch ex As Exception
            _logger.LogError($"[TextCleaner] Errore durante pulizia testo: {ex.Message}", ex)
            Return testo ' Ritorna il testo originale in caso di errore
        End Try
    End Function

    ''' <summary>
    ''' Pulizia specifica per placeholder immagini - più aggressiva
    ''' </summary>
    Public Function PulisciTestoPerPlaceholder(testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then
            Return "Immagine descrittiva"
        End If

        Try
            _logger.LogInfo($"[TextCleaner] Pulizia placeholder per: {testo}")

            Dim testoPulito As String = testo

            ' Pulizia extra-aggressiva per placeholder
            testoPulito = RimuoviBulletPointsAvanzati(testoPulito)
            testoPulito = RimuoviNumerazioni(testoPulito)
            testoPulito = RimuoviParoleChiave(testoPulito)
            testoPulito = NormalizzaSpazi(testoPulito)

            ' Se dopo la pulizia il testo è troppo corto o vuoto, usa un placeholder generico
            If testoPulito.Trim().Length < 5 Then
                testoPulito = "Immagine descrittiva per slide"
            End If

            ' Limita lunghezza per placeholder
            If testoPulito.Length > 150 Then
                testoPulito = testoPulito.Substring(0, 147) + "..."
            End If

            _logger.LogSuccess($"[TextCleaner] Placeholder pulito: {testoPulito}")
            Return testoPulito.Trim()

        Catch ex As Exception
            _logger.LogError($"[TextCleaner] Errore pulizia placeholder: {ex.Message}", ex)
            Return "Immagine descrittiva"
        End Try
    End Function

    ''' <summary>
    ''' Rimozione avanzata e iterativa di tutti i tipi di bullet points
    ''' </summary>
    Private Function RimuoviBulletPointsAvanzati(testo As String) As String
        Dim testoPulito As String = testo
        Dim iterazioni As Integer = 0
        Const maxIterazioni As Integer = 5

        ' Pattern di bullet points da rimuovere (ordinati per priorità)
        Dim bulletPatterns As String() = {
            "^[\s]*[•▪▫‣⁃◦∙⦾⦿]+[\s]*", ' Bullet Unicode standard
            "^[\s]*[-\*\+o]+[\s]*",     ' Bullet ASCII classici
            "^[\s]*[→⇒➤➜➔]+[\s]*",     ' Frecce
            "^[\s]*[✓✔️☑️]+[\s]*",      ' Checkmarks
            "^[\s]*[◆◇◊]+[\s]*",       ' Diamanti
            "^[\s]*[■□▲▼]+[\s]*"       ' Forme geometriche
        }

        Do While iterazioni < maxIterazioni
            Dim testoPreIterazione As String = testoPulito

            ' Applica tutti i pattern di rimozione bullet
            For Each pattern As String In bulletPatterns
                testoPulito = Regex.Replace(testoPulito, pattern, "", RegexOptions.Multiline)
            Next

            ' Se non ci sono stati cambiamenti, esci dal loop
            If testoPreIterazione = testoPulito Then
                Exit Do
            End If

            iterazioni += 1
        Loop

        _logger.LogInfo($"[TextCleaner] Rimossi bullet points in {iterazioni} iterazioni")
        Return testoPulito
    End Function

    ''' <summary>
    ''' Rimozione numerazioni (1., a), i., I., A., ecc.)
    ''' </summary>
    Private Function RimuoviNumerazioni(testo As String) As String
        Dim testoPulito As String = testo

        ' Pattern per numerazioni varie
        Dim numerationPatterns As String() = {
            "^[\s]*\d+\.[\s]*",           ' 1. 2. 3.
            "^[\s]*\d+\)[\s]*",           ' 1) 2) 3)
            "^[\s]*\(\d+\)[\s]*",         ' (1) (2) (3)
            "^[\s]*[a-z]\.[\s]*",         ' a. b. c.
            "^[\s]*[a-z]\)[\s]*",         ' a) b) c)
            "^[\s]*[A-Z]\.[\s]*",         ' A. B. C.
            "^[\s]*[A-Z]\)[\s]*",         ' A) B) C)
            "^[\s]*[ivxlc]+\.[\s]*",      ' i. ii. iii. (numeri romani minuscoli)
            "^[\s]*[IVXLC]+\.[\s]*",      ' I. II. III. (numeri romani maiuscoli)
            "^[\s]*[ivxlc]+\)[\s]*",      ' i) ii) iii)
            "^[\s]*[IVXLC]+\)[\s]*"       ' I) II) III)
        }

        For Each pattern As String In numerationPatterns
            testoPulito = Regex.Replace(testoPulito, pattern, "", RegexOptions.Multiline)
        Next

        Return testoPulito
    End Function

    ''' <summary>
    ''' Rimozione parole chiave comuni nei documenti strutturati
    ''' </summary>
    Private Function RimuoviParoleChiave(testo As String) As String
        Dim testoPulito As String = testo

        ' Pattern per rimuovere parole chiave comuni
        Dim parolechePatterns As String() = {
            "^[\s]*contenuto della slide:[\s]*",
            "^[\s]*testo della slide:[\s]*",
            "^[\s]*testo slide:[\s]*",
            "^[\s]*bullet point:[\s]*",
            "^[\s]*punti elenco:[\s]*",
            "^[\s]*slide content:[\s]*",
            "^[\s]*content:[\s]*",
            "^[\s]*main points:[\s]*",
            "^[\s]*immagine:[\s]*",
            "^[\s]*image:[\s]*",
            "^[\s]*visual:[\s]*",
            "^[\s]*immagine consigliata:[\s]*",
            "^[\s]*immagine suggerita:[\s]*"
        }

        For Each pattern As String In parolechePatterns
            testoPulito = Regex.Replace(testoPulito, pattern, "", RegexOptions.Multiline Or RegexOptions.IgnoreCase)
        Next

        Return testoPulito
    End Function

    ''' <summary>
    ''' Normalizzazione spazi, tab e caratteri speciali
    ''' </summary>
    Private Function NormalizzaSpazi(testo As String) As String
        Dim testoPulito As String = testo

        ' Rimuovi caratteri di controllo ASCII (tranne newline)
        testoPulito = Regex.Replace(testoPulito, "[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "")

        ' Normalizza spazi multipli in uno solo
        testoPulito = Regex.Replace(testoPulito, "[ \t]+", " ")

        ' Rimuovi spazi all'inizio e fine di ogni riga
        testoPulito = Regex.Replace(testoPulito, "^[ \t]+|[ \t]+$", "", RegexOptions.Multiline)

        Return testoPulito
    End Function

    ''' <summary>
    ''' Normalizzazione interruzioni di riga
    ''' </summary>
    Private Function NormalizzaInterruzioniRiga(testo As String) As String
        Dim testoPulito As String = testo

        ' Normalizza diverse tipologie di interruzioni riga
        testoPulito = testoPulito.Replace(vbCrLf, vbLf) ' Windows -> Unix
        testoPulito = testoPulito.Replace(vbCr, vbLf)   ' Mac -> Unix

        ' Preserva interruzioni singole ma rimuovi quelle multiple eccessive
        testoPulito = Regex.Replace(testoPulito, "\n{3,}", vbLf & vbLf)

        Return testoPulito
    End Function

    ''' <summary>
    ''' Rimozione righe vuote multiple consecutive
    ''' </summary>
    Private Function RimuoviRigheVuoteMultiple(testo As String) As String
        ' Rimuovi più righe vuote consecutive lasciando massimo una riga vuota
        Return Regex.Replace(testo, "(\r?\n\s*){3,}", vbCrLf & vbCrLf)
    End Function

    ''' <summary>
    ''' Valida che il testo sia stato pulito correttamente
    ''' </summary>
    Public Function ValidaPuliziaTesto(testo As String) As Boolean
        If String.IsNullOrWhiteSpace(testo) Then
            Return True ' Testo vuoto è considerato valido
        End If

        ' Pattern che NON dovrebbero essere presenti nel testo pulito
        Dim invalidPatterns As String() = {
            "[•▪▫‣⁃◦∙]",              ' Bullet points Unicode
            "^[\s]*[-\*\+o][\s]",     ' Bullet ASCII all'inizio riga
            "^[\s]*\d+\.[\s]",        ' Numerazioni
            "^[\s]*[a-zA-Z]\.[\s]",   ' Lettere con punto
            "contenuto della slide:", ' Parole chiave residue
            "testo della slide:"
        }

        For Each pattern As String In invalidPatterns
            If Regex.IsMatch(testo, pattern, RegexOptions.Multiline Or RegexOptions.IgnoreCase) Then
                _logger.LogWarning($"[TextCleaner] Validazione fallita: trovato pattern '{pattern}' nel testo pulito")
                Return False
            End If
        Next

        Return True
    End Function

End Class