Imports System.Text.RegularExpressions
Namespace WordSlideGenerator

    Public Class TextCleaner

        Public Shared Function PulisciTestoCompleto(testo As String) As String
            If Trim(testo) = "" Then Return ""

            Dim risultato As String = testo

            ' Normalizza interruzioni di riga (mantieni struttura)
            risultato = Replace(risultato, vbCrLf, vbLf)
            risultato = Replace(risultato, vbCr, vbLf)

            ' Rimuovi solo righe vuote multiple consecutive (lascia singole interruzioni)
            Do While InStr(risultato, vbLf & vbLf & vbLf) > 0
                risultato = Replace(risultato, vbLf & vbLf & vbLf, vbLf & vbLf)
            Loop

            Dim righe() As String = Split(risultato, vbLf)
            Dim righeFinali As New List(Of String)

            For Each riga As String In righe
                Dim rigaPulita As String = Trim(riga)

                ' Mantieni righe vuote singole per la formattazione
                If rigaPulita = "" Then
                    righeFinali.Add("")
                    Continue For
                End If

                ' Rimuovi tutti i tipi di punti elenco
                rigaPulita = RimuoviPuntiElenco(rigaPulita)

                ' Rimuovi spazi multipli
                Do While InStr(rigaPulita, "  ") > 0
                    rigaPulita = Replace(rigaPulita, "  ", " ")
                Loop

                ' Aggiungi la riga pulita
                righeFinali.Add(rigaPulita)
            Next

            ' Ricomponi il testo mantenendo le interruzioni di riga necessarie
            Dim testoFinale As String = String.Join(vbCrLf, righeFinali.ToArray())

            ' Rimuovi solo righe vuote eccessive all'inizio e alla fine
            testoFinale = testoFinale.Trim()

            ' Sostituisci righe vuote multiple con una sola riga vuota
            Do While InStr(testoFinale, vbCrLf & vbCrLf & vbCrLf) > 0
                testoFinale = Replace(testoFinale, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
            Loop

            Return testoFinale
        End Function

        Private Shared Function RimuoviPuntiElenco(rigaPulita As String) As String
            Do While True
                Dim rimozioneEffettuata As Boolean = False

                ' Punti elenco con spazio
                If rigaPulita.StartsWith("• ") Then
                    rigaPulita = Trim(rigaPulita.Substring(2))
                    rimozioneEffettuata = True
                ElseIf rigaPulita.StartsWith("- ") Then
                    rigaPulita = Trim(rigaPulita.Substring(2))
                    rimozioneEffettuata = True
                ElseIf rigaPulita.StartsWith("* ") Then
                    rigaPulita = Trim(rigaPulita.Substring(2))
                    rimozioneEffettuata = True
                ElseIf rigaPulita.StartsWith("o ") Then
                    rigaPulita = Trim(rigaPulita.Substring(2))
                    rimozioneEffettuata = True
                ElseIf rigaPulita.StartsWith("+ ") Then
                    rigaPulita = Trim(rigaPulita.Substring(2))
                    rimozioneEffettuata = True
                    ' Punti elenco senza spazio
                ElseIf rigaPulita.StartsWith("•") Then
                    rigaPulita = Trim(rigaPulita.Substring(1))
                    rimozioneEffettuata = True
                ElseIf rigaPulita.StartsWith("-") AndAlso rigaPulita.Length > 1 AndAlso rigaPulita.Substring(1, 1) <> "-" Then
                    rigaPulita = Trim(rigaPulita.Substring(1))
                    rimozioneEffettuata = True
                ElseIf rigaPulita.StartsWith("*") Then
                    rigaPulita = Trim(rigaPulita.Substring(1))
                    rimozioneEffettuata = True
                    ' Tab e spazi iniziali
                ElseIf rigaPulita.StartsWith(vbTab) Then
                    rigaPulita = Trim(rigaPulita.Substring(1))
                    rimozioneEffettuata = True
                    ' Numerazione (1., 2., a), b), etc.)
                ElseIf Regex.IsMatch(rigaPulita, "^[0-9]+\.\s") Then
                    Dim match = Regex.Match(rigaPulita, "^[0-9]+\.\s")
                    rigaPulita = Trim(rigaPulita.Substring(match.Length))
                    rimozioneEffettuata = True
                ElseIf Regex.IsMatch(rigaPulita, "^[a-zA-Z]\)\s") Then
                    rigaPulita = Trim(rigaPulita.Substring(3))
                    rimozioneEffettuata = True
                End If

                If Not rimozioneEffettuata Then Exit Do
            Loop

            Return rigaPulita
        End Function
    End Class
End Namespace