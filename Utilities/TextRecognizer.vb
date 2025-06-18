Namespace WordSlideGenerator
    Public Class TextRecognizer

        Public Shared Function RiconosceModulo(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (InStr(testoLower, "modulo") > 0 AndAlso InStr(testoLower, "didattico") > 0) OrElse
                   (testoLower.Length > 0 AndAlso testoLower.Substring(0, 1) = "#" AndAlso InStr(testoLower, "modulo") > 0)
        End Function

        Public Shared Function RiconosceLezione(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (testoLower.Length > 1 AndAlso testoLower.Substring(0, 2) = "##" AndAlso InStr(testoLower, "lezione") > 0) OrElse
                   (InStr(testoLower, "lezione") = 1 AndAlso InStr(testoLower, ":") > 0)
        End Function

        Public Shared Function RiconosceSlide(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            If testoLower.Length > 6 AndAlso testoLower.Substring(0, 6) = "slide " AndAlso testo.Length > 7 Then
                If IsNumeric(testo.Substring(6, 1)) Then
                    Return True
                End If
            End If
            Return False
        End Function

        Public Shared Function RiconosceVoceNarrante(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (InStr(testoLower, "voce narrante") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "testo voce narrante") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "speaker notes") > 0 AndAlso InStr(testoLower, ":") > 0)
        End Function

        Public Shared Function RiconosceTestoNarrazione(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (InStr(testoLower, "testo narrazione") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "narrazione") > 0 AndAlso InStr(testoLower, ":") > 0)
        End Function

        Public Shared Function RiconosceTestoSlide(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (InStr(testoLower, "testo della slide") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "testo slide") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "contenuto slide") > 0 AndAlso InStr(testoLower, ":") > 0)
        End Function

        Public Shared Function RiconosceImmagine(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (InStr(testoLower, "immagine consigliata") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "immagine suggerita") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "immagini suggerite") > 0 AndAlso InStr(testoLower, ":") > 0)
        End Function

        Public Shared Function RiconosceAppunti(testo As String) As Boolean
            Dim testoLower As String = LCase(testo)
            Return (InStr(testoLower, "appunti") > 0 AndAlso InStr(testoLower, ":") > 0) OrElse
                   (InStr(testoLower, "note aggiuntive") > 0 AndAlso InStr(testoLower, ":") > 0)
        End Function

        Public Shared Function EstraiTitoloModulo(testo As String) As String
            Dim risultato As String = testo

            ' Rimuovi marcatori Markdown
            If risultato.Length > 0 AndAlso risultato.Substring(0, 1) = "#" Then
                risultato = Trim(risultato.Substring(1))
                Do While risultato.Length > 0 AndAlso risultato.Substring(0, 1) = "#"
                    risultato = Trim(risultato.Substring(1))
                Loop
            End If

            ' Rimuovi prefisso "modulo didattico:"
            If InStr(LCase(risultato), "modulo didattico:") > 0 Then
                risultato = Trim(risultato.Substring(InStr(LCase(risultato), "modulo didattico:") + 17))
            End If

            Return risultato
        End Function

        Public Shared Function EstraiTitoloLezione(testo As String) As String
            Dim risultato As String = testo

            ' Rimuovi marcatori Markdown "##"
            If risultato.Length > 1 AndAlso risultato.Substring(0, 2) = "##" Then
                risultato = Trim(risultato.Substring(2))
            End If

            Return risultato
        End Function

        Public Shared Function EstraiSoloTitoloLezione(titoloCompleto As String) As String
            Dim risultato As String = titoloCompleto
            Dim posColon As Integer = InStr(risultato, ":")

            If posColon > 0 AndAlso posColon < risultato.Length Then
                risultato = Trim(risultato.Substring(posColon))
            End If

            Return risultato
        End Function

        Public Shared Function EstraiTitolo(riga As String) As String
            Dim posColon As Integer = InStr(riga, ":")

            If posColon > 0 Then
                Return Trim(riga.Substring(posColon))
            Else
                Return riga
            End If
        End Function

        Public Shared Function EstraiContenutoDopoEtichetta(testo As String) As String
            Dim posDuePunti As Integer = InStr(testo, ":")

            If posDuePunti > 0 AndAlso posDuePunti < testo.Length Then
                Dim risultato As String = Trim(testo.Substring(posDuePunti))

                ' Rimuovi virgolette se presenti
                If risultato.Length > 0 AndAlso risultato.Substring(0, 1) = """" Then
                    risultato = risultato.Substring(1)
                End If
                If risultato.Length > 0 AndAlso risultato.Substring(risultato.Length - 1, 1) = """" Then
                    risultato = risultato.Substring(0, risultato.Length - 1)
                End If

                Return risultato
            Else
                Return ""
            End If
        End Function
    End Class
End Namespace