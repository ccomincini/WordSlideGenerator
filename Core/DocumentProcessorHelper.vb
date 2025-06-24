Imports WordSlideGenerator.WordSlideGenerator

Public Class DocumentProcessorHelper
    Private _logger As Logger
    Private _textRecognizer As TextRecognizer
    Private _textCleaner As TextCleaner

    Public Sub New(logger As Logger)
        _logger = logger
        _textRecognizer = New TextRecognizer(logger)
        _textCleaner = New TextCleaner(logger)
    End Sub

    ''' <summary>
    ''' Pulizia testo avanzata - sostituisce il metodo esistente
    ''' </summary>
    Public Function PulisciTestoCompletoAvanzato(testo As String) As String
        Return _textCleaner.PulisciTestoCompleto(testo)
    End Function

    ''' <summary>
    ''' Riconoscimento slide migliorato
    ''' </summary>
    Public Function RiconosceSlideAvanzato(testoRiga As String) As Boolean
        Return _textRecognizer.RiconosceSlide(testoRiga)
    End Function

    ''' <summary>
    ''' Riconoscimento "Contenuto della slide:" - NUOVO
    ''' </summary>
    Public Function RiconosceContenutoSlide(testoRiga As String) As Boolean
        Return _textRecognizer.RiconosceTestoSlide(testoRiga)
    End Function

    ''' <summary>
    ''' Estrazione contenuto da "Contenuto della slide:" - NUOVO
    ''' </summary>
    Public Function EstraiContenutoSlide(testoRiga As String) As String
        Return _textRecognizer.EstraiContenutoSlide(testoRiga)
    End Function

    ''' <summary>
    ''' Pulizia specifica per placeholder immagini
    ''' </summary>
    Public Function PulisciTestoPerImmagine(descrizioneImmagine As String) As String
        Return _textCleaner.PulisciTestoPerPlaceholder(descrizioneImmagine)
    End Function

End Class