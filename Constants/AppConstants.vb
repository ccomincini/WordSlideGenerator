Imports System.Drawing

Namespace WordSlideGenerator

    Public Class AppConstants
        ' Configurazioni UI
        Public Const APP_TITLE As String = "Word to PowerPoint Converter - Macro VBA Edition"
        Public Const FORM_WIDTH As Integer = 700
        Public Const FORM_HEIGHT As Integer = 600

        ' Configurazioni PowerPoint
        Public Const DEFAULT_FONT_NAME As String = "Calibri"
        Public Const TITLE_FONT_SIZE As Integer = 36
        Public Const CONTENT_FONT_SIZE As Integer = 24
        Public Const CONTENT_FONT_SIZE_TWO_COLUMN As Integer = 20
        Public Const IMAGE_PLACEHOLDER_FONT_SIZE As Integer = 14

        ' Colori
        Public Shared ReadOnly TITLE_COLOR As Integer = Drawing.Color.FromArgb(0, 70, 140).ToArgb()
        Public Shared ReadOnly CONTENT_COLOR As Integer = Drawing.Color.FromArgb(64, 64, 64).ToArgb()
        Public Shared ReadOnly IMAGE_PLACEHOLDER_COLOR As Integer = Drawing.Color.FromArgb(80, 80, 80).ToArgb()
        Public Shared ReadOnly MODULE_BACKGROUND_COLOR As Integer = Drawing.Color.FromArgb(230, 240, 255).ToArgb()
        Public Shared ReadOnly LESSON_BACKGROUND_COLOR As Integer = Drawing.Color.FromArgb(240, 248, 255).ToArgb()

        ' Limiti
        Public Const MAX_IMAGES As Integer = 200

        ' Filtri file
        Public Const WORD_FILE_FILTER As String = "Documenti Word (*.docx;*.doc)|*.docx;*.doc"

        ' Messaggi
        Public Const NO_FILE_SELECTED As String = "Nessun file selezionato"
        Public Const READY_FOR_GENERATION As String = "Pronto per generazione"
        Public Const IMAGES_FUTURE_FEATURE As String = "Funzionalità disponibile in versione futura con integrazione AI."
    End Class

End Namespace
