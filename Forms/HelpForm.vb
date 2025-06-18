Namespace WordSlideGenerator
    Public Class HelpForm
        Inherits Form

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.SuspendLayout()

            ' Configurazione finestra guida
            Me.Text = "Guida - Word to PowerPoint Converter"
            Me.Size = New System.Drawing.Size(650, 700)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.ShowIcon = False

            ' TextBox con la guida
            Dim txtGuida As New TextBox()
            txtGuida.Multiline = True
            txtGuida.ScrollBars = ScrollBars.Vertical
            txtGuida.ReadOnly = True
            txtGuida.Location = New System.Drawing.Point(10, 10)
            txtGuida.Size = New System.Drawing.Size(610, 600)
            txtGuida.Font = New System.Drawing.Font("Segoe UI", 10)
            txtGuida.BackColor = System.Drawing.Color.White

            ' Contenuto della guida
            txtGuida.Text = GetHelpContent()

            ' Pulsante OK
            Dim btnOK As New Button()
            btnOK.Text = "OK"
            btnOK.Location = New System.Drawing.Point(285, 620)
            btnOK.Size = New System.Drawing.Size(80, 30)
            btnOK.FlatStyle = FlatStyle.System
            btnOK.DialogResult = DialogResult.OK

            ' Aggiungi controlli alla finestra
            Me.Controls.Add(txtGuida)
            Me.Controls.Add(btnOK)

            Me.ResumeLayout(False)
        End Sub

        Private Function GetHelpContent() As String
            Return "📖 GUIDA ALL'USO - WORD TO POWERPOINT CONVERTER" & vbCrLf & vbCrLf &
                "Questo strumento converte automaticamente documenti Word strutturati in presentazioni PowerPoint professionali." & vbCrLf & vbCrLf &
                "═══════════════════════════════════════════════════════════════" & vbCrLf & vbCrLf &
                "🔧 PREPARAZIONE DEL DOCUMENTO WORD:" & vbCrLf & vbCrLf &
                "Il documento Word deve essere strutturato con etichette specifiche:" & vbCrLf & vbCrLf &
                "📁 MODULI DIDATTICI:" & vbCrLf &
                "   # Modulo Didattico: Nome del Modulo" & vbCrLf &
                "   Oppure: Modulo Didattico: Nome del Modulo" & vbCrLf & vbCrLf &
                "📖 LEZIONI:" & vbCrLf &
                "   ## Lezione 1: Titolo della Lezione" & vbCrLf &
                "   Oppure: Lezione 1: Titolo della Lezione" & vbCrLf & vbCrLf &
                "📄 SLIDE NUMERATE:" & vbCrLf &
                "   Slide 1: Titolo della slide" & vbCrLf &
                "   Slide 2: Altro titolo" & vbCrLf & vbCrLf &
                "📝 CONTENUTI DELLE SLIDE:" & vbCrLf &
                "   Testo della slide: Il contenuto principale" & vbCrLf &
                "   Voce narrante: Note per il relatore" & vbCrLf &
                "   Immagine suggerita: Descrizione dell'immagine" & vbCrLf &
                "   Appunti: Note aggiuntive per il docente" & vbCrLf & vbCrLf &
                "═══════════════════════════════════════════════════════════════" & vbCrLf & vbCrLf &
                "🎯 PROCEDURA DI UTILIZZO:" & vbCrLf & vbCrLf &
                "1️⃣ SELEZIONE FILE:" & vbCrLf &
                "   • Clicca 'Seleziona File Word'" & vbCrLf &
                "   • Scegli il documento .docx strutturato" & vbCrLf &
                "   • Il nome del file apparirà nell'interfaccia" & vbCrLf & vbCrLf &
                "2️⃣ GENERAZIONE PRESENTAZIONE:" & vbCrLf &
                "   • Clicca 'Genera Presentazione Completa'" & vbCrLf &
                "   • L'applicazione elaborerà il documento" & vbCrLf &
                "   • Verranno create slide strutturate automaticamente" & vbCrLf &
                "   • PowerPoint si aprirà con la presentazione" & vbCrLf & vbCrLf &
                "3️⃣ GESTIONE IMMAGINI (se presenti):" & vbCrLf &
                "   • 'Genera Immagini': Funzionalità futura con AI" & vbCrLf &
                "   • 'Salta Immagini': Mantiene placeholder descrittivi" & vbCrLf & vbCrLf &
                "═══════════════════════════════════════════════════════════════" & vbCrLf & vbCrLf &
                "🔘 DESCRIZIONE PULSANTI:" & vbCrLf & vbCrLf &
                "📂 Seleziona File Word:" & vbCrLf &
                "   Apre la finestra per scegliere il documento da convertire" & vbCrLf & vbCrLf &
                "🚀 Genera Presentazione Completa:" & vbCrLf &
                "   Avvia la conversione completa del documento in slides" & vbCrLf & vbCrLf &
                "🎨 Genera Immagini:" & vbCrLf &
                "   Funzionalità futura per generare immagini con AI" & vbCrLf & vbCrLf &
                "⏭️ Salta Immagini:" & vbCrLf &
                "   Completa la presentazione mantenendo i placeholder" & vbCrLf & vbCrLf &
                "🛑 STOP:" & vbCrLf &
                "   Interrompe il processo e libera tutte le risorse" & vbCrLf & vbCrLf &
                "❓ Aiuto:" & vbCrLf &
                "   Mostra questa guida all'utilizzo" & vbCrLf & vbCrLf &
                "🚪 Chiudi Applicazione:" & vbCrLf &
                "   Chiude l'applicazione liberando automaticamente le risorse" & vbCrLf & vbCrLf &
                "═══════════════════════════════════════════════════════════════" & vbCrLf & vbCrLf &
                "📋 RISULTATO FINALE:" & vbCrLf & vbCrLf &
                "✅ Slide di separazione per ogni modulo" & vbCrLf &
                "✅ Slide di apertura per ogni lezione" & vbCrLf &
                "✅ Slide con layout 'Due contenuti' (testo + immagini)" & vbCrLf &
                "✅ Slide con layout 'Titolo e testo' (solo contenuto)" & vbCrLf &
                "✅ Note del relatore complete" & vbCrLf &
                "✅ Placeholder descrittivi per le immagini" & vbCrLf &
                "✅ Formattazione professionale automatica" & vbCrLf & vbCrLf &
                "═══════════════════════════════════════════════════════════════" & vbCrLf & vbCrLf &
                "⚠️ NOTE IMPORTANTI:" & vbCrLf & vbCrLf &
                "• Assicurati che Microsoft Office sia installato" & vbCrLf &
                "• Il documento Word deve seguire la struttura indicata" & vbCrLf &
                "• PowerPoint rimarrà aperto per modifiche successive" & vbCrLf &
                "• I punti elenco vengono automaticamente rimossi" & vbCrLf &
                "• Le interruzioni di riga vengono normalizzate" & vbCrLf & vbCrLf &
                "Per assistenza tecnica, verificare la formattazione del documento Word."
        End Function
    End Class
End Namespace
