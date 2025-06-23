Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports WordSlideGenerator.WordSlideGenerator

' RIMOSSO: Imports WordSlideGenerator.WordSlideGenerator (RIDONDANTE)
Imports PPT = Microsoft.Office.Interop.PowerPoint
Imports WRD = Microsoft.Office.Interop.Word

Public Class Form1
    Inherits Form

    ' Servizi e gestori
    Private logger As Logger
    Private officeManager As OfficeManager
    Private imageManager As ImageManager
    Private documentProcessor As DocumentProcessor
    Private slideGenerator As SlideGenerator

    ' Controlli UI
    Private lblSelectedFile As Label
    Private btnGeneratePresentation As Button
    Private lblStructureStatus As Label
    Private lblProgress As Label
    Private btnGenerateImages As Button
    Private btnSkipImages As Button
    Private btnStopProcess As Button
    Private txtLog As TextBox
    Private btnCloseApp As Button
    Private btnHelp As Button
    Private grpImages As GroupBox

    Public Sub New()
        InitializeComponent()
        InitializeServices()
    End Sub

    Private Sub InitializeServices()
        ' Inizializza logger
        logger = New Logger(txtLog)

        ' Inizializza gestori
        officeManager = New OfficeManager(logger)
        imageManager = New ImageManager(logger)
        documentProcessor = New DocumentProcessor(logger, imageManager)

        slideGenerator = New SlideGenerator(officeManager.PowerPointPresentation, logger, imageManager, New SectionGenerator(logger))

        logger.LogInfo("Applicazione inizializzata")
    End Sub

    Private Sub BtnSelectFile_Click(sender As Object, e As EventArgs)
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = AppConstants.WORD_FILE_FILTER
        openFileDialog.Title = "Seleziona il documento Word strutturato"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            lblSelectedFile.Text = Path.GetFileName(openFileDialog.FileName)
            lblSelectedFile.ForeColor = System.Drawing.Color.Black
            lblSelectedFile.Tag = openFileDialog.FileName

            btnGeneratePresentation.Enabled = True
            logger.LogInfo($"File selezionato: {openFileDialog.FileName}")
        End If
    End Sub

    Private Sub BtnGeneratePresentation_Click(sender As Object, e As EventArgs)
        Try
            ' Prima libera risorse precedenti se esistono
            officeManager.ReleaseAllResources()
            imageManager.Clear()

            Dim filePath As String = lblSelectedFile.Tag.ToString()

            logger.LogProcess("INIZIO ELABORAZIONE DOCUMENTO")

            ' Abilita pulsante STOP
            btnStopProcess.Enabled = True

            ' Inizializza Office
            If Not officeManager.InitializeApplications() Then
                MessageBox.Show("Errore nell'inizializzazione di Office", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Crea presentazione PowerPoint
            If Not officeManager.CreatePowerPointPresentation() Then
                MessageBox.Show("Errore nella creazione della presentazione PowerPoint", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            'Verifica che la presentazione sia stata creata
            If officeManager.PowerPointPresentation Is Nothing Then
                MessageBox.Show("Errore: la presentazione non è stata creata correttamente.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Apri documento Word
            If Not officeManager.OpenWordDocument(filePath) Then
                MessageBox.Show("Errore nell'apertura del documento Word", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Processa documento e genera slide
            logger.LogProcess("Generazione presentazione senza immagini...")
            Dim slideContents As List(Of SlideContent) = documentProcessor.ProcessDocument(officeManager.WordDocument)

            Dim sectionGenerator As New SectionGenerator(logger)

            ' Inizializza generatore slide
            slideGenerator = New SlideGenerator(officeManager.PowerPointPresentation, logger, imageManager, sectionGenerator)
            slideGenerator.GenerateSlides(slideContents)

            ' Aggiorna UI
            UpdateUIAfterGeneration()

        Catch ex As Exception
            logger.LogError("Errore durante la generazione", ex)
            MessageBox.Show($"Errore: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnStopProcess.Enabled = False
        End Try
    End Sub

    Private Sub UpdateUIAfterGeneration()
        lblStructureStatus.Text = $"Presentazione generata - {officeManager.GetSlideCount()} slide create"
        lblStructureStatus.ForeColor = System.Drawing.Color.Green

        ' Disabilita STOP
        btnStopProcess.Enabled = False

        ' Abilita sezione immagini se necessario
        If imageManager.HasImages() Then
            btnGenerateImages.Enabled = True
            btnSkipImages.Enabled = True

            lblProgress.Text = $"Trovate {imageManager.ImageCount} immagini da generare - Scegli un'opzione"
            lblProgress.ForeColor = System.Drawing.Color.DarkBlue

            ' Chiudi Word (mantieni solo PowerPoint aperto)
            officeManager.CloseWord()
        Else
            ' Nessuna immagine - chiudi Word
            officeManager.CloseWord()
            logger.LogSuccess("Presentazione completata - nessuna immagine da generare")
        End If

        logger.LogSuccess($"Presentazione completata con {officeManager.GetSlideCount()} slide")
        logger.LogInfo($"🎨 Trovate {imageManager.ImageCount} immagini da generare")
    End Sub

    Private Sub BtnGenerateImages_Click(sender As Object, e As EventArgs)
        imageManager.ShowImageGenerationMessage()
        officeManager.CloseWord()
    End Sub

    Private Sub BtnSkipImages_Click(sender As Object, e As EventArgs)
        logger.LogInfo("Generazione immagini saltata dall'utente")

        ' Chiudi Word
        officeManager.CloseWord()

        ' Disabilita pulsanti
        btnGenerateImages.Enabled = False
        btnSkipImages.Enabled = False

        ' Aggiorna messaggio di stato
        lblProgress.Text = $"Presentazione completata con {imageManager.ImageCount} placeholder per immagini"
        lblProgress.ForeColor = System.Drawing.Color.Green

        MessageBox.Show($"Presentazione senza immagini completata con successo!" & vbCrLf & vbCrLf &
                        $"• {officeManager.GetSlideCount()} slide create" & vbCrLf &
                        $"• {imageManager.ImageCount} placeholder per immagini" & vbCrLf & vbCrLf &
                        "La presentazione PowerPoint è ora disponibile per l'editing.",
                        "Presentazione Completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnStopProcess_Click(sender As Object, e As EventArgs)
        logger.LogWarning("Processo interrotto dall'utente")
        officeManager.ReleaseAllResources()

        ' Disabilita tutti i pulsanti
        btnStopProcess.Enabled = False
        btnGenerateImages.Enabled = False
        btnSkipImages.Enabled = False

        MessageBox.Show("Processo interrotto. Tutte le risorse sono state liberate.", "Processo Fermato", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub BtnCloseApp_Click(sender As Object, e As EventArgs)
        Dim result As DialogResult = MessageBox.Show("Chiudere l'applicazione?" & vbCrLf & "Tutte le risorse verranno liberate automaticamente.",
                                                     "Conferma Chiusura",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            logger.LogInfo("Chiusura applicazione richiesta dall'utente...")
            officeManager.ReleaseAllResources()
            System.Windows.Forms.Application.Exit()
        End If
    End Sub

    Private Sub BtnHelp_Click(sender As Object, e As EventArgs)
        Dim helpFilePath As String = IO.Path.Combine(System.Windows.Forms.Application.StartupPath, "Help", "Help.html")

        If IO.File.Exists(helpFilePath) Then
            Help.ShowHelp(Me, helpFilePath)
        Else
            MessageBox.Show("File della guida non trovato!", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        logger.LogInfo("Chiusura applicazione...")
        officeManager.ReleaseAllResources()
        MyBase.OnFormClosed(e)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub CreateImageGenerationGroup()
        grpImages = New GroupBox()
        With grpImages
            .Text = "3. Generazione Immagini (Opzionale)"
            .Location = New System.Drawing.Point(20, 260)
            .Size = New System.Drawing.Size(640, 120)
        End With

        Dim progressBar As New ProgressBar()
        With progressBar
            .Location = New System.Drawing.Point(20, 30)
            .Size = New System.Drawing.Size(600, 20)
        End With

        lblProgress = New Label()
        With lblProgress
            .Text = "Immagini verranno gestite dopo la creazione della struttura"
            .Location = New System.Drawing.Point(20, 55)
            .Size = New System.Drawing.Size(600, 20)
        End With

        btnGenerateImages = New Button()
        With btnGenerateImages
            .Text = "Genera Immagini"
            .Location = New System.Drawing.Point(20, 85)
            .Size = New System.Drawing.Size(120, 25)
            .Enabled = False
            .FlatStyle = FlatStyle.System
        End With
        AddHandler btnGenerateImages.Click, AddressOf BtnGenerateImages_Click

        btnSkipImages = New Button()
        With btnSkipImages
            .Text = "Salta Immagini"
            .Location = New System.Drawing.Point(150, 85)
            .Size = New System.Drawing.Size(100, 25)
            .Enabled = False
            .FlatStyle = FlatStyle.System
        End With
        AddHandler btnSkipImages.Click, AddressOf BtnSkipImages_Click

        btnStopProcess = New Button()
        With btnStopProcess
            .Text = "STOP"
            .Location = New System.Drawing.Point(260, 85)
            .Size = New System.Drawing.Size(60, 25)
            .Enabled = False
            .FlatStyle = FlatStyle.System
            .BackColor = System.Drawing.SystemColors.Control
            .ForeColor = System.Drawing.Color.DarkRed
            .Font = New System.Drawing.Font("Segoe UI", 9, FontStyle.Bold)
        End With
        AddHandler btnStopProcess.Click, AddressOf BtnStopProcess_Click

        grpImages.Controls.Add(progressBar)
        grpImages.Controls.Add(lblProgress)
        grpImages.Controls.Add(btnGenerateImages)
        grpImages.Controls.Add(btnSkipImages)
        grpImages.Controls.Add(btnStopProcess)
    End Sub

    Private Sub CreateControlButtons()
        btnCloseApp = New Button()
        With btnCloseApp
            .Text = "Chiudi Applicazione"
            .Location = New System.Drawing.Point(550, 530)
            .Size = New System.Drawing.Size(110, 30)
            .FlatStyle = FlatStyle.System
        End With
        AddHandler btnCloseApp.Click, AddressOf BtnCloseApp_Click

        btnHelp = New Button()
        With btnHelp
            .Text = "❓ Aiuto"
            .Location = New System.Drawing.Point(430, 530)
            .Size = New System.Drawing.Size(110, 30)
            .FlatStyle = FlatStyle.System
        End With
        AddHandler btnHelp.Click, AddressOf BtnHelp_Click
    End Sub

    Private Sub CreateUserInterface()
        ' File selection group
        Dim grpFileSelection As New GroupBox()
        grpFileSelection.Text = "1. Selezione File Word"
        grpFileSelection.Location = New System.Drawing.Point(20, 20)
        grpFileSelection.Size = New System.Drawing.Size(640, 80)

        Dim btnSelectFile As New Button()
        btnSelectFile.Text = "Seleziona File Word"
        btnSelectFile.Location = New System.Drawing.Point(20, 30)
        btnSelectFile.Size = New System.Drawing.Size(150, 30)
        btnSelectFile.FlatStyle = FlatStyle.System
        AddHandler btnSelectFile.Click, AddressOf BtnSelectFile_Click

        lblSelectedFile = New Label()
        lblSelectedFile.Text = AppConstants.NO_FILE_SELECTED
        lblSelectedFile.Location = New System.Drawing.Point(180, 35)
        lblSelectedFile.Size = New System.Drawing.Size(440, 20)
        lblSelectedFile.ForeColor = System.Drawing.Color.Gray

        grpFileSelection.Controls.Add(btnSelectFile)
        grpFileSelection.Controls.Add(lblSelectedFile)

        ' Structure generation group
        Dim grpStructure As New GroupBox()
        grpStructure.Text = "2. Generazione Presentazione Senza Immagini"
        grpStructure.Location = New System.Drawing.Point(20, 120)
        grpStructure.Size = New System.Drawing.Size(640, 120)

        btnGeneratePresentation = New Button()
        btnGeneratePresentation.Text = "Genera Presentazione Senza Immagini"
        btnGeneratePresentation.Location = New System.Drawing.Point(20, 30)
        btnGeneratePresentation.Size = New System.Drawing.Size(220, 35)
        btnGeneratePresentation.Enabled = False
        btnGeneratePresentation.FlatStyle = FlatStyle.System
        AddHandler btnGeneratePresentation.Click, AddressOf BtnGeneratePresentation_Click

        lblStructureStatus = New Label()
        lblStructureStatus.Text = AppConstants.READY_FOR_GENERATION
        lblStructureStatus.Location = New System.Drawing.Point(250, 35)
        lblStructureStatus.Size = New System.Drawing.Size(370, 20)
        lblStructureStatus.ForeColor = System.Drawing.Color.Gray

        Dim lblFeatures As New Label()
        lblFeatures.Text = "• Moduli e Lezioni • Slide strutturate • Note del relatore • Placeholder immagini"
        lblFeatures.Location = New System.Drawing.Point(20, 70)
        lblFeatures.Size = New System.Drawing.Size(600, 20)
        lblFeatures.ForeColor = System.Drawing.Color.DarkGreen
        lblFeatures.Font = New System.Drawing.Font("Calibri", 9, FontStyle.Italic)

        grpStructure.Controls.Add(btnGeneratePresentation)
        grpStructure.Controls.Add(lblStructureStatus)
        grpStructure.Controls.Add(lblFeatures)

        ' Image generation group
        CreateImageGenerationGroup()

        ' Status and log
        txtLog = New TextBox()
        txtLog.Location = New System.Drawing.Point(20, 400)
        txtLog.Size = New System.Drawing.Size(640, 120)
        txtLog.Multiline = True
        txtLog.ScrollBars = ScrollBars.Vertical
        txtLog.ReadOnly = True
        txtLog.BackColor = System.Drawing.Color.White
        txtLog.Font = New System.Drawing.Font("Consolas", 9)

        ' Control buttons
        CreateControlButtons()

        ' Add all controls to form
        Me.Controls.Add(grpFileSelection)
        Me.Controls.Add(grpStructure)
        Me.Controls.Add(grpImages)
        Me.Controls.Add(txtLog)
        Me.Controls.Add(btnCloseApp)
        Me.Controls.Add(btnHelp)
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        ' Form properties
        Me.Text = AppConstants.APP_TITLE
        Me.Size = New System.Drawing.Size(AppConstants.FORM_WIDTH, AppConstants.FORM_HEIGHT)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.ClientSize = New System.Drawing.Size(700, 600)
        Me.Name = "Form1"

        Me.ResumeLayout(False)
        CreateUserInterface()
    End Sub
End Class