Namespace WordSlideGenerator
    Public Class Logger
        Private txtLog As TextBox

        Public Sub New(logTextBox As TextBox)
            txtLog = logTextBox
        End Sub

        Public Sub LogMessage(message As String)
            If txtLog IsNot Nothing Then
                txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}")
                txtLog.SelectionStart = txtLog.Text.Length
                txtLog.ScrollToCaret()
                Application.DoEvents()
            End If
        End Sub

        Public Sub LogError(message As String, ex As Exception)
            LogMessage($"❌ ERRORE: {message} - {ex.Message}")
        End Sub

        Public Sub LogSuccess(message As String)
            LogMessage($"✅ {message}")
        End Sub

        Public Sub LogInfo(message As String)
            LogMessage($"ℹ️ {message}")
        End Sub

        Public Sub LogWarning(message As String)
            LogMessage($"⚠️ {message}")
        End Sub

        Public Sub LogProcess(message As String)
            LogMessage($"🚀 {message}")
        End Sub
    End Class
End Namespace