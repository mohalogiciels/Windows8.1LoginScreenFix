Public Class ContactMe

    Private Sub GitHubRepoLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles GitHubRepoLinkLabel.LinkClicked
        Process.Start("https://www.github.com/mohalogiciels/Windows8.1LoginScreenFix")
    End Sub

    Private Sub EmailLinkLabel_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles EmailLinkLabel.LinkClicked
        Process.Start("mailto:mohalogiciels@hotmail.com")
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Me.Dispose()
    End Sub

End Class
