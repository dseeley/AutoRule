Imports System.Windows.Forms

Public Class UserFormWait

    ' Don't allow form to be closed - only hidden
    Private Sub UserFormWait_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        e.Cancel = True
        Me.Hide()
    End Sub
End Class
