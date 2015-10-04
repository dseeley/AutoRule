Imports System.Windows.Forms

Public Class UserFormPreview
    Dim myEmailSender As String
    Dim myEmailSubject As String
    Dim myEmailPreview As String
    Dim myEmailTotal As String
    Dim myEmailCount As String

    Public Sub New(ByVal iEmailSender As String, ByVal iEmailSubject As String, ByVal iEmailPreview As String, ByVal iEmailTotal As Integer, ByVal iEmailCount As Integer)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        myEmailSender = iEmailSender
        myEmailSubject = iEmailSubject
        myEmailPreview = iEmailPreview
        myEmailTotal = iEmailTotal
        myEmailCount = iEmailCount
    End Sub

    Private Sub UserFormPreview_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        TextBoxSender.AppendText(Me.myEmailSender)
        TextBoxSubject.AppendText(Me.myEmailSubject)
        RichTextBoxEmailPreview.AppendText(Me.myEmailPreview)
        Me.Text = Me.Text + " (" + Me.myEmailCount + " of " + Me.myEmailTotal + ")"
    End Sub

    Private Sub Yes_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Yes_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub No_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles No_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
