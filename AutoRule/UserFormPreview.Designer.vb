<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserFormPreview
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Yes_Button = New System.Windows.Forms.Button()
        Me.No_Button = New System.Windows.Forms.Button()
        Me.RichTextBoxEmailPreview = New System.Windows.Forms.RichTextBox()
        Me.TextBoxSubject = New System.Windows.Forms.TextBox()
        Me.TextBoxSender = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Yes_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.No_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(664, 627)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(219, 45)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Yes_Button
        '
        Me.Yes_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Yes_Button.Location = New System.Drawing.Point(4, 5)
        Me.Yes_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Yes_Button.Name = "Yes_Button"
        Me.Yes_Button.Size = New System.Drawing.Size(100, 35)
        Me.Yes_Button.TabIndex = 0
        Me.Yes_Button.Text = "Yes"
        '
        'No_Button
        '
        Me.No_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.No_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.No_Button.Location = New System.Drawing.Point(114, 5)
        Me.No_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.No_Button.Name = "No_Button"
        Me.No_Button.Size = New System.Drawing.Size(100, 35)
        Me.No_Button.TabIndex = 1
        Me.No_Button.Text = "No"
        '
        'RichTextBoxEmailPreview
        '
        Me.RichTextBoxEmailPreview.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBoxEmailPreview.Location = New System.Drawing.Point(18, 74)
        Me.RichTextBoxEmailPreview.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RichTextBoxEmailPreview.Name = "RichTextBoxEmailPreview"
        Me.RichTextBoxEmailPreview.ReadOnly = True
        Me.RichTextBoxEmailPreview.Size = New System.Drawing.Size(864, 541)
        Me.RichTextBoxEmailPreview.TabIndex = 1
        Me.RichTextBoxEmailPreview.Text = ""
        '
        'TextBoxSubject
        '
        Me.TextBoxSubject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxSubject.Enabled = False
        Me.TextBoxSubject.Location = New System.Drawing.Point(207, 38)
        Me.TextBoxSubject.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBoxSubject.Name = "TextBoxSubject"
        Me.TextBoxSubject.Size = New System.Drawing.Size(676, 26)
        Me.TextBoxSubject.TabIndex = 2
        '
        'TextBoxSender
        '
        Me.TextBoxSender.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxSender.Enabled = False
        Me.TextBoxSender.Location = New System.Drawing.Point(18, 38)
        Me.TextBoxSender.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TextBoxSender.Name = "TextBoxSender"
        Me.TextBoxSender.Size = New System.Drawing.Size(176, 26)
        Me.TextBoxSender.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Sender"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(207, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 20)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Subject"
        '
        'UserFormPreview
        '
        Me.AcceptButton = Me.Yes_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.No_Button
        Me.ClientSize = New System.Drawing.Size(902, 690)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBoxSender)
        Me.Controls.Add(Me.TextBoxSubject)
        Me.Controls.Add(Me.RichTextBoxEmailPreview)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "UserFormPreview"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add an AutoRule for this email chain?"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Yes_Button As System.Windows.Forms.Button
    Friend WithEvents No_Button As System.Windows.Forms.Button
    Friend WithEvents RichTextBoxEmailPreview As System.Windows.Forms.RichTextBox
    Friend WithEvents TextBoxSubject As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxSender As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
