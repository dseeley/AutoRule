Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Outlook.OlObjectClass
Imports Microsoft.Office.Interop.Outlook.OlRuleType

Public Class UserformQueue
    Dim NewRuleFolderArr(100) As Outlook.Folder
    Dim CurrRules As Outlook.Rules

    Public Sub New(ByVal iCurrRules As Outlook.Rules)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        CurrRules = iCurrRules

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.SubmitNewRules()
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub UserformQueue_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ListView1.Columns.Add("Subject")
        ListView1.Columns.Add("Folder to move to")
        For Each elem As ColumnHeader In ListView1.Columns
            elem.Width = -2  'Size to fit header width
        Next
    End Sub

    Public Sub addItem(ByVal newRuleSubject As String, ByVal newRuleMoveToFolder As Outlook.Folder)
        Dim str(2) As String
        Dim itm As ListViewItem
        str(0) = newRuleSubject
        str(1) = newRuleMoveToFolder.Name
        itm = New ListViewItem(str)
        itm.Checked = True
        ListView1.Items.Add(itm)

        NewRuleFolderArr(ListView1.Items.Count - 1) = newRuleMoveToFolder

        For Each elem As ColumnHeader In ListView1.Columns
            elem.Width = -1  'Auto-resize to fit new column width
        Next
    End Sub

    'This is only here for completeness
    Public Function getCurrRules() As Outlook.Rules
        Return Me.CurrRules
    End Function

    Public Sub SubmitNewRules()
        Dim CheckedItemList As System.Windows.Forms.ListView.CheckedListViewItemCollection
        Dim UserFormWait As New UserFormWait

        CheckedItemList = ListView1.CheckedItems

        If CheckedItemList.Count > 0 Then
            'Get Rules from Session.DefaultStore object
            Dim objNamespace = Globals.ThisAddIn.Application.GetNamespace("MAPI")

            'CurrRules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules()

            For Each NewRule In CheckedItemList
                'Give the rule an easy-to-find name
                Dim oRuleName As String = "<AutoRule> " & ThisAddIn.CleanSubject(NewRule.Text)
                Dim SubjectArr() As String = {ThisAddIn.CleanSubject(NewRule.Text)}

                'Create the rule by adding a Receive Rule to Rules collection
                Dim oRule As Outlook.Rule = CurrRules.Create(oRuleName, olRuleReceive)

                oRule.Conditions.Subject.Enabled = True
                oRule.Conditions.Subject.Text = SubjectArr

                'Add rule to the end (by default created at the start), so that the 'Display Alert' rule runs
                oRule.ExecutionOrder = CurrRules.Count

                'Show the desktop alert  - NO: this creates a client-only rule.
                'oRule.Actions.DesktopAlert.Enabled = True

                'oRule.Actions.NewItemAlert.Text = "New AutoRule email"
                'oRule.Actions.NewItemAlert.Enabled = True

                'Stop processing more rules  - NO: this creates a client-only rule.
                'oRule.Actions.Stop.Enabled = True

                Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction = oRule.Actions.MoveToFolder
                With oMoveRuleAction
                    .Enabled = True
                    .Folder = NewRuleFolderArr(NewRule.Index)
                End With
            Next

            'Update the "Please wait" dialog
            UserFormWait.Label1.Text = "Syncing rules with server..."
            UserFormWait.Show()
            UserFormWait.Refresh()

            'Update the server
            CurrRules.Save()

            'Run the rule (it must be saved in the Rules collection first, so we have to then find it again)
            For Each oRule In CurrRules
                For Each NewRule In CheckedItemList
                    Dim oRuleName As String = "<AutoRule> " & ThisAddIn.CleanSubject(NewRule.Text)
                    If oRule.Name = oRuleName Then
                        oRule.Execute(False, Globals.ThisAddIn.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox))
                    End If
                Next
            Next
        End If

        UserFormWait.Dispose()
    End Sub

End Class
