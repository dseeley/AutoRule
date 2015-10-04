Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Outlook.OlRuleType

Public Class ThisAddIn

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New RibbonContext1(Me)
    End Function



    Public Sub GoChangeAutoRule(ByVal control As Office.IRibbonControl)
        Me.ChangeAutoRule_EvtHandler()
    End Sub
    Public Sub GoCreateAutoRule(ByVal control As Office.IRibbonControl)
        Me.CreateAutoRule_EvtHandler()
    End Sub
    Public Sub GoRuleifyInbox(ByVal control As Office.IRibbonControl)
        Me.RuleifyInbox()
    End Sub
    Public Sub GoCheckAutoRuleExpiry(ByVal control As Office.IRibbonControl)
        Me.CheckAutoRuleExpiry_EvtHandler()
    End Sub
    Public Sub GoFindInSubfolders(ByVal control As Office.IRibbonControl)
        Me.FindInSubFolders()
    End Sub




    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub


    Dim UserFormWait As New UserFormWait
    Dim oUserformQueue As UserformQueue


    'Go through all inbox items.  For any item that is new, and/or which has a subject that another item has: offer to create new autorule
    Private Sub RuleifyInbox()
        Dim objNamespace As Microsoft.Office.Interop.Outlook.NameSpace = Application.GetNamespace("MAPI")
        Dim folderInbox As Outlook.MAPIFolder = objNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)

        Dim oMail As Object 'Outlook.MailItem '

        Dim FullInboxDict As New Dictionary(Of String, Object)
        Dim ToRuleifyInboxDict As New Dictionary(Of String, Object)

        'Sort the folder descending (most recent first), so that the first email with the matching subject is the most recent
        Dim SortedFolderItems As Items
        SortedFolderItems = folderInbox.Items
        SortedFolderItems.Sort("[ReceivedTime]", True)

        For Each oMail In SortedFolderItems
            Dim SubjectStr As String = CleanSubject(oMail.Subject.ToString)

            'only add to the ToRuleifyInboxDict if this subject already exists
            If ((FullInboxDict.ContainsKey(SubjectStr) = True) Or (oMail.UnRead = True)) Then
                If (ToRuleifyInboxDict.ContainsKey(SubjectStr) = False) Then
                    ToRuleifyInboxDict.Add(SubjectStr, oMail)
                End If
            End If

            'Cope with the element already existing
            If (FullInboxDict.ContainsKey(SubjectStr) = False) Then
                FullInboxDict.Add(SubjectStr, oMail)
            End If
        Next oMail

        If (ToRuleifyInboxDict.Count > 0) Then
            Dim CurrRules As Outlook.Rules = GetCurrRules()
            Dim RulesLessExpiredRules As Outlook.Rules = GetRulesLessExpiredRules(CurrRules)
            Dim ToRuleifyInboxCounter = 1

            For Each InboxItem In ToRuleifyInboxDict.Values
                'Windows.Forms.MessageBox.Show(CleanSubject(InboxItem.Subject.ToString))
                Dim oUserFormPreview As New UserFormPreview(InboxItem.SenderName.ToString, InboxItem.Subject.ToString, InboxItem.Body, ToRuleifyInboxDict.Count, ToRuleifyInboxCounter)

                If oUserFormPreview.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    CreateAutoRule(InboxItem, RulesLessExpiredRules, True)
                End If

                oUserFormPreview.Dispose()
                ToRuleifyInboxCounter += 1
            Next
        Else
            System.Windows.Forms.MessageBox.Show("No new or unread emails in Inbox")
        End If

    End Sub


    Private Sub CreateAutoRule_EvtHandler()
        Dim CurrRules As Outlook.Rules

        'Get the CurrRules, but only if a QueueForm isn't already active.
        Dim QueueFormActive As Boolean = False
        If (Not oUserformQueue Is Nothing) Then
            If (oUserformQueue.IsDisposed = False) Then
                QueueFormActive = True
            End If
        End If
        If (QueueFormActive = False) Then
            CurrRules = GetCurrRules()
        Else
            CurrRules = oUserformQueue.getCurrRules
        End If


        Dim UniqueNewInboxItemsDict As New Dictionary(Of String, Object)
        Dim UniqueExistingInboxItemsDict As New Dictionary(Of String, Object)

        For Each oMail In Application.ActiveExplorer().Selection
            Dim SubjectStr As String = CleanSubject(oMail.Subject.ToString)

            'only add unique items to the UniqueNewInboxItemsDict
            If (UniqueNewInboxItemsDict.ContainsKey(SubjectStr) = False) Then

                'Check if this rule already exists - if so, offer to change it
                Dim oRuleName As String
                Dim Rulefound As Boolean = False
                oRuleName = "<AutoRule> " & SubjectStr

                For Each oRule In CurrRules
                    If (StrComp(oRule.Name, oRuleName, 0) = 0) Then
                        Rulefound = True
                        Exit For
                    End If
                Next

                If Rulefound = True Then
                    UniqueExistingInboxItemsDict.Add(SubjectStr, oMail)
                Else
                    UniqueNewInboxItemsDict.Add(SubjectStr, oMail)
                End If

            End If
        Next oMail


        'For each selected item
        Dim UniqueInboxItemCounter = 1

        For Each UniqueInboxItem In UniqueNewInboxItemsDict.Values
            Dim oUserFormPreview As New UserFormPreview(UniqueInboxItem.SenderName.ToString, UniqueInboxItem.Subject.ToString, UniqueInboxItem.Body, UniqueNewInboxItemsDict.Count, UniqueInboxItemCounter)

            If oUserFormPreview.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                If (Application.ActiveExplorer().Selection.Count > 1) Then
                    CreateAutoRule(UniqueInboxItem, CurrRules, True)
                Else
                    CreateAutoRule(UniqueInboxItem, CurrRules, False)
                End If
            End If

            oUserFormPreview.Dispose()
            UniqueInboxItemCounter += 1
        Next
    End Sub


    'Creates a rule that moves the email on receipt to the specified folder
    Private Sub CreateAutoRule(ByVal objItem As Object, ByVal CurrRules As Outlook.Rules, ByVal QueueMandatory As Boolean)
        Dim objNamespace As Microsoft.Office.Interop.Outlook.NameSpace = Application.GetNamespace("MAPI")
        Dim oRule As Outlook.Rule
        Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
        'Dim oSubjectCondition As Outlook.TextRuleCondition
        Dim oMoveTarget As Outlook.Folder
        Dim oRuleName As String

        Try
            'Give the rule an easy-to-find name
            oRuleName = "<AutoRule> " & CleanSubject(objItem.Subject)

            oMoveTarget = objNamespace.PickFolder
            If Not oMoveTarget Is Nothing Then
                Dim QueueFormActive As Boolean = False

                If (Not oUserformQueue Is Nothing) Then
                    If (oUserformQueue.IsDisposed = False) Then
                        QueueFormActive = True
                    End If
                End If

                If (QueueFormActive = False) Then
                    If (QueueMandatory = True) Then
                        oUserformQueue = New UserformQueue(CurrRules)
                        oUserformQueue.Show()
                        QueueFormActive = True
                    Else
                        If MsgBox("Add to queue (or add directly)?", vbYesNo) = MsgBoxResult.Yes Then
                            oUserformQueue = New UserformQueue(CurrRules)
                            oUserformQueue.Show()
                            QueueFormActive = True
                        End If
                    End If
                End If

                If (QueueFormActive = True) Then
                    oUserformQueue.Focus()
                    oUserformQueue.addItem(objItem.Subject, oMoveTarget)
                Else
                    'Show a "Please wait" dialog
                    UserFormWait.Label1.Text = "Syncing rules with server..."
                    UserFormWait.Show()
                    UserFormWait.Refresh()

                    Dim SubjectArr() As String = {CleanSubject(objItem.Subject)}

                    'Create the rule by adding a Receive Rule to Rules collection
                    oRule = CurrRules.Create(oRuleName, olRuleReceive)

                    'Add rule to the end (by default created at the start), so that the 'Display Alert' rule runs
                    oRule.ExecutionOrder = CurrRules.Count

                    oRule.Conditions.Subject.Enabled = True
                    oRule.Conditions.Subject.Text = SubjectArr

                    'Show the desktop alert  - NO: this creates a client-only rule.
                    'oRule.Actions.DesktopAlert.Enabled = True

                    'oRule.Actions.NewItemAlert.Text = "New AutoRule email"
                    'oRule.Actions.NewItemAlert.Enabled = True

                    'Stop processing more rules - NO: this creates a client-only rule.
                    'oRule.Actions.Stop.Enabled = True

                    oMoveRuleAction = oRule.Actions.MoveToFolder
                    With oMoveRuleAction
                        .Enabled = True
                        .Folder = oMoveTarget
                    End With

                    'Update the server
                    CurrRules.Save()

                    'Hide the "Please wait" dialog
                    UserFormWait.Hide()

                    'Run the rule (it must be saved in the Rules collection first, so we have to then find it again)
                    For Each oRule In CurrRules
                        If oRule.Name = oRuleName Then
                            oRule.Execute(False, objItem.Parent)     ' Need to specify folder, as this sub can be called from ChangeAutoRule
                        End If
                    Next
                End If
            End If

        Catch excep As System.Exception
            System.Windows.Forms.MessageBox.Show(excep.Message & "Error")
        End Try

        'Hide the "Please wait" dialog
        UserFormWait.Hide()

    End Sub


    'Finds an existing MoveTo rule, and changes the MoveTo folder
    Private Sub ChangeAutoRule_EvtHandler()
        Dim objNamespace As Microsoft.Office.Interop.Outlook.NameSpace
        Dim objItem As Object      ' Can be a MailItem, TaskItem or AppointmentItem

        Dim CurrRules As Outlook.Rules = GetCurrRules()     'Get Rules from Session.DefaultStore object
        Dim oRule As Outlook.Rule = Nothing
        Dim Rulefound As Boolean = False
        Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
        Dim oMoveTarget As Outlook.Folder
        Dim MsgAnswer As MsgBoxResult

        Dim oRuleName As String

        Try
            For Each objItem In Application.ActiveExplorer().Selection

                objNamespace = Application.GetNamespace("MAPI")

                'Give the rule an easy-to-find name
                oRuleName = "<AutoRule> " & CleanSubject(objItem.Subject)

                'Find the existing rule (it must be saved in the Rules collection first, so we have to then find it again)
                For Each oRule In CurrRules
                    If (StrComp(oRule.Name, oRuleName, 0) = 0) Then
                        Rulefound = True
                        Exit For
                    End If
                Next

                'Hide the "Please wait" dialog
                UserFormWait.Hide()

                'If the rule is found
                If (Rulefound = True) Then

                    oMoveTarget = objNamespace.PickFolder
                    If Not oMoveTarget Is Nothing Then

                        'Show the "Please wait" dialog
                        UserFormWait.Label1.Text = "Syncing rules with server..."
                        UserFormWait.Show()
                        UserFormWait.Refresh()


                        oMoveRuleAction = oRule.Actions.MoveToFolder
                        With oMoveRuleAction
                            .Enabled = True
                            .Folder = oMoveTarget
                        End With

                        'Update the server
                        CurrRules.Save()

                        'Run the rule on the current folder (it must be saved in the Rules collection first, so we have to then find it again)
                        For Each oRule In CurrRules
                            If oRule.Name = oRuleName Then
                                oRule.Execute(False, objItem.Parent)
                            End If
                        Next

                        'Hide the "Please wait" dialog
                        UserFormWait.Hide()

                    End If

                Else    'Rule is not found, create new?
                    MsgAnswer = MsgBox("No existing rule found.  Create new rule?", vbYesNo)

                    If MsgAnswer = vbYes Then
                        CreateAutoRule(objItem, CurrRules, False)
                    End If

                End If
            Next

        Catch excep As System.Exception
            System.Windows.Forms.MessageBox.Show(excep.Message & "Error")
        End Try

        'Hide the "Please wait" dialog
        UserFormWait.Hide()

    End Sub


    Private Sub CheckAutoRuleExpiry_EvtHandler()
        Dim CurrRules As Outlook.Rules = GetCurrRules()     'this seems to return a reference.
        Dim CurrRulesCount As Integer = CurrRules.Count
        Dim RulesLessExpiredRules As Outlook.Rules = GetRulesLessExpiredRules(CurrRules)

        'if rules haven't changed during the call to GetRulesLessExpiredRules
        If (Not CurrRulesCount = RulesLessExpiredRules.Count) Then
            If (Not RulesLessExpiredRules Is Nothing) Then
                UserFormWait.Label1.Text = "Updating rules with server..."
                UserFormWait.Show()
                UserFormWait.Refresh()

                CurrRules.Save()

                UserFormWait.Hide()
            End If
        End If

        UserFormWait.Hide()
    End Sub


    'Find old rules, delete them
    Private Function GetRulesLessExpiredRules(ByVal AllRules As Outlook.Rules) As Outlook.Rules
        Dim objItem As Object      ' Can be a MailItem, TaskItem or AppointmentItem
        Dim oRule As Outlook.Rule
        Dim oCheckFolder As Outlook.Folder
        Dim RulesChanged As Boolean = False
        Dim CheckDate As Date
        Dim UserFormDel As New UserFormDelete
        Dim RulesLessExpiredRules As Outlook.Rules = AllRules

        ' The user will be offered the chance to delete any rules where
        ' an email that was caught by this rule was received before this date.
        CheckDate = DateAdd(DateInterval.Weekday, -8, Now)  'DateInterval.Week

        Try
            'Show the "Please wait" dialog
            UserFormWait.Label1.Text = "Checking rule expiry..."
            UserFormWait.Show()
            UserFormWait.Refresh()

            Dim ExpiredItemsCount As Integer = 0

            'Find AutoRules
            For Each oRule In AllRules
                If InStr(oRule.Name, "AutoRule") <> 0 Then
                    oCheckFolder = oRule.Actions.MoveToFolder.Folder

                    'Sort the folder descending (most recent first), so that the first email with the matching subject is the most recent,
                    ' and we don't have to iterate the entire folder, matching all the subject names, and searching for the most recent manually.
                    Dim SortedFolderItems As Items
                    SortedFolderItems = oCheckFolder.Items
                    SortedFolderItems.Sort("[ReceivedTime]", True)

                    For Each objItem In SortedFolderItems
                        If (InStr(oRule.Name, CleanSubject(objItem.Subject))) Then
                            If (objItem.ReceivedTime < CheckDate) Then
                                ExpiredItemsCount += 1
                                UserFormDel.addItem(oRule.Name, objItem)
                            End If
                            Exit For  'Since we've sorted the folder, the first instance we find will be the most recent
                        End If
                    Next

                End If

            Next

            If ExpiredItemsCount > 0 Then
                If UserFormDel.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                    Dim ItemsToDel As System.Windows.Forms.ListView.CheckedListViewItemCollection = UserFormDel.GetItemList

                    If (ItemsToDel.Count > 0) Then
                        Dim DelItem As System.Windows.Forms.ListViewItem

                        For Each DelItem In ItemsToDel
                            'Windows.Forms.MessageBox.Show("DelItem:" & DelItem.Text)
                            RulesLessExpiredRules.Remove(DelItem.Text)
                        Next

                    End If
                End If
            End If

        Catch excep As System.Exception
            System.Windows.Forms.MessageBox.Show(excep.Message & "Error")
        End Try


        UserFormWait.Hide()
        UserFormDel.Dispose()

        'Return the new set of rules
        GetRulesLessExpiredRules = RulesLessExpiredRules
    End Function

    Private Sub FindInSubFolders()
        Dim objNamespace As Microsoft.Office.Interop.Outlook.NameSpace
        Dim objItem As Object      ' Can be a MailItem, TaskItem or AppointmentItem
        Dim oFolder As Outlook.Folder

        'Show the "Please wait" dialog
        UserFormWait.Label1.Text = "Searching..."
        UserFormWait.Show()
        UserFormWait.Refresh()

        Try
            objItem = Application.ActiveExplorer().Selection.Item(1)
            objNamespace = Application.GetNamespace("MAPI")

            oFolder = objNamespace.PickFolder
            If Not oFolder Is Nothing Then
                FindInSubFolders_SearchRecurse(oFolder, objItem.Subject)      'Don't put brackets around this, or it sends the oFolder name as a string!?
            End If


        Catch excep As System.Exception
            System.Windows.Forms.MessageBox.Show(excep.Message & "Error")
        End Try

        UserFormWait.Hide()

    End Sub


    Private Sub FindInSubFolders_SearchRecurse(ByVal oParent As Outlook.Folder, ByVal SubjectToFind As String)
        Dim oFolder As Outlook.Folder
        Dim oMail As Object
        Dim MsgBoxAns As MsgBoxResult

        For Each oMail In oParent.Items
            If (StrComp(CleanSubject(SubjectToFind), CleanSubject(oMail.Subject), 0) = 0) Then
                MsgBoxAns = MsgBox("Found " & SubjectToFind & vbNewLine & "(in " & oParent.Name & ")" & vbNewLine & vbNewLine & "Keep searching?", vbYesNo)
                If MsgBoxAns = vbNo Then
                    Exit Sub
                End If
                Exit For
            End If
        Next

        If (oParent.Folders.Count > 0) Then
            For Each oFolder In oParent.Folders
                FindInSubFolders_SearchRecurse(oFolder, SubjectToFind)
            Next
        End If
    End Sub


    'Strips "RE:", "FW:" and "Automatic reply:" from the subject
    Public Shared Function CleanSubject(ByVal oldStr As String) As String
        Dim NewStr As String

        NewStr = oldStr
        If Left(oldStr, 3).ToUpper = "RE:" Then
            NewStr = Trim(Mid(oldStr, 4, Len(oldStr) - 3))
        End If

        If Left(oldStr, 3).ToUpper = "FW:" Then
            NewStr = Trim(Mid(oldStr, 4, Len(oldStr) - 3))
        End If

        If Left(oldStr, 16).ToUpper = "AUTOMATIC REPLY:" Then
            NewStr = Trim(Mid(oldStr, 17, Len(oldStr) - 16))
        End If

        'Subjects can only be 256 characters long
        NewStr = Left(NewStr, 255 - Len("<AutoRule> "))

        CleanSubject = NewStr
    End Function

    Private Function GetCurrRules() As Outlook.Rules
        Dim CurrRules As Outlook.Rules

        'Show the "Please wait" dialog
        UserFormWait.Label1.Text = "Getting rules from server..."
        UserFormWait.Show()
        UserFormWait.Refresh()

        CurrRules = Application.Session.DefaultStore.GetRules()     'this seems to return a reference.

        UserFormWait.Hide()
        Return CurrRules
    End Function


    Private Function GetMessageClass(ByVal item As Object) As String
        Dim args As Object() = New Object() {}
        Dim t As Type = item.GetType()
        Return t.InvokeMember("messageClass",
            Reflection.BindingFlags.Public _
            Or Reflection.BindingFlags.GetField _
            Or Reflection.BindingFlags.GetProperty,
            Nothing, item, args).ToString()
    End Function

End Class
