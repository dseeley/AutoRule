Imports System.Windows.Forms

Public Class UserFormDelete
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub UserFormDelete_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ListView1.Columns.Add("Rule Name")
        ListView1.Columns.Add("Last used date")
        For Each elem As ColumnHeader In ListView1.Columns
            elem.Width = -2  'Size to fit header column width
        Next
    End Sub


    Public Sub addItem(ByVal newRule As String, ByVal newObj As Object)
        Dim str(2) As String
        Dim itm As ListViewItem
        str(0) = newRule
        str(1) = newObj.ReceivedTime
        itm = New ListViewItem(str)
        itm.Checked = True
        ListView1.Items.Add(itm)

        For Each elem As ColumnHeader In ListView1.Columns
            elem.Width = -1  'Auto-resize to fit new column width
        Next
    End Sub

    Public Function GetItemList() As System.Windows.Forms.ListView.CheckedListViewItemCollection
        GetItemList = ListView1.CheckedItems
    End Function

    Private Sub ButtonAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAll.Click
        For Each elem As ListViewItem In ListView1.Items
            elem.Checked = True
        Next
    End Sub

    Private Sub ButtonNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNone.Click
        For Each elem As ListViewItem In ListView1.Items
            elem.Checked = False
        Next
    End Sub
End Class
