
<Runtime.InteropServices.ComVisible(True)> _
Public Class RibbonContext1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Dim parentObj As Object

    Public Sub New(myParent As Object)
        Me.parentObj = myParent
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("AutoRule.RibbonContext1.xml")
    End Function


#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub


    Public Sub CreateAutoRule(ByVal control As Office.IRibbonControl)
        parentObj.GoCreateAutoRule(control)
    End Sub
    Public Sub ChangeAutoRule(ByVal control As Office.IRibbonControl)
        parentObj.GoChangeAutoRule(control)
    End Sub
    Public Sub RuleifyInbox(ByVal control As Office.IRibbonControl)
        parentObj.GoRuleifyInbox(control)
    End Sub
    Public Sub CheckAutoRuleExpiry(ByVal control As Office.IRibbonControl)
        parentObj.GoCheckAutoRuleExpiry(control)
    End Sub
    Public Sub FindInSubfolders(ByVal control As Office.IRibbonControl)
        parentObj.GoFindInSubfolders(control)
    End Sub

#End Region


#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
