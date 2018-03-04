Imports Microsoft.Office.Tools.Ribbon

Public Class ParcoursExcelRibbon

    Private Sub ParcoursExcelRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ToggleButtonParticipants_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButtonParticipants.Click
        Globals.ThisAddIn.TaskPane.Visible =
    TryCast(sender, Microsoft.Office.Tools.Ribbon.RibbonToggleButton).Checked
    End Sub
End Class
