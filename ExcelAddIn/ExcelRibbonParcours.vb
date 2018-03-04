Imports Microsoft.Office.Tools.Ribbon

Public Class ExcelRibbonParcours

    Private Sub ExcelRibbonParcours_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub ToggleUserControlParticipants_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleUserControlParticipants.Click
        Globals.ThisAddIn.TaskPane.Visible = TryCast(sender, RibbonToggleButton).Checked
    End Sub

End Class
