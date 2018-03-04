Public Class ThisAddIn

    Private myParticipantsUserControl As ParticipantsUserControl
    Private WithEvents myCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public ReadOnly Property TaskPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return myCustomTaskPane
        End Get
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        myParticipantsUserControl = New ParticipantsUserControl
        myCustomTaskPane = Me.CustomTaskPanes.Add(myParticipantsUserControl, "Teilnehmer")
        myCustomTaskPane.Visible = True

    End Sub

    Private Sub myCustomTaskPane_VisibleChanged(ByVal sender As Object,
        ByVal e As System.EventArgs) Handles myCustomTaskPane.VisibleChanged

        Globals.Ribbons.ParcoursExcelRibbon.ToggleButtonParticipants.Checked = myCustomTaskPane.Visible
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
