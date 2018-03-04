Partial Class ExcelRibbonParcours
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Teilnehmer = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ToggleUserControlParticipants = Me.Factory.CreateRibbonToggleButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Tab1.SuspendLayout()
        Me.Teilnehmer.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Teilnehmer
        '
        Me.Teilnehmer.Groups.Add(Me.Group2)
        Me.Teilnehmer.Groups.Add(Me.Group3)
        Me.Teilnehmer.Label = "Teilnehmer"
        Me.Teilnehmer.Name = "Teilnehmer"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ToggleUserControlParticipants)
        Me.Group2.Label = "Teilnehmer"
        Me.Group2.Name = "Group2"
        '
        'ToggleUserControlParticipants
        '
        Me.ToggleUserControlParticipants.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ToggleUserControlParticipants.Label = "Teilnehmer"
        Me.ToggleUserControlParticipants.Name = "ToggleUserControlParticipants"
        Me.ToggleUserControlParticipants.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Label = "Group3"
        Me.Group3.Name = "Group3"
        '
        'ExcelRibbonParcours
        '
        Me.Name = "ExcelRibbonParcours"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Teilnehmer)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Teilnehmer.ResumeLayout(False)
        Me.Teilnehmer.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Teilnehmer As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ToggleUserControlParticipants As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ExcelRibbonParcours() As ExcelRibbonParcours
        Get
            Return Me.GetRibbon(Of ExcelRibbonParcours)()
        End Get
    End Property
End Class
