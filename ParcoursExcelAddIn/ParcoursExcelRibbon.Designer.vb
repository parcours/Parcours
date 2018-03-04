Partial Class ParcoursExcelRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Erforderlich für die Unterstützung des Windows.Forms-Klassenkompositions-Designers
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Dieser Aufruf ist für den Komponenten-Designer erforderlich.
        InitializeComponent()

    End Sub

    'Die Komponente überschreibt den Löschvorgang zum Bereinigen der Komponentenliste.
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

    'Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    'Das Bearbeiten ist mit dem Komponenten-Designer möglich.
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TabParticipants = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.ToggleButtonParticipants = Me.Factory.CreateRibbonToggleButton
        Me.TabParticipants.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabParticipants
        '
        Me.TabParticipants.Groups.Add(Me.Group1)
        Me.TabParticipants.Label = "Teilnehmer"
        Me.TabParticipants.Name = "TabParticipants"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.ToggleButtonParticipants)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'ToggleButtonParticipants
        '
        Me.ToggleButtonParticipants.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ToggleButtonParticipants.Label = "ToggleButton1"
        Me.ToggleButtonParticipants.Name = "ToggleButtonParticipants"
        Me.ToggleButtonParticipants.ShowImage = True
        '
        'ParcoursExcelRibbon
        '
        Me.Name = "ParcoursExcelRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.TabParticipants)
        Me.TabParticipants.ResumeLayout(False)
        Me.TabParticipants.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabParticipants As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ToggleButtonParticipants As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ParcoursExcelRibbon() As ParcoursExcelRibbon
        Get
            Return Me.GetRibbon(Of ParcoursExcelRibbon)()
        End Get
    End Property
End Class
