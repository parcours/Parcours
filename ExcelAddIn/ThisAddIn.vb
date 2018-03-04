'    Parcours
'    Copyright (C) 2018 Stephan Cieszynski
'    
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'    
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'    
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'

Public Class ThisAddIn

    Private MyUserControlParticipants As UserControlParticipants
    Private WithEvents MyCustomTaskPane As Microsoft.Office.Tools.CustomTaskPane

    Public ReadOnly Property TaskPane() As Microsoft.Office.Tools.CustomTaskPane
        Get
            Return MyCustomTaskPane
        End Get
    End Property

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        MyUserControlParticipants = New UserControlParticipants()
        MyCustomTaskPane = Me.CustomTaskPanes.Add(MyUserControlParticipants, "Partipicants")
        MyCustomTaskPane.Visible = True
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub MyCustomTaskPane_VisibleChanged(ByVal sender As Object,
    ByVal e As System.EventArgs) Handles MyCustomTaskPane.VisibleChanged

        Globals.Ribbons.ExcelRibbonParcours.ToggleUserControlParticipants.Checked = MyCustomTaskPane.Visible
    End Sub
End Class
