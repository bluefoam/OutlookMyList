Imports Microsoft.Office.Tools.Ribbon

Public Class OutlookRibbon

    Private Sub ToggleButton1_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton1.Click
        Globals.ThisAddIn.ToggleTaskPane()
    End Sub

End Class