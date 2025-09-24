Imports Microsoft.Office.Tools.Ribbon

Imports System.Diagnostics

Public Class OutlookRibbon

    Private Sub OutlookRibbon_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ' 初始化分页按钮状态
        If Globals.ThisAddIn.MailThreadPaneInstance IsNot Nothing Then
            ToggleButtonPagination.Checked = Globals.ThisAddIn.MailThreadPaneInstance.IsPaginationEnabled
        End If
    End Sub

    Private Sub ToggleButton1_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton1.Click
        Globals.ThisAddIn.ToggleTaskPane()
    End Sub

    Private Sub ToggleButtonPagination_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButtonPagination.Click
        ' 获取当前的MailThreadPane实例并切换分页功能
        If Globals.ThisAddIn.MailThreadPaneInstance IsNot Nothing Then
            Globals.ThisAddIn.MailThreadPaneInstance.IsPaginationEnabled = ToggleButtonPagination.Checked
        End If
    End Sub

    ' 更新分页按钮状态的公共方法
    Public Sub UpdatePaginationButtonState(enabled As Boolean)
        Try
            ToggleButtonPagination.Checked = enabled
        Catch ex As Exception
            Debug.WriteLine($"Error updating pagination button state: {ex.Message}")
        End Try
    End Sub

    Private Sub ToggleButtonBottomPane_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButtonBottomPane.Click
        Globals.ThisAddIn.ToggleBottomPane()
        ' 更新按钮状态
        ToggleButtonBottomPane.Checked = Globals.ThisAddIn.IsBottomPaneVisible
    End Sub

    Private Sub ButtonMinimizeBottomPane_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonMinimizeBottomPane.Click
        Globals.ThisAddIn.MinimizeBottomPane()
        ' 更新按钮文本
        If Globals.ThisAddIn.BottomPaneInstance IsNot Nothing Then
            ButtonMinimizeBottomPane.Label = If(Globals.ThisAddIn.BottomPaneInstance.IsMinimized, "还原", "最小化")
        End If
    End Sub

    ' 更新底部面板按钮状态的公共方法
    Public Sub UpdateBottomPaneButtonState(visible As Boolean, minimized As Boolean)
        Try
            ToggleButtonBottomPane.Checked = visible
            ButtonMinimizeBottomPane.Label = If(minimized, "还原", "最小化")
        Catch ex As Exception
            Debug.WriteLine($"Error updating bottom pane button state: {ex.Message}")
        End Try
    End Sub

    Private Sub ToggleButtonEmbeddedPane_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButtonEmbeddedPane.Click
        Globals.ThisAddIn.ToggleEmbeddedBottomPane()
        ' 更新按钮状态
        ToggleButtonEmbeddedPane.Checked = Globals.ThisAddIn.IsEmbeddedBottomPaneVisible
    End Sub

End Class