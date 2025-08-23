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

End Class