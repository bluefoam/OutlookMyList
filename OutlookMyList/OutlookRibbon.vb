Imports Microsoft.Office.Tools.Ribbon

Imports System.Diagnostics

Public Class OutlookRibbon

    Private Sub OutlookRibbon_Load(sender As Object, e As RibbonUIEventArgs) Handles MyBase.Load
        ' 初始化分页按钮状态
        If Globals.ThisAddIn.MailThreadPaneInstance IsNot Nothing Then
            ToggleButtonPagination.Checked = Globals.ThisAddIn.MailThreadPaneInstance.IsPaginationEnabled
        End If

        ' 同步缓存开关按钮状态
        ToggleButtonCacheEnabled.Checked = Globals.ThisAddIn.CacheEnabled
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

    Private Sub ToggleButtonCacheEnabled_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButtonCacheEnabled.Click
        ' 切换缓存开关
        Globals.ThisAddIn.SaveCacheEnabledToRegistry(ToggleButtonCacheEnabled.Checked)
    End Sub

    Private Sub ButtonMergeConversation_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonMergeConversation.Click
        ' 最简单方式：直接复用合并逻辑
        Globals.ThisAddIn.HandleMergeCustomConversation()
    End Sub

    ' 更新“合并自定义会话”按钮的启用状态
    Public Sub UpdateMergeButtonState(enabled As Boolean)
        Try
            ButtonMergeConversation.Enabled = enabled
        Catch ex As Exception
            Debug.WriteLine($"Error updating merge button state: {ex.Message}")
        End Try
    End Sub

    ' 加载时根据当前选择初始化按钮状态
    Private Sub OutlookRibbon_AfterLoad() Handles MyBase.Load
        Try
            Dim explorer = Globals.ThisAddIn.Application?.ActiveExplorer
            Dim selCount As Integer = 0
            If explorer IsNot Nothing Then
                selCount = explorer.Selection.Count
            End If
            ButtonMergeConversation.Enabled = (selCount >= 2)
        Catch
            ' 忽略初始化异常，避免影响Ribbon显示
        End Try
    End Sub

End Class