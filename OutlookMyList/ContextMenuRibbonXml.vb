Imports Microsoft.Office.Core
Imports System.Text

Public Class ContextMenuRibbonXml
    Implements IRibbonExtensibility

    Private ribbon As IRibbonUI

    Public Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Try
            Globals.ThisAddIn.LogInfo($"GetCustomUI 调用: ribbonID={ribbonID}")
        Catch
        End Try

        ' 提供多个上下文菜单的覆盖，兼容不同视图/选择场景
        Dim xml As New StringBuilder()
        xml.Append("" & _
            "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>" & _
            "  <contextMenus>" & _
            "    <contextMenu idMso='ContextMenuMailItem'>" & _
            "      <button id='MergeCustomConversationButton' label='合并自定义会话'" & _
            "              imageMso='GroupInsertTable' onAction='OnMergeCustomConversation'" & _
            "              getEnabled='GetMergeEnabled' insertBeforeMso='ReplyAll' />" & _
            "    </contextMenu>" & _
            "    <contextMenu idMso='ContextMenuReadOnlyMailItem'>" & _
            "      <button id='MergeCustomConversationButton_ReadOnly' label='合并自定义会话'" & _
            "              imageMso='GroupInsertTable' onAction='OnMergeCustomConversation'" & _
            "              getEnabled='GetMergeEnabled' />" & _
            "    </contextMenu>" & _
            "    <contextMenu idMso='ContextMenuReadingPane'>" & _
            "      <button id='MergeCustomConversationButton_ReadingPane' label='合并自定义会话'" & _
            "              imageMso='GroupInsertTable' onAction='OnMergeCustomConversation'" & _
            "              getEnabled='GetMergeEnabled' />" & _
            "    </contextMenu>" & _
            "  </contextMenus>" & _
            "</customUI>")
        Return xml.ToString()
    End Function

    ' Ribbon 加载
    Public Sub OnRibbonLoad(ribbonUI As IRibbonUI)
        ribbon = ribbonUI
        Try
            Globals.ThisAddIn.LogInfo("Ribbon XML 上下文菜单已加载")
        Catch
        End Try
    End Sub

    ' 合并自定义会话回调
    Public Sub OnMergeCustomConversation(control As IRibbonControl)
        Try
            Globals.ThisAddIn.HandleMergeCustomConversation()
        Catch ex As Exception
            Try
                Globals.ThisAddIn.LogException(ex, "OnMergeCustomConversation")
            Catch
            End Try
        End Try
    End Sub

    ' 根据选择数量控制启用状态
    Public Function GetMergeEnabled(control As IRibbonControl) As Boolean
        Try
            Dim count As Integer = 0
            Dim explorer = Globals.ThisAddIn.Application.ActiveExplorer
            If explorer IsNot Nothing AndAlso explorer.Selection IsNot Nothing Then
                count = explorer.Selection.Count
            End If
            Return count >= 2
        Catch
            Return True
        End Try
    End Function
End Class