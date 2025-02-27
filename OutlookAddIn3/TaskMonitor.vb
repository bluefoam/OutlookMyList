Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics

Public Class TaskMonitor
    Private WithEvents taskFolder As Outlook.MAPIFolder
    Private WithEvents items As Outlook.Items
    Private WithEvents explorer As Outlook.Explorer

    Public Sub Initialize()
        Try
            Dim outlook As Outlook.Application = Globals.ThisAddIn.Application
            taskFolder = outlook.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderTasks)
            items = taskFolder.Items
            explorer = outlook.ActiveExplorer()

            AddHandler explorer.SelectionChange, AddressOf Explorer_SelectionChange
        Catch ex As System.Exception
            Debug.WriteLine($"初始化TaskMonitor时出? {ex.Message}")
        End Try
    End Sub

    Private Sub Explorer_SelectionChange()
        Try
            ' 处理任务?
            If explorer.CurrentFolder.DefaultItemType = OlItemType.olTaskItem Then
                If explorer.Selection.Count > 0 Then
                    'Dim selectedItem As Object = explorer.Selection(1)
                    'If TypeOf selectedItem Is Outlook.TaskItem Then
                    '    Dim task As Outlook.TaskItem = DirectCast(selectedItem, Outlook.TaskItem)
                    '    PrintTaskProperties(task)
                    'End If
                    ' 处理邮件�?
                    'If TypeOf selectedItem Is Outlook.MailItem Then
                    '    Dim mail As Outlook.MailItem = DirectCast(selectedItem, Outlook.MailItem)
                    '    PrintMailProperties(mail)
                    'End If
                End If
                Return
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"处理选择变更时出? {ex.Message}")
        End Try
    End Sub

    Public Sub Cleanup()
        Try
            If explorer IsNot Nothing Then
                RemoveHandler explorer.SelectionChange, AddressOf Explorer_SelectionChange
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"清理TaskMonitor时出? {ex.Message}")
        End Try
    End Sub
End Class