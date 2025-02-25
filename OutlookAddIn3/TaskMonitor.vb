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
            Debug.WriteLine($"初始化TaskMonitor时出错: {ex.Message}")
        End Try
    End Sub

    Private Sub Explorer_SelectionChange()
        Try
            ' 处理任务项
            If explorer.CurrentFolder.DefaultItemType = OlItemType.olTaskItem Then
                If explorer.Selection.Count > 0 Then
                    Dim selectedItem As Object = explorer.Selection(1)
                    If TypeOf selectedItem Is Outlook.TaskItem Then
                        Dim task As Outlook.TaskItem = DirectCast(selectedItem, Outlook.TaskItem)
                        PrintTaskProperties(task)
                    End If
                    ' 处理邮件项
                    If TypeOf selectedItem Is Outlook.MailItem Then
                        Dim mail As Outlook.MailItem = DirectCast(selectedItem, Outlook.MailItem)
                        PrintMailProperties(mail)
                    End If
                End If
                Return
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"处理选择变更时出错: {ex.Message}")
        End Try
    End Sub

    ' 添加邮件属性打印方法
    Private Sub PrintMailProperties(mail As Outlook.MailItem)
        Debug.WriteLine("========== 邮件属性 ==========")
        Debug.WriteLine($"基本属性:")
        Debug.WriteLine($"- 主题: {mail.Subject}")
        Debug.WriteLine($"- EntryID: {mail.EntryID}")
        Debug.WriteLine($"- 会话ID: {mail.ConversationID}")
        Debug.WriteLine($"- 发件人: {mail.SenderName}")
        Debug.WriteLine($"- 发件人邮箱: {mail.SenderEmailAddress}")
        Debug.WriteLine($"- 收件人: {mail.To}")
        Debug.WriteLine($"- 抄送: {mail.CC}")
        Debug.WriteLine($"- 密送: {mail.BCC}")
        Debug.WriteLine($"- 发送时间: {mail.SentOn}")
        Debug.WriteLine($"- 接收时间: {mail.ReceivedTime}")
        Debug.WriteLine($"- 大小: {mail.Size} bytes")
        Debug.WriteLine($"- 是否已读: {mail.UnRead}")
        Debug.WriteLine($"- 重要性: {mail.Importance}")
        Debug.WriteLine($"- 分类: {mail.Categories}")

        Debug.WriteLine($"Links属性:")
        Try
            If mail.Links IsNot Nothing Then
                Debug.WriteLine($"- Links数量: {mail.Links.Count}")
                For Each link As Outlook.Link In mail.Links
                    Debug.WriteLine($"  - Link类型: {link.Item.GetType().Name}")
                    If TypeOf link.Item Is Outlook.TaskItem Then
                        Dim linkedTask As Outlook.TaskItem = DirectCast(link.Item, Outlook.TaskItem)
                        Debug.WriteLine($"  - 关联任务主题: {linkedTask.Subject}")
                        Debug.WriteLine($"  - 关联任务EntryID: {linkedTask.EntryID}")
                    End If
                Next
            Else
                Debug.WriteLine("- Links不可用")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"- 访问Links时出错: {ex.Message}")
        End Try

        Debug.WriteLine($"自定义属性:")
        For Each prop As Outlook.UserProperty In mail.UserProperties
            Debug.WriteLine($"- {prop.Name}: {prop.Value}")
        Next

        Debug.WriteLine($"所有属性:")
        Try
            For Each prop As Object In mail.ItemProperties
                Try
                    Debug.WriteLine($"- {prop.Name}: {prop.Value}")
                Catch propEx As System.Exception
                    Debug.WriteLine($"- 无法读取属性值: {prop.Name}")
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"读取ItemProperties时出错: {ex.Message}")
        End Try
        Debug.WriteLine("============================")
    End Sub

    Private Sub PrintTaskProperties(task As Outlook.TaskItem)
        Debug.WriteLine("========== 任务属性 ==========")
        Debug.WriteLine($"基本属性:")
        Debug.WriteLine($"- 主题: {task.Subject}")
        Debug.WriteLine($"- EntryID: {task.EntryID}")
        Debug.WriteLine($"- 创建时间: {task.CreationTime}")
        Debug.WriteLine($"- 最后修改时间: {task.LastModificationTime}")
        Debug.WriteLine($"- 状态: {task.Status}")
        Debug.WriteLine($"- 完成百分比: {task.PercentComplete}")
        Debug.WriteLine($"- 到期日: {If(task.DueDate = DateTime.MinValue, "无", task.DueDate.ToString())}")
        Debug.WriteLine($"- 优先级: {task.Importance}")
        Debug.WriteLine($"- 分类: {task.Categories}")
        Debug.WriteLine($"- 主体内容: {task.Body}")

        Debug.WriteLine($"Links属性:")
        Try
            If task.Links IsNot Nothing Then
                Debug.WriteLine($"- Links数量: {task.Links.Count}")
                For Each link As Outlook.Link In task.Links
                    Debug.WriteLine($"  - Link类型: {link.Item.GetType().Name}")
                    If TypeOf link.Item Is Outlook.MailItem Then
                        Dim linkedMail As Outlook.MailItem = DirectCast(link.Item, Outlook.MailItem)
                        Debug.WriteLine($"  - 关联邮件主题: {linkedMail.Subject}")
                        Debug.WriteLine($"  - 关联邮件EntryID: {linkedMail.EntryID}")
                    End If
                Next
            Else
                Debug.WriteLine("- Links不可用")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"- 访问Links时出错: {ex.Message}")
        End Try

        Debug.WriteLine($"自定义属性:")
        For Each prop As Outlook.UserProperty In task.UserProperties
            Debug.WriteLine($"- {prop.Name}: {prop.Value}")
        Next

        Debug.WriteLine($"所有属性:")
        Try
            For Each prop As Object In task.ItemProperties
                Try
                    Debug.WriteLine($"- {prop.Name}: {prop.Value}")
                Catch propEx As System.Exception
                    Debug.WriteLine($"- 无法读取属性值: {prop.Name}")
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"读取ItemProperties时出错: {ex.Message}")
        End Try
        Debug.WriteLine("============================")
    End Sub

    Public Sub Cleanup()
        Try
            If explorer IsNot Nothing Then
                RemoveHandler explorer.SelectionChange, AddressOf Explorer_SelectionChange
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"清理TaskMonitor时出错: {ex.Message}")
        End Try
    End Sub
End Class