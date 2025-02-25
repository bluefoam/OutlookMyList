Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics
Imports System.Windows.Forms
Imports OutlookAddIn3.Models

Namespace OutlookAddIn3.Handlers
    Public Class TaskHandler
        Public Shared Function GetTaskStatusText(status As OlTaskStatus) As String
            Select Case status
                Case OlTaskStatus.olTaskNotStarted
                    Return "未开始"
                Case OlTaskStatus.olTaskInProgress
                    Return "进行中"
                Case OlTaskStatus.olTaskComplete
                    Return "已完成"
                Case OlTaskStatus.olTaskWaiting
                    Return "等待中"
                Case OlTaskStatus.olTaskDeferred
                    Return "已推迟"
                Case Else
                    Return "未知"
            End Select
        End Function

        Public Shared Function GetTaskMailEntryID(task As TaskItem) As String
            Try
                If task.Links IsNot Nothing AndAlso task.Links.Count > 0 Then
                    For Each link As Link In task.Links
                        If TypeOf link.Item Is MailItem Then
                            Return DirectCast(link.Item, MailItem).EntryID
                        End If
                    Next
                End If

                For Each prop As UserProperty In task.UserProperties
                    If prop.Name = "MailEntryID" Then
                        Return prop.Value.ToString()
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"读取任务关联邮件时出错: {ex.Message}")
            End Try
            Return String.Empty
        End Function

        Public Shared Function GetMailItem(mailEntryID As String) As MailItem
            Try
                Dim item As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailEntryID)
                If item IsNot Nothing AndAlso TypeOf item Is MailItem Then
                    Return DirectCast(item, MailItem)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"获取邮件失败: {ex.Message}")
            End Try
            Return Nothing
        End Function

        Public Shared Sub SetupTaskList(taskList As ListView)
            taskList.View = Windows.Forms.View.Details
            taskList.FullRowSelect = True
            taskList.GridLines = True

            ' 添加列
            taskList.Columns.Clear()  ' 清除现有列
            taskList.Columns.Add("主题", 200)
            taskList.Columns.Add("到期日", 100)
            taskList.Columns.Add("状态", 80)
            taskList.Columns.Add("完成度", 80)
            taskList.Columns.Add("关联邮件", 200)
        End Sub

        ' 修改 TaskInfo 的引用，使用完整命名空间
        Private Shared Sub AddTaskToList(taskList As ListView, task As TaskItem, linkedMailSubject As String)
            Dim taskInfo As New OutlookAddIn3.Models.TaskInfo With {
                .TaskEntryID = task.EntryID,
                .MailEntryID = If(task.Links.Count > 0, DirectCast(task.Links(1).Item, MailItem).EntryID, String.Empty),
                .Subject = task.Subject,
                .DueDate = If(task.DueDate = #12:00:00 AM#, Nothing, task.DueDate),
                .Status = task.Status.ToString(),
                .PercentComplete = task.PercentComplete,
                .LinkedMailSubject = linkedMailSubject
            }

            Try
                Dim listItem As New ListViewItem(task.Subject)
                listItem.SubItems.Add(If(task.DueDate = DateTime.MinValue, "", task.DueDate.ToString("yyyy-MM-dd")))
                listItem.SubItems.Add(GetTaskStatusText(task.Status))
                listItem.SubItems.Add($"{task.PercentComplete}%")
                listItem.SubItems.Add("(标准任务)")
                listItem.Tag = taskInfo
                taskList.Items.Add(listItem)
            Catch ex As System.Exception
                Debug.WriteLine($"添加任务到列表时出错: {ex.Message}")
            End Try
        End Sub

        Private Shared Function GetMailConversationID(mailEntryID As String) As String
            Try
                Dim mail As MailItem = DirectCast(
                    Globals.ThisAddIn.Application.Session.GetItemFromID(mailEntryID),
                    MailItem)
                Return mail.ConversationID
            Catch ex As System.Exception
                Debug.WriteLine($"获取邮件会话ID时出错: {ex.Message}")
                Return String.Empty
            End Try
        End Function

        Public Shared Sub LoadOutlookTasks(taskList As ListView, conversationId As String)
            Try
                taskList.Items.Clear()
                If String.IsNullOrEmpty(conversationId) Then
                    Return
                End If
                ' 加载邮件标注的任务
                LoadAnnotatedTasks(taskList, conversationId)
                ' 加载关联的Outlook任务
                LoadLinkedOutlookTasks(taskList, conversationId)
            Catch ex As System.Exception
                Debug.WriteLine($"LoadOutlookTasks error: {ex.Message}")
            End Try
        End Sub

        Private Shared Sub LoadAnnotatedTasks(taskList As ListView, conversationId As String)
            Try
                ' 获取会话中的所有邮件
                Dim outlookApp = Globals.ThisAddIn.Application
                Dim inbox = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

                ' 只查找当前会话的邮件
                Dim filter As String = $"[ConversationID] = '{conversationId}'"
                Dim items = inbox.Items.Restrict(filter)

                For Each item As Object In items
                    If TypeOf item Is MailItem Then
                        Dim mail As MailItem = DirectCast(item, MailItem)
                        ' 直接使用 IsMarkedAsTask 属性判断
                        If mail.IsMarkedAsTask Then
                            Dim taskInfo As New TaskInfo With {
                                .Subject = mail.TaskSubject,
                                .MailEntryID = mail.EntryID,
                                .RelatedMailSubject = mail.Subject,
                                .DueDate = If(mail.TaskDueDate = DateTime.MinValue, Nothing, mail.TaskDueDate),
                                .Status = GetTaskStatusText(mail.TaskStatus),
                                .PercentComplete = mail.PercentComplete
                            }

                            Dim listItem As New ListViewItem(taskInfo.Subject)
                            listItem.SubItems.Add(If(taskInfo.DueDate.HasValue, taskInfo.DueDate.Value.ToString("yyyy-MM-dd"), ""))
                            listItem.SubItems.Add(taskInfo.Status)
                            listItem.SubItems.Add($"{taskInfo.PercentComplete}%")
                            listItem.SubItems.Add(taskInfo.RelatedMailSubject)
                            listItem.Tag = taskInfo
                            taskList.Items.Add(listItem)
                        End If
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"LoadAnnotatedTasks error: {ex.Message}")
            End Try
        End Sub

        Private Shared Function GetAnnotatedTasksFromMails(conversationId As String) As List(Of TaskInfo)
            Dim tasks As New List(Of TaskInfo)
            Try
                ' 获取会话中的所有邮件
                Dim outlookApp = Globals.ThisAddIn.Application
                Dim inbox = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

                ' 修改查找方式：查找标记为任务的邮件
                Dim filter As String = $"[ConversationID] = '{conversationId}' AND [IsMarkedAsTask] = True"
                Dim items = inbox.Items.Restrict(filter)

                For Each item As Object In items
                    If TypeOf item Is MailItem Then
                        Dim mail As MailItem = DirectCast(item, MailItem)
                        Dim props As ItemProperties = mail.ItemProperties

                        ' 从邮件的任务属性中获取信息
                        If props("TaskSubject") IsNot Nothing Then
                            tasks.Add(New TaskInfo With {
                                .Subject = props("TaskSubject").Value.ToString(),
                                .MailEntryID = mail.EntryID,
                                .RelatedMailSubject = mail.Subject,
                                .DueDate = If(props("TaskDueDate")?.Value IsNot Nothing,
                                            CDate(props("TaskDueDate").Value), Nothing),
                                .Status = If(props("TaskStatus")?.Value IsNot Nothing,
                                           props("TaskStatus").Value.ToString(), "未开始"),
                                .PercentComplete = If(props("TaskComplete")?.Value IsNot Nothing,
                                                    CInt(props("TaskComplete").Value), 0)
                            })
                        End If
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"GetAnnotatedTasksFromMails error: {ex.Message}")
            End Try
            Return tasks
        End Function
        Private Shared Sub LoadLinkedOutlookTasks(taskList As ListView, conversationId As String)
            Try
                ' 获取所有任务文件夹
                Dim outlookApp = Globals.ThisAddIn.Application
                Dim taskFolder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderTasks)

                ' 修改筛选条件：查找标记为任务的项目
                Dim filter As String = $"[IsMarkedAsTask] = True"
                Dim items = taskFolder.Items.Restrict(filter)

                For Each item As Object In items
                    Try
                        Dim props As Outlook.ItemProperties = item.ItemProperties

                        ' 检查是否包含必要的任务属性
                        If props("TaskSubject") IsNot Nothing Then
                            ' 检查是否属于当前会话
                            If props("ConversationID")?.Value?.ToString() = conversationId Then
                                ' 使用 AddMarkedMailTaskToList 来添加标记任务
                                AddMarkedMailTaskToList(taskList, item)
                            End If
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine($"处理标记任务时出错: {ex.Message}")
                        Continue For
                    End Try
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"LoadLinkedOutlookTasks error: {ex.Message}")
            End Try
        End Sub

        Private Shared Function ParseTasksFromMail(mail As Outlook.MailItem) As List(Of TaskInfo)
            Dim tasks As New List(Of TaskInfo)
            Try
                ' 在这里实现你的邮件任务标记解析逻辑
                ' 例如：查找特定格式的标记，如 [Task]、TODO: 等
                ' 这是一个示例实现
                Dim body As String = mail.Body
                Dim lines = body.Split(New String() {vbCrLf, vbCr, vbLf}, StringSplitOptions.None)

                For Each line In lines
                    If line.Trim().StartsWith("[Task]") OrElse line.Trim().StartsWith("TODO:") Then
                        tasks.Add(New TaskInfo With {
                            .Subject = line.Trim(),
                            .MailEntryID = mail.EntryID,
                            .RelatedMailSubject = mail.Subject
                        })
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"ParseTasksFromMail error: {ex.Message}")
            End Try
            Return tasks
        End Function

        ' 添加任务信息类
        Public Class TaskInfo
            Public Property Subject As String
            Public Property DueDate As DateTime?
            Public Property MailEntryID As String
            Public Property RelatedMailSubject As String
            Public Property TaskEntryID As String
            Public Property Status As String
            Public Property PercentComplete As Integer
            Public Property LinkedMailSubject As String
        End Class
        Public Shared Sub CreateNewTask(conversationId As String, mailEntryID As String)
            Try
                Dim outlookApp As Outlook.Application = Globals.ThisAddIn.Application
                Dim task As Outlook.TaskItem = DirectCast(outlookApp.CreateItem(Outlook.OlItemType.olTaskItem), Outlook.TaskItem)

                task.Subject = "新任务"
                task.Body = $"关联邮件ID: {mailEntryID}"
                task.UserProperties.Add("ConversationID", Outlook.OlUserPropertyType.olText).Value = conversationId
                task.UserProperties.Add("RelatedMailID", Outlook.OlUserPropertyType.olText).Value = mailEntryID

                task.Display(False)
            Catch ex As System.Exception
                Debug.WriteLine($"CreateNewTask error: {ex.Message}")
                Throw
            End Try
        End Sub

        Private Shared Sub AddMarkedMailTaskToList(taskList As ListView, item As Object)
            Try
                Dim props As ItemProperties = item.ItemProperties
                Dim listItem As New ListViewItem(props("TaskSubject").Value.ToString())
                listItem.SubItems.Add(If(props("TaskDueDate").Value Is Nothing, "",
                                       CDate(props("TaskDueDate").Value).ToString("yyyy-MM-dd")))
                listItem.SubItems.Add(GetTaskStatusText(CInt(props("TaskStatus").Value)))
                listItem.SubItems.Add($"{props("TaskComplete").Value}%")
                listItem.SubItems.Add("(邮件标记任务)")
                listItem.Tag = New OutlookAddIn3.Models.TaskInfo With {
                    .TaskEntryID = item.EntryID,
                    .MailEntryID = item.EntryID
                }
                taskList.Items.Add(listItem)
            Catch ex As System.Exception
                Debug.WriteLine($"添加邮件标记任务到列表时出错: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace