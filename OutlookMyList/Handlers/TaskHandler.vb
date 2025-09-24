Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics
Imports System.Windows.Forms

Namespace OutlookMyList.Handlers
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
                Dim item As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mailEntryID)
                If item IsNot Nothing AndAlso TypeOf item Is MailItem Then
                    Return DirectCast(item, MailItem)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"获取邮件失败: {ex.Message}")
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 根据邮件EntryID获取关联的任务信息
        ''' </summary>
        ''' <param name="mailEntryID">邮件EntryID</param>
        ''' <returns>关联的任务信息，如果没有关联任务则返回Nothing</returns>
        Public Shared Function GetTaskByMailEntryID(mailEntryID As String) As OutlookMyList.Models.TaskInfo
            Try
                If String.IsNullOrEmpty(mailEntryID) Then
                    Return Nothing
                End If

                ' 首先检查邮件是否被标记为任务
                Dim mailItem As MailItem = GetMailItem(mailEntryID)
                If mailItem IsNot Nothing Then
                    Try
                        If mailItem.IsMarkedAsTask Then
                            ' 创建基于邮件标记的任务信息
                            Dim taskInfo As New OutlookMyList.Models.TaskInfo With {
                                .Subject = mailItem.TaskSubject,
                                .MailEntryID = mailEntryID,
                                .RelatedMailSubject = mailItem.Subject,
                                .DueDate = mailItem.TaskDueDate,
                                .Status = GetTaskStatusText(mailItem.TaskStatus),
                                .PercentComplete = mailItem.TaskCompletedDate.CompareTo(DateTime.MinValue)
                            }
                            Return taskInfo
                        End If
                    Finally
                        Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                    End Try
                End If

                ' 然后检查是否有独立的任务项关联到这个邮件
                Dim outlookApp = Globals.ThisAddIn.Application
                Dim tasksFolder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderTasks)
                
                For Each item As Object In tasksFolder.Items
                    If TypeOf item Is TaskItem Then
                        Dim task As TaskItem = DirectCast(item, TaskItem)
                        Try
                            ' 检查任务的链接项
                            If task.Links IsNot Nothing AndAlso task.Links.Count > 0 Then
                                For Each link As Link In task.Links
                                    If TypeOf link.Item Is MailItem Then
                                        Dim linkedMail As MailItem = DirectCast(link.Item, MailItem)
                                        If linkedMail.EntryID = mailEntryID Then
                                            ' 找到关联的任务
                                            Dim taskInfo As New OutlookMyList.Models.TaskInfo With {
                                                .TaskEntryID = task.EntryID,
                                                .MailEntryID = mailEntryID,
                                                .Subject = task.Subject,
                                                .DueDate = task.DueDate,
                                                .Status = GetTaskStatusText(task.Status),
                                                .PercentComplete = task.PercentComplete,
                                                .RelatedMailSubject = linkedMail.Subject
                                            }
                                            Runtime.InteropServices.Marshal.ReleaseComObject(linkedMail)
                                            Runtime.InteropServices.Marshal.ReleaseComObject(task)
                                            Return taskInfo
                                        End If
                                        Runtime.InteropServices.Marshal.ReleaseComObject(linkedMail)
                                    End If
                                Next
                            End If
                            
                            ' 检查用户属性中的邮件ID关联
                            For Each prop As UserProperty In task.UserProperties
                                If prop.Name = "MailEntryID" AndAlso prop.Value?.ToString() = mailEntryID Then
                                    Dim taskInfo As New OutlookMyList.Models.TaskInfo With {
                                        .TaskEntryID = task.EntryID,
                                        .MailEntryID = mailEntryID,
                                        .Subject = task.Subject,
                                        .DueDate = task.DueDate,
                                        .Status = GetTaskStatusText(task.Status),
                                        .PercentComplete = task.PercentComplete
                                    }
                                    Runtime.InteropServices.Marshal.ReleaseComObject(task)
                                    Return taskInfo
                                End If
                            Next
                        Finally
                            Runtime.InteropServices.Marshal.ReleaseComObject(task)
                        End Try
                    End If
                Next
                
            Catch ex As System.Exception
                Debug.WriteLine($"GetTaskByMailEntryID error: {ex.Message}")
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
            Try
                Dim storeIdForTask As String = Nothing
                Try
                    Dim parentFolder = TryCast(task.Parent, MAPIFolder)
                    If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                        storeIdForTask = parentFolder.Store.StoreID
                    End If
                Catch
                End Try

                Dim mailLink As MailItem = Nothing
                Dim mailStoreId As String = Nothing
                Try
                    If task.Links IsNot Nothing AndAlso task.Links.Count > 0 Then
                        mailLink = TryCast(task.Links(1).Item, MailItem)
                        If mailLink IsNot Nothing Then
                            Dim parentFolder2 = TryCast(mailLink.Parent, MAPIFolder)
                            If parentFolder2 IsNot Nothing AndAlso parentFolder2.Store IsNot Nothing Then
                                mailStoreId = parentFolder2.Store.StoreID
                            End If
                        End If
                    End If
                Catch
                End Try

                Dim taskInfo As New OutlookMyList.Models.TaskInfo With {
                    .TaskEntryID = task.EntryID,
                    .MailEntryID = If(mailLink IsNot Nothing, mailLink.EntryID, String.Empty),
                    .Subject = task.Subject,
                    .DueDate = If(task.DueDate = #12:00:00 AM#, Nothing, task.DueDate),
                    .Status = task.Status.ToString(),
                    .PercentComplete = task.PercentComplete,
                    .LinkedMailSubject = linkedMailSubject,
                    .StoreID = If(Not String.IsNullOrEmpty(mailStoreId), mailStoreId, storeIdForTask)
                }

                Dim listItem As New ListViewItem(task.Subject)
                listItem.SubItems.Add(If(task.DueDate = DateTime.MinValue, "", task.DueDate.ToString("yyyy-MM-dd")))
                listItem.SubItems.Add(GetTaskStatusText(task.Status))
                listItem.SubItems.Add($"{task.PercentComplete}%")
                listItem.SubItems.Add("(标准任务)")
                listItem.Tag = taskInfo
            Catch ex As System.Runtime.InteropServices.COMException
                Debug.WriteLine($"COM异常访问任务属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Dim listItem As New ListViewItem("无法访问任务")
                listItem.SubItems.Add("无法访问")
                listItem.SubItems.Add("无法访问")
                listItem.SubItems.Add("无法访问")
                listItem.SubItems.Add("(标准任务)")
                listItem.Tag = Nothing
                taskList.Items.Add(listItem)
            Catch ex As System.Exception
                Debug.WriteLine($"添加任务到列表时出错: {ex.Message}")
            End Try
        End Sub

        Private Shared Function GetMailConversationID(mailEntryID As String) As String
            Try
                Dim mail As MailItem = DirectCast(
                    OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mailEntryID),
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
                        Try
                            ' 直接使用 IsMarkedAsTask 属性判断
                            If mail.IsMarkedAsTask Then
                                Dim storeId As String = Nothing
                                Try
                                    Dim parentFolder = TryCast(mail.Parent, MAPIFolder)
                                    If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                                        storeId = parentFolder.Store.StoreID
                                    End If
                                Catch
                                End Try

                                Dim taskInfo As New OutlookMyList.Models.TaskInfo With {
                                    .Subject = mail.TaskSubject,
                                    .MailEntryID = mail.EntryID,
                                    .RelatedMailSubject = mail.Subject,
                                    .DueDate = If(mail.TaskDueDate = DateTime.MinValue, Nothing, mail.TaskDueDate),
                                    .Status = GetTaskStatusText(mail.TaskStatus),
                                    .PercentComplete = mail.PercentComplete,
                                    .StoreID = storeId
                                }

                                Dim listItem As New ListViewItem(taskInfo.Subject)
                                listItem.SubItems.Add(If(taskInfo.DueDate.HasValue, taskInfo.DueDate.Value.ToString("yyyy-MM-dd"), ""))
                                listItem.SubItems.Add(taskInfo.Status)
                                listItem.SubItems.Add($"{taskInfo.PercentComplete}%")
                                listItem.SubItems.Add(taskInfo.RelatedMailSubject)
                                listItem.Tag = taskInfo
                                taskList.Items.Add(listItem)
                            End If
                        Catch ex As System.Runtime.InteropServices.COMException
                            Debug.WriteLine($"COM异常访问邮件任务属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                            Dim listItem As New ListViewItem("无法访问任务")
                            listItem.SubItems.Add("无法访问")
                            listItem.SubItems.Add("无法访问")
                            listItem.SubItems.Add("无法访问")
                            listItem.SubItems.Add("无法访问")
                            listItem.Tag = Nothing
                            taskList.Items.Add(listItem)
                        Catch ex As System.Exception
                            Debug.WriteLine($"访问邮件任务属性时发生异常: {ex.Message}")
                            Dim listItem As New ListViewItem("无法访问任务")
                            listItem.SubItems.Add("无法访问")
                            listItem.SubItems.Add("无法访问")
                            listItem.SubItems.Add("无法访问")
                            listItem.SubItems.Add("无法访问")
                            listItem.Tag = Nothing
                            taskList.Items.Add(listItem)
                        End Try
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"LoadAnnotatedTasks error: {ex.Message}")
            End Try
        End Sub

        Private Shared Function GetAnnotatedTasksFromMails(conversationId As String) As List(Of OutlookMyList.Models.TaskInfo)
            Dim tasks As New List(Of OutlookMyList.Models.TaskInfo)
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
                        Try
                            Dim props As ItemProperties = mail.ItemProperties

                            ' 从邮件的任务属性中获取信息
                            If props("TaskSubject") IsNot Nothing Then
                                tasks.Add(New OutlookMyList.Models.TaskInfo With {
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
                        Catch ex As System.Runtime.InteropServices.COMException
                            Debug.WriteLine($"COM异常访问邮件任务属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                        Catch ex As System.Exception
                            Debug.WriteLine($"访问邮件任务属性时发生异常: {ex.Message}")
                        End Try
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

        Private Shared Function ParseTasksFromMail(mail As Outlook.MailItem) As List(Of OutlookMyList.Models.TaskInfo)
            Dim tasks As New List(Of OutlookMyList.Models.TaskInfo)
            Try
                ' 在这里实现你的邮件任务标记解析逻辑
                ' 例如：查找特定格式的标记，如 [Task]、TODO: 等
                ' 这是一个示例实现
                Try
                    Dim body As String = mail.Body
                    Dim lines = body.Split(New String() {vbCrLf, vbCr, vbLf}, StringSplitOptions.None)

                    For Each line In lines
                        If line.Trim().StartsWith("[Task]") OrElse line.Trim().StartsWith("TODO:") Then
                            tasks.Add(New OutlookMyList.Models.TaskInfo With {
                                .Subject = line.Trim(),
                                .MailEntryID = mail.EntryID,
                                .RelatedMailSubject = mail.Subject
                            })
                        End If
                    Next
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Catch ex As System.Exception
                    Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                End Try
            Catch ex As System.Exception
                Debug.WriteLine($"ParseTasksFromMail error: {ex.Message}")
            End Try
            Return tasks
        End Function

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
                Dim storeId As String = Nothing
                Try
                    Dim parentFolder = TryCast(item.Parent, MAPIFolder)
                    If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                        storeId = parentFolder.Store.StoreID
                    End If
                Catch
                End Try

                listItem.Tag = New OutlookMyList.Models.TaskInfo With {
                    .TaskEntryID = item.EntryID,
                    .MailEntryID = item.EntryID,
                    .StoreID = storeId
                }
                taskList.Items.Add(listItem)
            Catch ex As System.Exception
                Debug.WriteLine($"添加邮件标记任务到列表时出错: {ex.Message}")
            End Try
        End Sub
    End Class
End Namespace