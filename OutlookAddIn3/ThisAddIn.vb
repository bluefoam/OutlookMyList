Imports System.Diagnostics  ' Add this import statement at the top of the file

Public Class ThisAddIn
    Private WithEvents currentExplorer As Outlook.Explorer
    Private customTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private mailThreadPane As MailThreadPane
    Private taskMonitor As TaskMonitor

    ' 添加防重复调用变量
    Private lastUpdateTime As DateTime = DateTime.MinValue
    Private lastMailEntryID As String = String.Empty
    Private Const UpdateThreshold As Integer = 500 ' 毫秒
    Private isUpdating As Boolean = False

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        currentExplorer = Me.Application.ActiveExplorer
        InitializeMailPane()

        ' 初始化任务监视器
        taskMonitor = New TaskMonitor()
        taskMonitor.Initialize()
    End Sub

    Private Sub InitializeMailPane()
        mailThreadPane = New MailThreadPane()
        customTaskPane = Me.CustomTaskPanes.Add(mailThreadPane, "相关邮件")
        customTaskPane.Width = 400
        customTaskPane.Visible = True
        ' Initialize with empty values
        mailThreadPane.UpdateMailList(String.Empty, String.Empty)
    End Sub

    Private Sub currentExplorer_SelectionChange() Handles currentExplorer.SelectionChange
        If mailThreadPane Is Nothing OrElse customTaskPane Is Nothing OrElse Not customTaskPane.Visible Then Return

        If currentExplorer.Selection.Count > 0 Then
            Dim selection As Object = currentExplorer.Selection(1)
            UpdateMailContent(selection)
        End If
    End Sub

    Public Sub ToggleTaskPane()
        If customTaskPane IsNot Nothing Then
            customTaskPane.Visible = Not customTaskPane.Visible
            ' 显示窗格时，获取当前选中项并更新内容
            If customTaskPane.Visible Then
                If currentExplorer IsNot Nothing AndAlso currentExplorer.Selection.Count > 0 Then
                    Dim currentItem As Object = currentExplorer.Selection(1)
                    UpdateMailContent(currentItem)
                Else
                    ' 如果没有选中项，清空内容
                    mailThreadPane?.UpdateMailList(String.Empty, String.Empty)
                End If
            End If
        End If
    End Sub

    Public Sub UpdateMailList()
        If mailThreadPane IsNot Nothing Then
            mailThreadPane.UpdateMailList(String.Empty, String.Empty)
        End If
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' 注销事件处理程序
        If currentExplorer IsNot Nothing Then
            ' 显式移除事件处理程序
            RemoveHandler currentExplorer.SelectionChange, AddressOf currentExplorer_SelectionChange
        End If

        ' 清理任务监视器
        If taskMonitor IsNot Nothing Then
            taskMonitor.Cleanup()
        End If

        ' 释放资源
        If mailThreadPane IsNot Nothing Then
            mailThreadPane.Dispose()
        End If
    End Sub

    Private Sub Application_ItemLoad(item As Object) Handles Application.ItemLoad
        Try
            ' 检查任务窗格是否可见
            If mailThreadPane IsNot Nothing AndAlso customTaskPane IsNot Nothing AndAlso customTaskPane.Visible Then
                ' 避免与 SelectionChange 事件冲突
                System.Threading.Thread.Sleep(100)
                UpdateMailContent(item)
            End If
        Catch ex As Exception
            Debug.WriteLine($"ItemLoad error: {ex.Message}")
        End Try
    End Sub

    Private Sub UpdateMailContentOld(item As Object)
        Try
            ' 防重复调用检查
            If isUpdating Then
                Debug.WriteLine("UpdateMailContent: 已有更新操作正在进行中，跳过")
                Return
            End If

            ' 获取当前邮件的 EntryID
            Dim mailEntryID As String = String.Empty
            Dim conversationID As String = String.Empty

            If TypeOf item Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                If mail.ConversationID IsNot Nothing Then
                    mailThreadPane.UpdateMailList(mail.ConversationID, mail.EntryID)
                End If
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(item, Outlook.AppointmentItem)
                If appointment.GlobalAppointmentID IsNot Nothing Then
                    mailThreadPane.UpdateMailList(appointment.GlobalAppointmentID, appointment.EntryID)
                End If
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                Dim meeting As Outlook.MeetingItem = DirectCast(item, Outlook.MeetingItem)
                mailThreadPane.UpdateMailList(String.Empty, meeting.EntryID)
            ElseIf TypeOf item Is Outlook.TaskItem Then
                Dim task As Outlook.TaskItem = DirectCast(item, Outlook.TaskItem)
                mailThreadPane.UpdateMailList(String.Empty, task.EntryID)
            ElseIf TypeOf item Is Outlook.ContactItem Then
                Dim contact As Outlook.ContactItem = DirectCast(item, Outlook.ContactItem)
                mailThreadPane.UpdateMailList(String.Empty, contact.EntryID)
            End If
        Catch ex As Exception
            'Debug.WriteLine($"UpdateMailContent error: {ex.Message}")
        End Try
    End Sub
    Private Sub UpdateMailContent(item As Object)
        Try
            ' 防重复调用检查
            If isUpdating Then
                Debug.WriteLine("UpdateMailContent: 已有更新操作正在进行中，跳过")
                Return
            End If

            ' 获取当前邮件的 EntryID
            Dim mailEntryID As String = String.Empty
            Dim conversationID As String = String.Empty

            If TypeOf item Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                mailEntryID = mail.EntryID
                conversationID = mail.ConversationID
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(item, Outlook.AppointmentItem)
                mailEntryID = appointment.EntryID
                conversationID = appointment.GlobalAppointmentID
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                Dim meeting As Outlook.MeetingItem = DirectCast(item, Outlook.MeetingItem)
                mailEntryID = meeting.EntryID
            ElseIf TypeOf item Is Outlook.TaskItem Then
                Dim task As Outlook.TaskItem = DirectCast(item, Outlook.TaskItem)
                mailEntryID = task.EntryID
            ElseIf TypeOf item Is Outlook.ContactItem Then
                Dim contact As Outlook.ContactItem = DirectCast(item, Outlook.ContactItem)
                mailEntryID = contact.EntryID
            End If

            ' 检查是否是同一封邮件的重复调用
            Dim currentTime = DateTime.Now
            If mailEntryID = lastMailEntryID AndAlso
               (currentTime - lastUpdateTime).TotalMilliseconds < UpdateThreshold Then
                Debug.WriteLine($"跳过重复更新，时间间隔: {(currentTime - lastUpdateTime).TotalMilliseconds}ms")
                Return
            End If

            ' 更新最后处理的邮件和时间
            lastMailEntryID = mailEntryID
            lastUpdateTime = currentTime

            ' 设置更新标志
            isUpdating = True

            ' 调用更新方法
            If Not String.IsNullOrEmpty(mailEntryID) Then
                mailThreadPane.UpdateMailList(conversationID, mailEntryID)
            End If

        Catch ex As Exception
            Debug.WriteLine($"UpdateMailContent error: {ex.Message}")
        Finally
            ' 重置更新标志
            isUpdating = False
        End Try
    End Sub
End Class
