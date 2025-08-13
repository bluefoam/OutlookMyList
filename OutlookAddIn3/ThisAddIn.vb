Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports Microsoft.Win32

Public Class ThisAddIn
    Private WithEvents currentExplorer As Outlook.Explorer
    Private customTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private mailThreadPane As MailThreadPane
    Private taskMonitor As TaskMonitor

    ' 添加Inspector相关变量
    Private WithEvents inspectors As Outlook.Inspectors
    Private inspectorTaskPanes As New Dictionary(Of String, Microsoft.Office.Tools.CustomTaskPane)

    ' 添加防重复调用变量
    Private lastUpdateTime As DateTime = DateTime.MinValue
    Private lastMailEntryID As String = String.Empty
    Private Const UpdateThreshold As Integer = 500 ' 毫秒
    Private isUpdating As Boolean = False

    ' 添加主题相关变量
    Private currentTheme As Integer = -1

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' 初始化Application对象已在Designer中完成

        ' 获取当前Outlook主题
        GetCurrentOutlookTheme()

        ' 初始化Explorer窗口的任务窗格
        currentExplorer = Me.Application.ActiveExplorer
        InitializeMailPane()

        ' 初始化Inspectors集合并添加事件处理
        inspectors = Me.Application.Inspectors
        AddHandler inspectors.NewInspector, AddressOf Inspectors_NewInspector

        ' 处理已经打开的Inspector窗口
        For Each inspector As Outlook.Inspector In inspectors
            Inspectors_NewInspector(inspector)
        Next

        ' 初始化任务监视器
        taskMonitor = New TaskMonitor()
        taskMonitor.Initialize()

        ' 添加主题变化监听
        AddHandler SystemEvents.UserPreferenceChanged, AddressOf SystemEvents_UserPreferenceChanged
    End Sub

    Private Sub InitializeMailPane()
        mailThreadPane = New MailThreadPane()
        customTaskPane = Me.CustomTaskPanes.Add(mailThreadPane, "相关邮件v1")
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

        ' 注销Inspectors事件处理程序
        If inspectors IsNot Nothing Then
            RemoveHandler inspectors.NewInspector, AddressOf Inspectors_NewInspector
        End If

        ' 移除主题变化监听
        RemoveHandler SystemEvents.UserPreferenceChanged, AddressOf SystemEvents_UserPreferenceChanged

        ' 清理所有Inspector任务窗格
        For Each taskPane In inspectorTaskPanes.Values
            If taskPane IsNot Nothing Then
                taskPane.Dispose()
            End If
        Next
        inspectorTaskPanes.Clear()

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
    ' 处理新打开的Inspector窗口
    Private Sub Inspectors_NewInspector(inspector As Outlook.Inspector)
        Try
            ' 检查Inspector是否包含MailItem
            Dim mailItem As Object = inspector.CurrentItem
            If mailItem Is Nothing Then Return

            ' 为Inspector创建任务窗格
            Dim inspectorPane As New MailThreadPane()
            Dim inspectorTaskPane As Microsoft.Office.Tools.CustomTaskPane = Me.CustomTaskPanes.Add(inspectorPane, "相关邮件v1", inspector)
            inspectorTaskPane.Width = 400
            inspectorTaskPane.Visible = True

            ' 存储任务窗格引用
            Dim inspectorId As String = inspector.Caption & DateTime.Now.Ticks.ToString()
            inspectorTaskPanes(inspectorId) = inspectorTaskPane

            ' 添加Inspector关闭事件处理
            AddHandler CType(inspector, Outlook.InspectorEvents_Event).Close, Sub() InspectorClose(inspectorId)

            ' 添加Inspector当前项变化事件处理
            AddHandler CType(inspector, Outlook.InspectorEvents_Event).Activate, Sub() InspectorActivate(inspector, inspectorPane)

            ' 更新邮件列表
            UpdateInspectorMailContent(mailItem, inspectorPane)
        Catch ex As Exception
            Debug.WriteLine($"创建Inspector任务窗格出错: {ex.Message}")
        End Try
    End Sub

    ' 处理Inspector激活事件
    Private Sub InspectorActivate(inspector As Outlook.Inspector, inspectorPane As MailThreadPane)
        Try
            Dim mailItem As Object = inspector.CurrentItem
            If mailItem IsNot Nothing Then
                UpdateInspectorMailContent(mailItem, inspectorPane)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Inspector激活事件处理出错: {ex.Message}")
        End Try
    End Sub

    ' 处理Inspector关闭事件
    Private Sub InspectorClose(inspectorId As String)
        Try
            If inspectorTaskPanes.ContainsKey(inspectorId) Then
                Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = inspectorTaskPanes(inspectorId)
                If taskPane IsNot Nothing Then
                    taskPane.Dispose()
                End If
                inspectorTaskPanes.Remove(inspectorId)
            End If
        Catch ex As Exception
            Debug.WriteLine($"关闭Inspector任务窗格出错: {ex.Message}")
        End Try
    End Sub

    ' 更新Inspector窗口中的邮件内容
    Private Sub UpdateInspectorMailContent(item As Object, inspectorPane As MailThreadPane)
        Try
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

            ' 更新邮件列表
            If Not String.IsNullOrEmpty(mailEntryID) Then
                inspectorPane.UpdateMailList(conversationID, mailEntryID)
            End If
        Catch ex As Exception
            Debug.WriteLine($"更新Inspector邮件内容出错: {ex.Message}")
        End Try
    End Sub
    ' 获取当前Outlook主题
    Private Sub GetCurrentOutlookTheme()
        Try
            ' 尝试从注册表获取Outlook主题设置
            Dim themeValue As Integer = -1
            Dim outlookVersion As String = Application.Version
            Dim registryPath As String = "Software\\Microsoft\\Office\" & outlookVersion.Substring(0, 2) & ".0\\Common"

            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(registryPath)
                If key IsNot Nothing Then
                    Dim value As Object = key.GetValue("UI Theme")
                    If value IsNot Nothing Then
                        themeValue = Convert.ToInt32(value)
                    End If
                End If
            End Using

            ' 如果注册表中没有找到，使用默认值0（浅色/彩色主题）
            If themeValue = -1 Then
                themeValue = 0
            End If

            ' 如果主题发生变化，更新UI
            If themeValue <> currentTheme Then
                currentTheme = themeValue
                ApplyThemeToControls()
            End If

            Debug.WriteLine($"当前Outlook主题: {currentTheme}")
        Catch ex As Exception
            Debug.WriteLine($"获取Outlook主题出错: {ex.Message}")
        End Try
    End Sub

    ' 系统主题变化事件处理
    Private Sub SystemEvents_UserPreferenceChanged(sender As Object, e As UserPreferenceChangedEventArgs)
        If e.Category = UserPreferenceCategory.Color Then
            GetCurrentOutlookTheme()
        End If
    End Sub

    ' 应用主题到控件
    Private Sub ApplyThemeToControls()
        Try
            ' 根据Outlook版本和主题值设置颜色
            Dim backgroundColor As Color
            Dim foregroundColor As Color
            Dim outlookVersion As String = Application.Version.Substring(0, 2)

            ' Outlook 2016及以上版本
            If Convert.ToInt32(outlookVersion) >= 16 Then
                Select Case currentTheme
                    Case 0 ' Colorful
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                    Case 1 ' Dark Gray
                        backgroundColor = Color.FromArgb(68, 68, 68)
                        foregroundColor = Color.White
                    Case 2 ' Black
                        backgroundColor = Color.FromArgb(32, 32, 32)
                        foregroundColor = Color.White
                    Case 3 ' White
                        backgroundColor = Color.White
                        foregroundColor = Color.Black
                    Case Else
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                End Select
            Else ' Outlook 2013及以下版本
                Select Case currentTheme
                    Case 0 ' White
                        backgroundColor = Color.White
                        foregroundColor = Color.Black
                    Case 1 ' Light Gray
                        backgroundColor = Color.FromArgb(240, 240, 240)
                        foregroundColor = Color.Black
                    Case 2 ' Dark Gray
                        backgroundColor = Color.FromArgb(68, 68, 68)
                        foregroundColor = Color.White
                    Case Else
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                End Select
            End If

            ' 应用颜色到主任务窗格
            If mailThreadPane IsNot Nothing Then
                mailThreadPane.ApplyTheme(backgroundColor, foregroundColor)
            End If

            ' 应用颜色到所有Inspector任务窗格
            For Each taskPane In inspectorTaskPanes.Values
                Dim inspectorPane As MailThreadPane = TryCast(taskPane.Control, MailThreadPane)
                If inspectorPane IsNot Nothing Then
                    inspectorPane.ApplyTheme(backgroundColor, foregroundColor)
                End If
            Next

            Debug.WriteLine($"已应用主题颜色: 背景={backgroundColor}, 前景={foregroundColor}")
        Catch ex As Exception
            Debug.WriteLine($"应用主题颜色出错: {ex.Message}")
        End Try
    End Sub
End Class
