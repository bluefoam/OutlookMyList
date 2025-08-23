Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Threading.Tasks
Imports Microsoft.Win32
Imports OutlookAddIn3.Utils

Public Class ThisAddIn
    Private WithEvents currentExplorer As Outlook.Explorer
    Private customTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private mailThreadPane As MailThreadPane
    Private taskMonitor As TaskMonitor

    ' 公共属性用于访问MailThreadPane实例
    Public ReadOnly Property MailThreadPaneInstance As MailThreadPane
        Get
            Return mailThreadPane
        End Get
    End Property

    ' 添加Inspector相关变量
    Private WithEvents inspectors As Outlook.Inspectors
    Private inspectorTaskPanes As New Dictionary(Of String, Microsoft.Office.Tools.CustomTaskPane)
    
    ' 添加Inspector防重复调用变量（以Inspector为粒度进行防重）
    Private inspectorUpdateHistory As New Dictionary(Of String, DateTime)
    Private Const InspectorUpdateThreshold As Integer = 1000 ' 毫秒，Inspector更新阈值更长

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
        customTaskPane = Me.CustomTaskPanes.Add(mailThreadPane, "相关邮件v1.1")
        customTaskPane.Width = 400
        customTaskPane.Visible = True

        ' 添加分页状态改变事件处理程序
        AddHandler mailThreadPane.PaginationEnabledChanged, AddressOf MailThreadPane_PaginationEnabledChanged

        ' 初始化后，检查是否有当前选中的邮件并加载内容
        If currentExplorer IsNot Nothing AndAlso currentExplorer.Selection.Count > 0 Then
            Dim currentItem As Object = currentExplorer.Selection(1)
            If Not mailThreadPane.IsWebViewUpdateSuppressed Then
                ' 延迟加载当前选中的邮件内容
                mailThreadPane.BeginInvoke(Sub() UpdateMailContent(currentItem))
            End If
        End If
    End Sub

    Private Sub currentExplorer_SelectionChange() Handles currentExplorer.SelectionChange
        If mailThreadPane Is Nothing OrElse customTaskPane Is Nothing OrElse Not customTaskPane.Visible Then Return

        If currentExplorer.Selection.Count > 0 Then
            Dim selection As Object = currentExplorer.Selection(1)
            ' 在抑制期间跳过更新，避免 ContactInfoList 构造时不断刷新 WebView
            If Not mailThreadPane.IsWebViewUpdateSuppressed Then
                ' 推迟处理，避免在事件回调中访问项目属性
                mailThreadPane.BeginInvoke(Sub() UpdateMailContent(selection))
            End If
        End If
    End Sub

    Public Sub ToggleTaskPane()
        If customTaskPane IsNot Nothing Then
            customTaskPane.Visible = Not customTaskPane.Visible
            ' 显示窗格时，获取当前选中项并更新内容
            If customTaskPane.Visible Then
                If currentExplorer IsNot Nothing AndAlso currentExplorer.Selection.Count > 0 Then
                    Dim currentItem As Object = currentExplorer.Selection(1)
                    If Not mailThreadPane.IsWebViewUpdateSuppressed Then
                        ' 使用 BeginInvoke 延迟处理，避免在事件回调中访问项目属性
                        mailThreadPane.BeginInvoke(Sub() UpdateMailContent(currentItem))
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub UpdateMailList()
        If mailThreadPane IsNot Nothing Then
            If currentExplorer IsNot Nothing AndAlso currentExplorer.Selection.Count > 0 Then
                Dim currentItem As Object = currentExplorer.Selection(1)
                mailThreadPane.BeginInvoke(Sub() UpdateMailContent(currentItem))
            End If
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

    ' 注释掉 ItemLoad 事件处理，避免会话邮件加载过程中的大量 COM 异常
    ' ItemLoad 事件在会话邮件批量加载时会被频繁触发，导致性能问题和异常日志
    ' 我们已通过 SelectionChange 和 Inspector 事件充分覆盖了邮件选择和打开的场景
    'Private Sub Application_ItemLoad(item As Object) Handles Application.ItemLoad
    '    ' 已禁用：避免会话邮件加载过程中的 COM 异常和性能问题
    'End Sub

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
                mailThreadPane.UpdateMailList(String.Empty, mail.EntryID)
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(item, Outlook.AppointmentItem)
                mailThreadPane.UpdateMailList(String.Empty, appointment.EntryID)
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

            ' 避免直接在事件处理程序中使用项目的属性，仅获取ID
            Try
                If TypeOf item Is Outlook.MailItem Then
                    Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                    mailEntryID = mail.EntryID
                    conversationID = mail.ConversationID
                ElseIf TypeOf item Is Outlook.AppointmentItem Then
                    Dim appointment As Outlook.AppointmentItem = DirectCast(item, Outlook.AppointmentItem)
                    mailEntryID = appointment.EntryID
                    conversationID = appointment.ConversationID
                ElseIf TypeOf item Is Outlook.MeetingItem Then
                    Dim meeting As Outlook.MeetingItem = DirectCast(item, Outlook.MeetingItem)
                    mailEntryID = meeting.EntryID
                    ' MeetingItem 没有 ConversationID 属性，保持为空
                ElseIf TypeOf item Is Outlook.TaskItem Then
                    Dim task As Outlook.TaskItem = DirectCast(item, Outlook.TaskItem)
                    mailEntryID = task.EntryID
                    ' TaskItem 没有 ConversationID 属性，保持为空
                ElseIf TypeOf item Is Outlook.ContactItem Then
                    Dim contact As Outlook.ContactItem = DirectCast(item, Outlook.ContactItem)
                    mailEntryID = contact.EntryID
                    ' ContactItem 没有 ConversationID 属性，保持为空
                End If
            Catch comEx As System.Runtime.InteropServices.COMException
                Debug.WriteLine($"COM异常：无法在当前事件处理程序中访问项目属性: {comEx.Message}")
                Return
            End Try

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

            ' 延迟执行更新操作，让当前事件处理程序完成
            If Not String.IsNullOrEmpty(mailEntryID) Then
                Task.Run(Sub()
                             Try
                                 ' 在新线程中调用更新方法
                                 mailThreadPane.UpdateMailList(conversationID, mailEntryID)
                             Catch ex As Exception
                                 Debug.WriteLine($"异步UpdateMailList调用错误: {ex.Message}")
                             Finally
                                 ' 线程安全地重置更新标志
                                 SyncLock Me
                                     isUpdating = False
                                 End SyncLock
                             End Try
                         End Sub)
            Else
                isUpdating = False
            End If

        Catch ex As Exception
            Debug.WriteLine($"UpdateMailContent error: {ex.Message}")
            isUpdating = False
        End Try
    End Sub
    ' 处理新打开的Inspector窗口
    Private Sub Inspectors_NewInspector(inspector As Outlook.Inspector)
        Try
            ' 检查Inspector是否包含MailItem
            Dim mailItem As Object = inspector.CurrentItem

            ' 为Inspector创建任务窗格
            Dim inspectorPane As New MailThreadPane()
            Dim inspectorTaskPane As Microsoft.Office.Tools.CustomTaskPane = Me.CustomTaskPanes.Add(inspectorPane, "相关邮件v1.1", inspector)
            inspectorTaskPane.Width = 400
            inspectorTaskPane.Visible = True

            ' 为该Inspector生成唯一标识并存储其任务窗格
            Dim inspectorId As String = inspector.Caption & "|" & inspector.GetHashCode().ToString()
            inspectorTaskPanes(inspectorId) = inspectorTaskPane

            ' Add Inspector close event handler
            AddHandler CType(inspector, Outlook.InspectorEvents_Event).Close, Sub() InspectorClose(inspectorId)
            ' Add Inspector current item change event handler
            AddHandler CType(inspector, Outlook.InspectorEvents_Event).Activate, Sub() InspectorActivate(inspector, inspectorPane, inspectorId)

            ' 初始化时更新一次（若未处于抑制状态）。使用 BeginInvoke 保证不在事件过程内直接访问 EntryID。
            If Not inspectorPane.IsWebViewUpdateSuppressed Then
                ' 检查控件句柄是否已创建，避免 BeginInvoke 异常
                If inspectorPane.IsHandleCreated Then
                    inspectorPane.BeginInvoke(Sub()
                                                 Try
                                                     UpdateInspectorMailContent(mailItem, inspectorPane)
                                                 Catch ex As System.Exception
                                                     Debug.WriteLine($"Inspector 初始更新异常: {ex.Message}")
                                                 End Try
                                             End Sub)
                Else
                    ' 句柄未创建时，延迟到句柄创建后再执行
                    AddHandler inspectorPane.HandleCreated, Sub()
                                                                Try
                                                                    UpdateInspectorMailContent(mailItem, inspectorPane)
                                                                Catch ex As System.Exception
                                                                    Debug.WriteLine($"Inspector 初始延迟更新异常: {ex.Message}")
                                                                End Try
                                                            End Sub
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error creating Inspector task pane: {ex.Message}")
        End Try
    End Sub

    ' Handle Inspector activate event
    Private Sub InspectorActivate(inspector As Outlook.Inspector, inspectorPane As MailThreadPane, inspectorId As String)
        Try
            Dim mailItem As Object = inspector.CurrentItem

            ' 根据 Inspector 粒度进行防抖：短时间内重复激活只更新一次
            Dim now As DateTime = DateTime.Now
            Dim lastTime As DateTime = DateTime.MinValue
            If inspectorUpdateHistory.TryGetValue(inspectorId, lastTime) Then
                If (now - lastTime).TotalMilliseconds < InspectorUpdateThreshold Then
                    Debug.WriteLine($"InspectorActivate: 跳过重复更新（{(now - lastTime).TotalMilliseconds}ms 内） InspectorId={inspectorId}")
                    Return
                End If
            End If
            inspectorUpdateHistory(inspectorId) = now

            ' 抑制期间不进行内容更新，避免 ContactInfoList 构造时触发 WebView 刷新
            If inspectorPane Is Nothing OrElse inspectorPane.IsWebViewUpdateSuppressed Then Return
            
            ' 检查控件句柄是否已创建，避免 BeginInvoke 异常
            If inspectorPane.IsHandleCreated Then
                inspectorPane.BeginInvoke(Sub()
                                             Try
                                                 UpdateInspectorMailContent(mailItem, inspectorPane)
                                             Catch ex As System.Exception
                                                 Debug.WriteLine($"Inspector Activate 更新异常: {ex.Message}")
                                             End Try
                                         End Sub)
            Else
                ' 句柄未创建时，延迟到句柄创建后再执行
                AddHandler inspectorPane.HandleCreated, Sub()
                                                            Try
                                                                UpdateInspectorMailContent(mailItem, inspectorPane)
                                                            Catch ex As System.Exception
                                                                Debug.WriteLine($"Inspector 延迟更新异常: {ex.Message}")
                                                            End Try
                                                        End Sub
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error handling Inspector activate event: {ex.Message}")
        End Try
    End Sub

    ' Update mail content in Inspector window
    Private Sub UpdateInspectorMailContent(item As Object, inspectorPane As MailThreadPane)
        Try
            ' 抑制期间跳过更新，避免联系人信息列表构造时刷新
            If inspectorPane Is Nothing OrElse inspectorPane.IsWebViewUpdateSuppressed Then Return

            Dim mailEntryID As String = String.Empty
            Dim conversationID As String = String.Empty

            ' 仅读取 EntryID，避免在事件回调后立即访问其他属性
            If TypeOf item Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                mailEntryID = mail.EntryID
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(item, Outlook.AppointmentItem)
                mailEntryID = appointment.EntryID
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

            ' Update mail list
            If Not String.IsNullOrEmpty(mailEntryID) Then
                inspectorPane.UpdateMailList(conversationID, mailEntryID)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error updating Inspector mail content: {ex.Message}")
        End Try
    End Sub
    ' Get current Outlook theme
    Private Sub GetCurrentOutlookTheme()
        Try
            ' Attempt to get Outlook theme settings from the registry
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

            ' If not found in registry, use default value 0 (light/colorful theme)
            If themeValue = -1 Then
                themeValue = 0
            End If

            ' If theme changed, update UI
            If themeValue <> currentTheme Then
                currentTheme = themeValue
                ApplyThemeToControls()
            End If

            Debug.WriteLine($"Current Outlook theme: {currentTheme}")
        Catch ex As Exception
            Debug.WriteLine($"Error getting Outlook theme: {ex.Message}")
        End Try
    End Sub

    ' System theme change event handler
    Private Sub SystemEvents_UserPreferenceChanged(sender As Object, e As UserPreferenceChangedEventArgs)
        If e.Category = UserPreferenceCategory.Color Then
            GetCurrentOutlookTheme()
        End If
    End Sub

    ' Apply theme to controls
    Private Sub ApplyThemeToControls()
        Try
            ' Set colors based on Outlook version and theme value
            Dim backgroundColor As Color
            Dim foregroundColor As Color
            Dim outlookVersion As String = Application.Version.Substring(0, 2)

            ' Outlook 2016 and above
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
            Else ' Outlook 2013 and below
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

            ' Apply colors to main task pane
            If mailThreadPane IsNot Nothing Then
                mailThreadPane.ApplyTheme(backgroundColor, foregroundColor)
            End If

            ' Apply colors to all Inspector task panes
            For Each taskPane In inspectorTaskPanes.Values
                Dim inspectorPane As MailThreadPane = TryCast(taskPane.Control, MailThreadPane)
                If inspectorPane IsNot Nothing Then
                    inspectorPane.ApplyTheme(backgroundColor, foregroundColor)
                End If
            Next

            Debug.WriteLine($"Theme colors applied: Background={{backgroundColor}}, Foreground={{foregroundColor}}")
        Catch ex As Exception
            Debug.WriteLine($"Error applying theme colors: {ex.Message}")
        End Try
    End Sub
    ' Handle Inspector close event
    Private Sub InspectorClose(inspectorId As String)
        Try
            If inspectorTaskPanes.ContainsKey(inspectorId) Then
                Dim taskPane As Microsoft.Office.Tools.CustomTaskPane = inspectorTaskPanes(inspectorId)
                taskPane?.Dispose()
                inspectorTaskPanes.Remove(inspectorId)
            End If
            ' 清理Inspector防抖记录
            If inspectorUpdateHistory.ContainsKey(inspectorId) Then
                inspectorUpdateHistory.Remove(inspectorId)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error closing Inspector task pane: {ex.Message}")
        End Try
    End Sub

    ' 处理MailThreadPane分页状态改变事件
    Private Sub MailThreadPane_PaginationEnabledChanged(enabled As Boolean)
        Try
            ' 更新Ribbon按钮状态
            Dim ribbon As OutlookRibbon = TryCast(Me.Application.ActiveExplorer()?.CommandBars.GetRibbonX(), OutlookRibbon)
            If ribbon IsNot Nothing Then
                ribbon.UpdatePaginationButtonState(enabled)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error updating pagination button state: {ex.Message}")
        End Try
    End Sub
End Class
