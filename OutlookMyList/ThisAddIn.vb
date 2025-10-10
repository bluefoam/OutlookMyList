Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Threading.Tasks
Imports Microsoft.Win32
Imports System.Timers
Imports System.IO
' 添加DirectMergeHelper引用
Imports OutlookMyList

Public Class ThisAddIn
    Private WithEvents currentExplorer As Outlook.Explorer
    Private customTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private mailThreadPane As MailThreadPane
    Private taskMonitor As TaskMonitor

    ' 兼容遗留引用：底部/嵌入面板占位字段（功能已移除，保持为 Nothing）
    Private bottomPaneTaskPane As Microsoft.Office.Tools.CustomTaskPane = Nothing
    Private bottomPane As BottomPane = Nothing
    Private embeddedBottomPane As EmbeddedBottomPane = Nothing
    Private embeddedPaneForm As Form = Nothing

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
    Private WithEvents themeMonitorTimer As System.Timers.Timer

    ' 全局缓存开关
    Public Property CacheEnabled As Boolean = True
    
    ' 添加CommandBar相关变量
    Private mergeConversationButton As Microsoft.Office.Core.CommandBarButton

    ' 已移除：底部面板功能

    ' 取消自定义Ribbon XML覆盖，改用Ribbon设计器以确保工具栏稳定显示
    '（经典桌面版Outlook的右键菜单不稳定，故不使用IRibbonExtensibility返回XML）

    ' 已移除：嵌入式底部面板功能

    Private Sub InitializeCommandBars()
        Try
            LogInfo("初始化CommandBars: 开始")
            ' 确保ActiveExplorer存在
            If Me.Application.ActiveExplorer Is Nothing Then
                Debug.WriteLine("ActiveExplorer为空，延迟初始化CommandBar")
                ' 延迟初始化
                System.Threading.Tasks.Task.Delay(1000).ContinueWith(Sub() InitializeCommandBars())
                Return
            End If
            
            ' 获取Outlook的主窗口CommandBar
            Dim commandBars As Microsoft.Office.Core.CommandBars = Me.Application.ActiveExplorer.CommandBars
            
            ' 调试：列出所有可用的CommandBar
            Debug.WriteLine("=== 可用的CommandBar列表 ===")
            LogInfo("列出当前可用的CommandBar")
            For Each bar As Microsoft.Office.Core.CommandBar In commandBars
                Try
                    Debug.WriteLine($"CommandBar名称: '{bar.Name}', 类型: {bar.Type}, 可见: {bar.Visible}, 启用: {bar.Enabled}")
                Catch ex As Exception
                    Debug.WriteLine($"无法访问CommandBar属性: {ex.Message}")
                End Try
            Next
            Debug.WriteLine("=== CommandBar列表结束 ===")
            
            ' 查找邮件列表的右键菜单 - 尝试多个可能的名称
            Dim contextMenu As Microsoft.Office.Core.CommandBar = Nothing
            Dim possibleNames As String() = {"Item", "Context Menu", "Mail Item", "MailItem", "Message", "List View", "Table", "Reading Pane", "Folder List"}
            
            For Each name As String In possibleNames
                Try
                    contextMenu = commandBars(name)
                    Debug.WriteLine($"找到CommandBar: {name}")
                    LogInfo($"通过名称找到CommandBar: {name}")
                    Exit For
                Catch
                    Debug.WriteLine($"未找到CommandBar: {name}")
                    LogInfo($"未找到CommandBar: {name}")
                End Try
            Next
            
            ' 如果还是没有找到，尝试通过类型查找
            If contextMenu Is Nothing Then
                Debug.WriteLine("=== 尝试通过类型查找CommandBar ===")
                LogInfo("尝试通过类型查找CommandBar (msoBarTypePopup)")
                For Each bar As Microsoft.Office.Core.CommandBar In commandBars
                    Try
                        If bar.Type = Microsoft.Office.Core.MsoBarType.msoBarTypePopup Then
                            Debug.WriteLine($"发现弹出菜单: '{bar.Name}', 可见: {bar.Visible}")
                            LogInfo($"发现弹出菜单: '{bar.Name}', 可见: {bar.Visible}")
                            If bar.Name.ToLower().Contains("item") OrElse 
                               bar.Name.ToLower().Contains("mail") OrElse
                               bar.Name.ToLower().Contains("list") OrElse
                               bar.Name.ToLower().Contains("context") Then
                                contextMenu = bar
                                Debug.WriteLine($"*** 通过类型找到CommandBar: '{bar.Name}' ***")
                                LogInfo($"通过类型找到CommandBar: '{bar.Name}'")
                                Exit For
                            End If
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"检查CommandBar时出错: {ex.Message}")
                    End Try
                Next
                Debug.WriteLine("=== 类型查找结束 ===")
                LogInfo("类型查找结束")
            End If
            
            If contextMenu Is Nothing Then
                Debug.WriteLine("*** 警告: 无法找到邮件项目上下文菜单 ***")
                Debug.WriteLine("可能的原因: 1) Outlook版本不同 2) 需要等待用户操作 3) 权限问题")
                LogInfo("无法找到邮件项目上下文菜单")
                Return
            Else
                Debug.WriteLine($"*** 成功找到目标CommandBar: '{contextMenu.Name}' ***")
                LogInfo($"成功找到目标CommandBar: '{contextMenu.Name}'")
            End If
            
            ' 检查是否已经添加过菜单项
            For Each control As Microsoft.Office.Core.CommandBarControl In contextMenu.Controls
                If control.Tag = "MergeConversations" Then
                    Debug.WriteLine("菜单项已存在，跳过添加")
                    Return
                End If
            Next
            
            ' 添加合并会话菜单项
            mergeConversationButton = CType(contextMenu.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, , , , True), Microsoft.Office.Core.CommandBarButton)
            mergeConversationButton.Caption = "合并自定义会话"
            mergeConversationButton.Tag = "MergeConversations"
            mergeConversationButton.FaceId = 1000 ' 使用一个通用图标
            mergeConversationButton.Visible = True
            mergeConversationButton.Enabled = True
            mergeConversationButton.BeginGroup = True ' 添加分组分隔符
            
            ' 绑定点击事件
            AddHandler mergeConversationButton.Click, AddressOf MergeConversationButton_Click
            
            Debug.WriteLine($"成功添加菜单项到CommandBar: {contextMenu.Name}")
            LogInfo($"成功添加菜单项到CommandBar: {contextMenu.Name}")
            
        Catch ex As Exception
            Debug.WriteLine("初始化CommandBar失败: " & ex.Message)
            Debug.WriteLine($"错误详情: {ex.StackTrace}")
            LogException(ex, "InitializeCommandBars")
        End Try
    End Sub

    ' Ribbon XML 回调复用：在上下文菜单中执行合并逻辑
    Public Sub HandleMergeCustomConversation()
        Try
            ' 获取当前选中的邮件
            If currentExplorer Is Nothing OrElse currentExplorer.Selection.Count = 0 Then
                MessageBox.Show("请先选择要合并的邮件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 检查是否选择了多个邮件
            If currentExplorer.Selection.Count < 2 Then
                MessageBox.Show("请选择至少两个邮件进行合并。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 确认操作
            Dim result As System.Windows.Forms.DialogResult = MessageBox.Show(
                $"确定要将选中的 {currentExplorer.Selection.Count} 个邮件合并到同一个自定义会话中吗？" & vbCrLf & vbCrLf &
                "第一个选中的邮件的会话ID将作为目标会话ID。",
                "确认合并会话",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If result = System.Windows.Forms.DialogResult.No Then
                Return
            End If

            ' 执行合并操作
            Dim targetConversationId As String = ""
            Dim processedCount As Integer = 0
            Dim errorCount As Integer = 0

            ' 首先检查所有选中的邮件，查找是否有任何一个已存在自定义会话ID
            Dim foundCustomId As Boolean = False
            For i As Integer = 1 To currentExplorer.Selection.Count
                Try
                    Dim mailItem As Object = currentExplorer.Selection(i)
                    If TypeOf mailItem Is Outlook.MailItem OrElse TypeOf mailItem Is Outlook.AppointmentItem OrElse TypeOf mailItem Is Outlook.MeetingItem Then
                        Dim customId As String = mailThreadPane.ReadCustomConversationIdFromItem(mailItem)
                        If Not String.IsNullOrEmpty(customId) Then
                            targetConversationId = customId
                            foundCustomId = True
                            Exit For
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"检查邮件 {i} 的自定义会话ID时出错: {ex.Message}")
                End Try
            Next

            ' 如果没有找到任何自定义会话ID，则使用第一个邮件的原始ConversationID
            If Not foundCustomId Then
                Dim firstMailItem As Object = currentExplorer.Selection(1)
                If TypeOf firstMailItem Is Outlook.MailItem OrElse TypeOf firstMailItem Is Outlook.AppointmentItem OrElse TypeOf firstMailItem Is Outlook.MeetingItem Then
                    If TypeOf firstMailItem Is Outlook.MailItem Then
                        targetConversationId = DirectCast(firstMailItem, Outlook.MailItem).ConversationID
                    ElseIf TypeOf firstMailItem Is Outlook.AppointmentItem Then
                        targetConversationId = DirectCast(firstMailItem, Outlook.AppointmentItem).ConversationID
                    ElseIf TypeOf firstMailItem Is Outlook.MeetingItem Then
                        targetConversationId = DirectCast(firstMailItem, Outlook.MeetingItem).ConversationID
                    End If
                End If
            End If

            If String.IsNullOrEmpty(targetConversationId) Then
                MessageBox.Show("无法确定目标会话ID。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 显示进度
            mailThreadPane.ShowProgress("正在合并会话...", currentExplorer.Selection.Count)

            ' 执行合并操作
            For i As Integer = 1 To currentExplorer.Selection.Count
                Try
                    Dim mailItem As Object = currentExplorer.Selection(i)

                    ' 更新进度
                    mailThreadPane.UpdateProgress(i, $"正在处理第 {i} 个邮件...")
                    System.Windows.Forms.Application.DoEvents() ' 允许UI更新

                    ' 检查是否为支持的邮件类型
                    If TypeOf mailItem Is Outlook.MailItem OrElse TypeOf mailItem Is Outlook.AppointmentItem OrElse TypeOf mailItem Is Outlook.MeetingItem Then
                        Dim entryId As String = ""

                        ' 获取EntryID
                        If TypeOf mailItem Is Outlook.MailItem Then
                            entryId = DirectCast(mailItem, Outlook.MailItem).EntryID
                        ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                            entryId = DirectCast(mailItem, Outlook.AppointmentItem).EntryID
                        ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                            entryId = DirectCast(mailItem, Outlook.MeetingItem).EntryID
                        End If

                        ' 设置自定义会话ID
                        If Not String.IsNullOrEmpty(entryId) Then
                            ' 获取 StoreID 以确保跨邮箱检索
                            Dim storeId As String = Nothing
                            Try
                                Dim parentFolder As Outlook.MAPIFolder = Nothing
                                If TypeOf mailItem Is Outlook.MailItem Then
                                    parentFolder = TryCast(DirectCast(mailItem, Outlook.MailItem).Parent, Outlook.MAPIFolder)
                                ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                                    parentFolder = TryCast(DirectCast(mailItem, Outlook.AppointmentItem).Parent, Outlook.MAPIFolder)
                                ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                                    parentFolder = TryCast(DirectCast(mailItem, Outlook.MeetingItem).Parent, Outlook.MAPIFolder)
                                End If
                                If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                                    storeId = parentFolder.Store.StoreID
                                End If
                            Catch
                            End Try

                            Debug.WriteLine($"尝试设置邮件 {i} 的自定义会话ID: entryId={entryId}, targetConversationId={targetConversationId}")
                            Dim setResult = mailThreadPane.SetCustomConversationIdByEntryID(entryId, targetConversationId, storeId)
                            If setResult Then
                                Debug.WriteLine($"成功设置邮件 {i} 的自定义会话ID")
                                processedCount += 1
                            Else
                                Debug.WriteLine($"设置邮件 {i} 的自定义会话ID失败")
                                errorCount += 1
                            End If
                        Else
                            errorCount += 1
                        End If
                    Else
                        errorCount += 1
                    End If

                Catch ex As Exception
                    errorCount += 1
                    Debug.WriteLine($"处理邮件 {i} 时出错: {ex.Message}")
                End Try
            Next

            ' 隐藏进度条
            mailThreadPane.HideProgress()

            ' 显示结果
            Dim message As String = $"合并完成！" & vbCrLf &
                                  $"成功处理: {processedCount} 个邮件" & vbCrLf &
                                  $"失败: {errorCount} 个邮件"

            If errorCount > 0 Then
                message &= vbCrLf & vbCrLf & "部分邮件可能由于权限或其他原因无法修改。"
            End If

            MessageBox.Show(message, "合并结果", MessageBoxButtons.OK, 
                          If(errorCount = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))

            ' 强制刷新邮件列表
            If mailThreadPane IsNot Nothing AndAlso currentExplorer IsNot Nothing AndAlso currentExplorer.Selection.Count > 0 Then
                Dim currentItem As Object = currentExplorer.Selection(1)
                If TypeOf currentItem Is Outlook.MailItem Then
                    Dim currentMail As Outlook.MailItem = CType(currentItem, Outlook.MailItem)
                    mailThreadPane.UpdateMailList(targetConversationId, currentMail.EntryID)
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    Dim currentAppt As Outlook.AppointmentItem = CType(currentItem, Outlook.AppointmentItem)
                    mailThreadPane.UpdateMailList(targetConversationId, currentAppt.EntryID)
                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                    Dim currentMeeting As Outlook.MeetingItem = CType(currentItem, Outlook.MeetingItem)
                    mailThreadPane.UpdateMailList(targetConversationId, currentMeeting.EntryID)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"合并会话时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Debug.WriteLine($"HandleMergeCustomConversation错误: {ex.Message}")
            Try
                LogException(ex, "HandleMergeCustomConversation")
            Catch
            End Try
        End Try
    End Sub

    Private Sub MergeConversationButton_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            ' 获取当前选中的邮件
            If currentExplorer Is Nothing OrElse currentExplorer.Selection.Count = 0 Then
                MessageBox.Show("请先选择要合并的邮件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            
            ' 检查是否选择了多个邮件
            If currentExplorer.Selection.Count < 2 Then
                MessageBox.Show("请选择至少两个邮件进行合并。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If
            
            ' 确认操作
            Dim dialogResult As System.Windows.Forms.DialogResult = MessageBox.Show(
                $"确定要将选中的 {currentExplorer.Selection.Count} 个邮件合并到同一个自定义会话中吗？" & vbCrLf & vbCrLf &
                "系统将优先使用已有的自定义会话ID，如果没有则使用第一个邮件的会话ID。",
                "确认合并会话",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)
            
            If dialogResult <> System.Windows.Forms.DialogResult.Yes Then
                Return
            End If
            
            ' 显示进度
            mailThreadPane.ShowProgress("正在合并会话...", currentExplorer.Selection.Count)
            
            ' 使用新的DirectMergeHelper类执行合并操作
            Dim mergeInfo = OutlookMyList.DirectMergeHelper.MergeConversation(currentExplorer.Selection)
            
            ' 隐藏进度条
            mailThreadPane.HideProgress()
            
            ' 显示结果
            If mergeInfo.success Then
                Dim message As String = $"合并完成！" & vbCrLf &
                                      $"成功处理: {mergeInfo.processedCount} 个邮件" & vbCrLf &
                                      $"失败: {mergeInfo.errorCount} 个邮件"
    
                If mergeInfo.errorCount > 0 Then
                    message &= vbCrLf & vbCrLf & "部分邮件可能由于权限或其他原因无法修改。"
                End If
    
                MessageBox.Show(message, "合并结果", MessageBoxButtons.OK, 
                              If(mergeInfo.errorCount = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))
                
                ' 强制刷新邮件列表
                If Not String.IsNullOrEmpty(mergeInfo.targetConversationId) AndAlso currentExplorer.Selection.Count > 0 Then
                    Dim currentItem As Object = currentExplorer.Selection(1)
                    Dim entryId As String = String.Empty
                    
                    If TypeOf currentItem Is Outlook.MailItem Then
                        entryId = DirectCast(currentItem, Outlook.MailItem).EntryID
                    ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                        entryId = DirectCast(currentItem, Outlook.AppointmentItem).EntryID
                    ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                        entryId = DirectCast(currentItem, Outlook.MeetingItem).EntryID
                    End If
                    
                    If Not String.IsNullOrEmpty(entryId) Then
                        mailThreadPane.UpdateMailList(mergeInfo.targetConversationId, entryId)
                    End If
                End If
            Else
                MessageBox.Show("合并操作失败，无法确定目标会话ID。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            ' 记录错误并显示错误消息
            Debug.WriteLine($"MergeConversationButton_Click错误: {ex.Message}")
            MessageBox.Show($"合并会话时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            
            ' 隐藏进度条（如果已显示）
            Try
                mailThreadPane.HideProgress()
            Catch
                ' 忽略隐藏进度条时的错误
            End Try
        End Try
    End Sub

    Public Sub ToggleEmbeddedBottomPane()
        ' 功能已移除：不执行任何操作
        Return
    End Sub

    Public Sub MinimizeBottomPane()
        ' 功能已移除：不执行任何操作
    End Sub

    Public ReadOnly Property BottomPaneInstance As BottomPane
        Get
            ' 功能已移除：返回 Nothing
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property IsBottomPaneVisible As Boolean
        Get
            ' 功能已移除：始终返回 False
            Return False
        End Get
    End Property

    Public ReadOnly Property IsEmbeddedBottomPaneVisible As Boolean
        Get
            ' 功能已移除：始终返回 False
            Return False
        End Get
    End Property

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
        
        ' 清理主题监听定时器
        If themeMonitorTimer IsNot Nothing Then
            themeMonitorTimer.Stop()
            themeMonitorTimer.Dispose()
            themeMonitorTimer = Nothing
        End If
        
        ' 清理CommandBar
        Try
            If mergeConversationButton IsNot Nothing Then
                RemoveHandler mergeConversationButton.Click, AddressOf MergeConversationButton_Click
                mergeConversationButton.Delete()
                mergeConversationButton = Nothing
            End If
            
            ' 清理所有带有我们标签的控件
            If Me.Application.ActiveExplorer IsNot Nothing Then
                Dim commandBars As Microsoft.Office.Core.CommandBars = Me.Application.ActiveExplorer.CommandBars
                For Each bar As Microsoft.Office.Core.CommandBar In commandBars
                    Try
                        Dim controlsToDelete As New List(Of Microsoft.Office.Core.CommandBarControl)
                        For Each control As Microsoft.Office.Core.CommandBarControl In bar.Controls
                            If control.Tag = "MergeConversations" Then
                                controlsToDelete.Add(control)
                            End If
                        Next
                        
                        For Each control In controlsToDelete
                            control.Delete()
                        Next
                    Catch
                        ' 忽略清理错误
                    End Try
                Next
            End If
        Catch ex As Exception
            Debug.WriteLine("清理CommandBar时出错: " & ex.Message)
        End Try
        
        Debug.WriteLine("主题监听器已停止")

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
        
        ' 清理底部面板
        If bottomPaneTaskPane IsNot Nothing Then
            bottomPaneTaskPane.Dispose()
        End If
        If bottomPane IsNot Nothing Then
            bottomPane.Dispose()
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
                    conversationID = meeting.ConversationID
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

            ' 抑制期间不进行内容更新，避免 ContactInfoTree 构造时触发 WebView 刷新
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

    ' 初始化主题监听器
    Private Sub InitializeThemeMonitor()
        Try
            ' 创建定时器，每2秒检查一次主题变化
            themeMonitorTimer = New System.Timers.Timer(2000)
            themeMonitorTimer.AutoReset = True
            themeMonitorTimer.Enabled = True
            Debug.WriteLine("主题监听器已启动")
        Catch ex As Exception
            Debug.WriteLine($"初始化主题监听器失败: {ex.Message}")
        End Try
    End Sub

    ' 定时器事件处理程序
    Private Sub ThemeMonitorTimer_Elapsed(sender As Object, e As ElapsedEventArgs) Handles themeMonitorTimer.Elapsed
        Try
            ' 在UI线程上执行主题检查
            If mailThreadPane IsNot Nothing Then
                mailThreadPane.BeginInvoke(Sub() GetCurrentOutlookTheme())
            End If
        Catch ex As Exception
            Debug.WriteLine($"主题监听器检查失败: {ex.Message}")
        End Try
    End Sub

    ' Get current Outlook theme
    Private Sub GetCurrentOutlookTheme()
        Try
            ' Attempt to get Outlook theme settings from the registry
            Dim themeValue As Integer = -1
            Dim outlookVersion As String = Application.Version
            Dim majorVersion As String = outlookVersion.Substring(0, 2)
            
            ' 尝试多个可能的注册表路径
            Dim registryPaths As String() = {
                $"Software\\Microsoft\\Office\\{majorVersion}.0\\Common\\UI",
                $"Software\\Microsoft\\Office\\{majorVersion}.0\\Common",
                $"Software\\Microsoft\\Office\\{majorVersion}.0\\Outlook\\Options\\General"
            }
            
            Debug.WriteLine($"Outlook版本: {outlookVersion}, 主版本: {majorVersion}")
            
            For Each registryPath As String In registryPaths
                Debug.WriteLine($"尝试注册表路径: {registryPath}")
                
                Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(registryPath)
                    If key IsNot Nothing Then
                        Debug.WriteLine($"成功打开注册表键: {registryPath}")
                        
                        ' 尝试多个可能的键名
                        Dim keyNames As String() = {"UI Theme", "Theme", "UITheme"}
                        
                        For Each keyName As String In keyNames
                            Dim value As Object = key.GetValue(keyName)
                            If value IsNot Nothing Then
                                themeValue = SafeParseTheme(value)
                                If themeValue <> -1 Then
                                    Debug.WriteLine($"找到主题值: {keyName} = {themeValue}")
                                    LogInfo($"Outlook主题解析: {keyName} = {themeValue}")
                                    Exit For
                                End If
                            End If
                        Next
                        
                        If themeValue <> -1 Then Exit For
                    Else
                        Debug.WriteLine($"无法打开注册表键: {registryPath}")
                    End If
                End Using
            Next

            ' If not found in registry, use default value 0 (light/colorful theme)
            If themeValue = -1 Then
                themeValue = 0
                Debug.WriteLine("未找到主题设置，使用默认值 0")
            End If

            ' If theme changed, update UI
            If themeValue <> currentTheme Then
                Debug.WriteLine($"主题变化: {currentTheme} -> {themeValue}")
                currentTheme = themeValue
                ApplyThemeToControls()
            Else
                ' 即使主题值没有变化，也要确保在启动时应用主题
                Debug.WriteLine($"主题值未变化，但强制应用主题以确保正确初始化: {themeValue}")
                ApplyThemeToControls()
            End If

            Debug.WriteLine($"当前Outlook主题: {currentTheme}")
        Catch ex As Exception
            Debug.WriteLine($"获取Outlook主题时出错: {ex.Message}")
        End Try
    End Sub

    ' 测试方法：手动触发主题检测
    Public Sub TestThemeDetection()
        Debug.WriteLine("=== 开始手动主题检测测试 ===")
        GetCurrentOutlookTheme()
        Debug.WriteLine("=== 主题检测测试完成 ===")
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
                        Debug.WriteLine("应用主题: Colorful (浅色)")
                    Case 1 ' Dark Gray
                        backgroundColor = Color.FromArgb(68, 68, 68)
                        foregroundColor = Color.White
                        Debug.WriteLine("应用主题: Dark Gray (深灰)")
                    Case 2 ' Black
                        backgroundColor = Color.FromArgb(32, 32, 32)
                        foregroundColor = Color.White
                        Debug.WriteLine("应用主题: Black (黑色)")
                    Case 3 ' White
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                        Debug.WriteLine("应用主题: White (白色) - 使用系统颜色")
                    Case 4 ' Dark Mode (新版本的黑色主题)
                        backgroundColor = Color.FromArgb(32, 32, 32)
                        foregroundColor = Color.White
                        Debug.WriteLine("应用主题: Dark Mode (深色模式)")
                    Case 5 ' System theme
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                        Debug.WriteLine("应用主题: System (系统主题)")
                    Case Else
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                        Debug.WriteLine($"应用主题: 未知主题值 {currentTheme}，使用默认浅色")
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
                Debug.WriteLine($"[ApplyThemeToControls] 调用 mailThreadPane.ApplyTheme，背景色: {backgroundColor}, 前景色: {foregroundColor}")
                Debug.WriteLine($"[ApplyThemeToControls] 调用前全局变量值: 背景={MailThreadPane.globalThemeBackgroundColor}, 前景={MailThreadPane.globalThemeForegroundColor}")
                mailThreadPane.ApplyTheme(backgroundColor, foregroundColor)
                Debug.WriteLine($"[ApplyThemeToControls] 调用后全局变量值: 背景={MailThreadPane.globalThemeBackgroundColor}, 前景={MailThreadPane.globalThemeForegroundColor}")
            Else
                Debug.WriteLine("[ApplyThemeToControls] 警告: mailThreadPane 为 Nothing")
            End If

            ' Apply colors to all Inspector task panes
            For Each taskPane In inspectorTaskPanes.Values
                Dim inspectorPane As MailThreadPane = TryCast(taskPane.Control, MailThreadPane)
                If inspectorPane IsNot Nothing Then
                    inspectorPane.ApplyTheme(backgroundColor, foregroundColor)
                End If
            Next

            ' Apply colors to bottom pane
            If bottomPane IsNot Nothing Then
                bottomPane.ApplyTheme(backgroundColor, foregroundColor)
            End If

            ' Apply colors to embedded bottom pane
            If embeddedBottomPane IsNot Nothing Then
                embeddedBottomPane.ApplyTheme(backgroundColor, foregroundColor)
            End If

            Debug.WriteLine($"Theme colors applied: Background={{backgroundColor}}, Foreground={{foregroundColor}}")
        Catch ex As Exception
            Debug.WriteLine($"Error applying theme colors: {ex.Message}")
        End Try
    End Sub

    ' 公共方法：获取当前主题颜色
    ' 公共方法：强制刷新主题
    Public Sub RefreshTheme()
        Try
            GetCurrentOutlookTheme()
        Catch ex As Exception
            Debug.WriteLine($"RefreshTheme error: {ex.Message}")
        End Try
    End Sub

    Public Function GetCurrentThemeColors() As (backgroundColor As Color, foregroundColor As Color)
        Try
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
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
                    Case 4 ' Dark Mode
                        backgroundColor = Color.FromArgb(32, 32, 32)
                        foregroundColor = Color.White
                    Case 5 ' System theme
                        backgroundColor = SystemColors.Window
                        foregroundColor = SystemColors.WindowText
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

            Debug.WriteLine($"ThisAddIn.GetCurrentThemeColors: 主题={currentTheme}, 背景={backgroundColor}, 前景={foregroundColor}")
            Return (backgroundColor, foregroundColor)
        Catch ex As Exception
            Debug.WriteLine($"获取主题颜色失败: {ex.Message}")
            Return (SystemColors.Window, SystemColors.WindowText)
        End Try
    End Function
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
            ' 使用 Globals.Ribbons 访问设计器生成的 Ribbon 实例
            If Globals.Ribbons IsNot Nothing AndAlso Globals.Ribbons.Ribbon1 IsNot Nothing Then
                Globals.Ribbons.Ribbon1.UpdatePaginationButtonState(enabled)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error updating pagination button state: {ex.Message}")
        End Try
    End Sub

    ' 添加缺失的方法
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' 初始化Application对象已在Designer中完成

        ' 注册全局异常处理，避免未处理异常弹窗，并记录日志
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf GlobalUnhandledExceptionHandler
        AddHandler System.Windows.Forms.Application.ThreadException, AddressOf GlobalThreadExceptionHandler

        ' 从注册表加载缓存开关
        LoadCacheEnabledFromRegistry()

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
        
        ' 初始化CommandBar右键菜单 - 延迟执行以确保Outlook完全加载
        System.Threading.Tasks.Task.Delay(2000).ContinueWith(Sub() InitializeCommandBars())
    End Sub

    Private Sub GlobalUnhandledExceptionHandler(sender As Object, e As UnhandledExceptionEventArgs)
        Try
            Dim ex = TryCast(e.ExceptionObject, Exception)
            LogException(ex, "Unhandled")
        Catch
        End Try
    End Sub

    Private Sub GlobalThreadExceptionHandler(sender As Object, e As Threading.ThreadExceptionEventArgs)
        Try
            LogException(e.Exception, "Thread")
        Catch
        End Try
    End Sub

    Public Sub LogException(ex As Exception, prefix As String)
        Try
            Dim dir As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookMyList")
            System.IO.Directory.CreateDirectory(dir)
            Dim logPath As String = System.IO.Path.Combine(dir, "error.log")
            System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{prefix}] {ex?.ToString()}{Environment.NewLine}")
            Debug.WriteLine($"[{prefix}] {ex?.Message}")
        Catch
        End Try
    End Sub

    Public Sub LogInfo(message As String)
        Try
            Dim dir As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookMyList")
            System.IO.Directory.CreateDirectory(dir)
            Dim logPath As String = System.IO.Path.Combine(dir, "error.log")
            System.IO.File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [Info] {message}{Environment.NewLine}")
            Debug.WriteLine(message)
        Catch
        End Try
    End Sub

    Private Function SafeParseTheme(valueObj As Object) As Integer
        Try
            If valueObj Is Nothing Then Return -1
            If TypeOf valueObj Is Integer Then
                Return DirectCast(valueObj, Integer)
            ElseIf TypeOf valueObj Is String Then
                Dim s As String = DirectCast(valueObj, String)
                Dim tmp As Integer
                If Integer.TryParse(s, tmp) Then
                    Return tmp
                End If
                Select Case s.Trim().ToLower()
                    Case "colorful", "light", "white"
                        Return 0
                    Case "darkgray", "dark gray"
                        Return 1
                    Case "black", "dark"
                        Return 2
                    Case Else
                        Return -1
                End Select
            Else
                Return Convert.ToInt32(valueObj)
            End If
        Catch
            Return -1
        End Try
    End Function

    Private Sub currentExplorer_SelectionChange() Handles currentExplorer.SelectionChange
        If mailThreadPane Is Nothing OrElse customTaskPane Is Nothing OrElse Not customTaskPane.Visible Then Return

        If currentExplorer.Selection.Count > 0 Then
            Dim selection As Object = currentExplorer.Selection(1)
            ' 在抑制期间跳过更新，避免 ContactInfoTree 构造时不断刷新 WebView
            If Not mailThreadPane.IsWebViewUpdateSuppressed Then
                ' 推迟处理，避免在事件回调中访问项目属性
                mailThreadPane.BeginInvoke(Sub() UpdateMailContent(selection))
            End If
        End If

        ' 同步更新Ribbon中"合并自定义会话"按钮的启用状态（选择数≥2启用）
        Try
            Dim enabled As Boolean = (currentExplorer.Selection.Count >= 2)
            If Globals.Ribbons IsNot Nothing AndAlso Globals.Ribbons.Ribbon1 IsNot Nothing Then
                Globals.Ribbons.Ribbon1.UpdateMergeButtonState(enabled)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Error updating merge button state: {ex.Message}")
        End Try
    End Sub

    Private Sub InitializeMailPane()
        mailThreadPane = New MailThreadPane()
        customTaskPane = Me.CustomTaskPanes.Add(mailThreadPane, "相关邮件v1.1")
        customTaskPane.Width = 400
        customTaskPane.Visible = True

        ' 立即应用主题到新创建的MailThreadPane
        ApplyThemeToControls()

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

    Private Sub LoadCacheEnabledFromRegistry()
        Try
            Dim basePath As String = "Software\\OutlookMyList\\Settings"
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey(basePath)
                If key IsNot Nothing Then
                    Dim value As Object = key.GetValue("CacheEnabled", True)
                    Dim enabled As Boolean = True
                    If TypeOf value Is Integer Then
                        enabled = (DirectCast(value, Integer) <> 0)
                    ElseIf TypeOf value Is String Then
                        enabled = Boolean.TryParse(DirectCast(value, String), enabled)
                    ElseIf TypeOf value Is Boolean Then
                        enabled = DirectCast(value, Boolean)
                    End If
                    CacheEnabled = enabled
                    LogInfo($"加载缓存开关: {CacheEnabled}")
                Else
                    CacheEnabled = True
                End If
            End Using
        Catch ex As Exception
            Debug.WriteLine($"加载缓存开关失败: {ex.Message}")
            CacheEnabled = True
        End Try
    End Sub

    Public Sub ToggleTaskPane()
        Try
            If bottomPaneTaskPane IsNot Nothing Then
                bottomPaneTaskPane.Visible = Not bottomPaneTaskPane.Visible
            End If
        Catch ex As System.Exception
             LogException(ex, "ToggleTaskPane")
         End Try
    End Sub

    Public Sub SaveCacheEnabledToRegistry(enabled As Boolean)
        Try
            Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\OutlookMyList")
            key.SetValue("CacheEnabled", enabled)
            key.Close()
            CacheEnabled = enabled
        Catch ex As System.Exception
             LogException(ex, "SaveCacheEnabledToRegistry")
         End Try
    End Sub

End Class
