Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports OutlookAddIn3.Utils
Imports OutlookAddIn3.Models
Imports OutlookAddIn3.Handlers
Imports System.Drawing
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports System.Threading.Tasks
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Linq


<ComVisible(True)>
Public Class MailThreadPane
    Inherits UserControl





    ' 添加类级别的字体缓存
    Private ReadOnly iconFont As Font
    Private Shadows ReadOnly defaultFont As Font
    Private ReadOnly highlightFont As Font
    Private ReadOnly normalFont As Font
    Private ReadOnly highlightColor As Color = Color.FromArgb(255, 255, 200)

    ' MessageClass映射缓存 - 提高类型判断效率
    Private Shared ReadOnly MessageClassBaseIndex As New Dictionary(Of String, Integer) From {
        {"IPM.Note", 0},
        {"IPM.Appointment", 6},
        {"IPM.Schedule.Meeting", 6},
        {"IPM.Task", 12},
        {"IPM.Contact", 18}
    }

    ' 图标组合查找表 - 通过索引快速获取图标组合
    Private Shared ReadOnly IconCombinations As String() = {
        "📧",           ' 0: 邮件
        "📧📎",         ' 1: 邮件+附件
        "📧🚩",         ' 2: 邮件+进行中旗标
        "📧📎🚩",       ' 3: 邮件+附件+进行中旗标
        "📧⚑",         ' 4: 邮件+已完成旗标
        "📧📎⚑",       ' 5: 邮件+附件+已完成旗标
        "📅",           ' 6: 日历
        "📅📎",         ' 7: 日历+附件
        "📅🚩",         ' 8: 日历+进行中旗标
        "📅📎🚩",       ' 9: 日历+附件+进行中旗标
        "📅⚑",         ' 10: 日历+已完成旗标
        "📅📎⚑",       ' 11: 日历+附件+已完成旗标
        "📋",           ' 12: 任务
        "📋📎",         ' 13: 任务+附件
        "📋🚩",         ' 14: 任务+进行中旗标
        "📋📎🚩",       ' 15: 任务+附件+进行中旗标
        "📋⚑",         ' 16: 任务+已完成旗标
        "📋📎⚑",       ' 17: 任务+附件+已完成旗标
        "👤",           ' 18: 联系人
        "👤📎",         ' 19: 联系人+附件
        "👤🚩",         ' 20: 联系人+进行中旗标
        "👤📎🚩",       ' 21: 联系人+附件+进行中旗标
        "👤⚑",         ' 22: 联系人+已完成旗标
        "👤📎⚑"        ' 23: 联系人+附件+已完成旗标
    }

    ' 主题颜色
    Private currentBackColor As Color = SystemColors.Window
    Private currentForeColor As Color = SystemColors.WindowText

    ' 抑制在列表构造/填充时触发 WebView 刷新或加载的标志
    Private suppressWebViewUpdate As Integer = 0 ' 使用计数器以支持嵌套调用
    
    ' 暴露抑制状态以供外部检查
    Public ReadOnly Property IsWebViewUpdateSuppressed As Boolean
        Get
            Return suppressWebViewUpdate > 0
        End Get
    End Property

    ' 分页功能开关的私有字段
    Private _isPaginationEnabled As Boolean = False

    ' 分页状态改变事件
    Public Event PaginationEnabledChanged(enabled As Boolean)

    ' 分页功能开关属性
    Public Property IsPaginationEnabled As Boolean
        Get
            Return _isPaginationEnabled
        End Get
        Set(value As Boolean)
            If _isPaginationEnabled <> value Then
                _isPaginationEnabled = value
                Debug.WriteLine($"分页功能开关已{If(value, "启用", "禁用")}")
                ' 触发事件通知状态改变
                RaiseEvent PaginationEnabledChanged(_isPaginationEnabled)
                ' 如果当前有邮件列表，重新应用分页设置
                If allListViewItems IsNot Nothing AndAlso allListViewItems.Count > 0 Then
                    EnableVirtualMode(allListViewItems.Count)
                    ' 重新加载当前页面
                    If isVirtualMode Then
                        LoadPage(0)
                    Else
                        ' 非虚拟模式：显示所有项目
                        lvMails.BeginUpdate()
                        Try
                            lvMails.Items.Clear()
                            mailItems.Clear()
                            For i As Integer = 0 To allListViewItems.Count - 1
                                Dim item = allListViewItems(i)
                                Dim clonedItem = CType(item.Clone(), ListViewItem)
                                lvMails.Items.Add(clonedItem)
                                mailItems.Add((i, ConvertEntryIDToString(item.Tag)))
                            Next
                        Finally
                            lvMails.EndUpdate()
                        End Try
                    End If
                    UpdatePaginationUI()
                End If
            End If
        End Set
    End Property

    ' 切换分页功能开关的便捷方法
    Public Sub TogglePagination()
        IsPaginationEnabled = Not IsPaginationEnabled
    End Sub

    ' 应用主题颜色
    Public Sub ApplyTheme(backgroundColor As Color, foregroundColor As Color)
        Try
            ' 保存当前主题颜色
            currentBackColor = backgroundColor
            currentForeColor = foregroundColor

            ' 应用到控件
            Me.BackColor = backgroundColor

            ' 应用到ListView
            If lvMails IsNot Nothing Then
                lvMails.BackColor = backgroundColor
                lvMails.ForeColor = foregroundColor
            End If

            ' 应用到任务列表
            If taskList IsNot Nothing Then
                taskList.BackColor = backgroundColor
                taskList.ForeColor = foregroundColor
            End If

            ' 应用到分隔控件
            If splitter1 IsNot Nothing Then
                splitter1.BackColor = backgroundColor
                splitter1.Panel1.BackColor = backgroundColor
                splitter1.Panel2.BackColor = backgroundColor
            End If

            If splitter2 IsNot Nothing Then
                splitter2.BackColor = backgroundColor
                splitter2.Panel1.BackColor = backgroundColor
                splitter2.Panel2.BackColor = backgroundColor
            End If



            ' 应用到按钮面板
            If btnPanel IsNot Nothing Then
                btnPanel.BackColor = backgroundColor

                ' 应用到按钮面板中的所有控件
                For Each ctrl As Control In btnPanel.Controls
                    If TypeOf ctrl Is Button Then
                        ' 按钮保持系统默认颜色
                    Else
                        ctrl.BackColor = backgroundColor
                        ctrl.ForeColor = foregroundColor
                    End If
                Next
            End If

            ' 强制重绘
            Me.Invalidate(True)
        Catch ex As System.Exception
            Debug.WriteLine("ApplyTheme error: " & ex.Message)
        End Try
    End Sub


    Private WithEvents lvMails As ListView
    Private WithEvents taskList As ListView
    Private WithEvents contactInfoList As ListView
    Private WithEvents mailBrowser As WebBrowser
    Private splitter1, splitter2 As SplitContainer
    Private tabControl As TabControl
    Private btnPanel As Panel

    ' 进度指示器相关控件
    Private progressBar As ProgressBar
    Private progressLabel As Label
    Private cancelButton As Button
    Private progressPanel As Panel
    Private cancellationTokenSource As Threading.CancellationTokenSource
    Private currentConversationId As String = String.Empty
    Private currentMailEntryID As String = String.Empty
    Private currentSortColumn As Integer = 0
    Private currentSortOrder As SortOrder = SortOrder.Ascending
    Private currentHighlightEntryID As String
    
    ' EntryID比较缓存，提升高亮匹配性能
    Private entryIdCompareCache As New Dictionary(Of String, String)  ' key: itemEntryID, value: normalized form
    Private entryIdCacheExpireTime As DateTime = DateTime.MinValue
    Private Const CacheExpireMinutes As Integer = 5  ' 缓存5分钟后过期

    Private mailItems As New List(Of (Index As Integer, EntryID As String))  ' 移到这里

    ' 虚拟化ListView相关变量
    Private allMailItems As New List(Of (Index As Integer, EntryID As String))  ' 所有邮件项的完整列表
    Private allListViewItems As New List(Of ListViewItem)  ' 所有ListView项的完整列表
    Private currentPage As Integer = 0  ' 当前页码
    Private totalPages As Integer = 0  ' 总页数
    Private isVirtualMode As Boolean = False  ' 是否启用虚拟模式
    Private isLoadingPage As Boolean = False  ' 是否正在加载页面


    ' 批量属性获取结构
    Private Structure MailItemProperties
        Public EntryID As String
        Public ReceivedTime As DateTime
        Public SenderName As String
        Public Subject As String
        Public MessageClass As String
        Public CreationTime As DateTime
        Public IsValid As Boolean
    End Structure

    ' 在类级别添加一个字典来存储链接和EntryID的映射

    ' 智能缓存机制 - 扩展缓存系统
    Private Shared contactMailCache As New Dictionary(Of String, (Data As String, CacheTime As DateTime))
    Private Shared meetingStatsCache As New Dictionary(Of String, MeetingStatsData)
    Private Shared conversationMailsCache As New Dictionary(Of String, (MailItems As List(Of (Index As Integer, EntryID As String)), ListViewItems As List(Of ListViewItem), CacheTime As DateTime))
    Private Shared contactInfoCache As New Dictionary(Of String, (BusinessPhone As String, MobilePhone As String, Department As String, Company As String, CacheTime As DateTime))
    ' 邮件属性缓存 - 避免重复COM调用
    Private Shared mailPropertiesCache As New Dictionary(Of String, (Properties As MailItemProperties, CacheTime As DateTime))

    Private Const CacheExpiryMinutes As Integer = 30
    Private Const ConversationCacheExpiryMinutes As Integer = 10 ' 会话缓存较短，因为邮件可能频繁更新
    Private Const MeetingStatsCacheExpiryMinutes As Integer = 60 ' 会议统计缓存1小时
    Private Const ContactInfoCacheExpiryMinutes As Integer = 120 ' 联系人信息缓存2小时
    Private Const MailPropertiesCacheExpiryMinutes As Integer = 15 ' 邮件属性缓存15分钟

    ' 虚拟化ListView相关常量
    Private Const PageSize As Integer = 15  ' 每页显示的邮件数量
    Private Const PreloadPages As Integer = 1  ' 预加载的页数

    ' 会议统计数据结构
    Public Structure MeetingStatsData
        Public TotalMeetings As Integer
        Public ProjectStats As Dictionary(Of String, Integer)
        Public UpcomingMeetings As List(Of (MeetingDate As DateTime, Title As String))
        Public CacheTime As DateTime
    End Structure

    ' 清理过期缓存的方法 - 支持多种缓存类型
    Private Shared Sub CleanExpiredCache()
        Try
            ' 清理联系人邮件缓存
            Dim expiredKeys As New List(Of String)
            For Each kvp In contactMailCache
                If DateTime.Now.Subtract(kvp.Value.CacheTime).TotalMinutes >= CacheExpiryMinutes Then
                    expiredKeys.Add(kvp.Key)
                End If
            Next
            For Each key In expiredKeys
                contactMailCache.Remove(key)
            Next

            ' 清理会议统计缓存
            expiredKeys.Clear()
            For Each kvp In meetingStatsCache
                If DateTime.Now.Subtract(kvp.Value.CacheTime).TotalMinutes >= MeetingStatsCacheExpiryMinutes Then
                    expiredKeys.Add(kvp.Key)
                End If
            Next
            For Each key In expiredKeys
                meetingStatsCache.Remove(key)
            Next

            ' 清理会话邮件缓存
            expiredKeys.Clear()
            For Each kvp In conversationMailsCache
                If DateTime.Now.Subtract(kvp.Value.CacheTime).TotalMinutes >= ConversationCacheExpiryMinutes Then
                    expiredKeys.Add(kvp.Key)
                End If
            Next
            For Each key In expiredKeys
                conversationMailsCache.Remove(key)
            Next

            ' 清理联系人信息缓存
            expiredKeys.Clear()
            For Each kvp In contactInfoCache
                If DateTime.Now.Subtract(kvp.Value.CacheTime).TotalMinutes >= ContactInfoCacheExpiryMinutes Then
                    expiredKeys.Add(kvp.Key)
                End If
            Next
            For Each key In expiredKeys
                contactInfoCache.Remove(key)
            Next

            Debug.WriteLine($"缓存清理完成: 联系人邮件{contactMailCache.Count}项, 会议统计{meetingStatsCache.Count}项, 会话邮件{conversationMailsCache.Count}项, 联系人信息{contactInfoCache.Count}项")
        Catch ex As System.Exception
            Debug.WriteLine($"清理缓存时出错: {ex.Message}")
        End Try
    End Sub

    ' 获取缓存的联系人信息
    Private Shared Function GetCachedContactInfo(senderEmail As String) As (BusinessPhone As String, MobilePhone As String, Department As String, Company As String, Found As Boolean)
        If contactInfoCache.ContainsKey(senderEmail) Then
            Dim cached = contactInfoCache(senderEmail)
            If DateTime.Now.Subtract(cached.CacheTime).TotalMinutes < ContactInfoCacheExpiryMinutes Then
                Return (cached.BusinessPhone, cached.MobilePhone, cached.Department, cached.Company, True)
            End If
        End If
        Return ("", "", "", "", False)
    End Function

    ' 缓存联系人信息
    Private Shared Sub CacheContactInfo(senderEmail As String, businessPhone As String, mobilePhone As String, department As String, company As String)
        contactInfoCache(senderEmail) = (businessPhone, mobilePhone, department, company, DateTime.Now)
    End Sub

    ' 删除原来的 mailIndexMap

    Private Sub SetupControls()
        InitializeSplitContainers()
        SetupProgressIndicator()
        SetupMailList()

        ' 延迟加载标签页 - 使用Task.Delay替代Thread.Sleep
        Task.Run(Async Function()
                     ' 使用Task.Delay代替Thread.Sleep，不会阻塞线程
                     Await Task.Delay(100)
                     ' 检查控件是否已经初始化完成
                     If Me.IsHandleCreated Then
                         Me.BeginInvoke(Sub()
                                            SetupTabPages()
                                            BindEvents()
                                        End Sub)
                     Else
                         ' 如果控件尚未完成初始化，等待控件句柄创建完成
                         AddHandler Me.HandleCreated, Sub(s, e)
                                                          Task.Run(Async Function()
                                                                       Await Task.Delay(50)
                                                                       Me.BeginInvoke(Sub()
                                                                                          SetupTabPages()
                                                                                          BindEvents()
                                                                                      End Sub)
                                                                   End Function)
                                                      End Sub
                     End If
                 End Function)
    End Sub

    Private Sub InitializeSplitContainers()
        ' 创建第一个分隔控件
        splitter1 = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .Panel1MinSize = 100,
            .Panel2MinSize = 150,
            .SplitterWidth = 5
        }

        ' 创建第二个分隔控件
        splitter2 = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .Panel1MinSize = 100,
            .Panel2MinSize = 50,
            .SplitterWidth = 5
        }

        ' 先添加第二个分隔控件到第一个分隔控件的Panel2
        splitter1.Panel2.Controls.Add(splitter2)

        ' 在第二个分隔控件的上半部分添加用于显示HTML详情的WebBrowser
        mailBrowser = New WebBrowser With {
            .Dock = DockStyle.Fill,
            .AllowWebBrowserDrop = False,
            .IsWebBrowserContextMenuEnabled = False,
            .ScriptErrorsSuppressed = True
        }
        ' 允许JS调用到VB方法（用于点击链接时可能需要）
        mailBrowser.ObjectForScripting = Me
        splitter2.Panel1.Controls.Add(mailBrowser)

        ' 然后添加第一个分隔控件到窗体
        Me.Controls.Add(splitter1)

        ' 添加尺寸改变事件处理
        AddHandler Me.SizeChanged, AddressOf Control_Resize
        AddHandler splitter1.Panel2.SizeChanged, AddressOf Panel2_SizeChanged
    End Sub

    Private Sub SetupProgressIndicator()
        ' 创建进度面板
        progressPanel = New Panel With {
            .Size = New Size(300, 80),
            .BackColor = Color.LightBlue,
            .BorderStyle = BorderStyle.FixedSingle,
            .Visible = False
        }

        ' 创建进度条
        progressBar = New ProgressBar With {
            .Location = New Point(10, 30),
            .Size = New Size(200, 20),
            .Style = ProgressBarStyle.Continuous
        }

        ' 创建进度标签
        progressLabel = New Label With {
            .Location = New Point(10, 10),
            .Size = New Size(280, 15),
            .Text = "正在处理...",
            .Font = New Font("Microsoft YaHei", 8)
        }

        ' 创建取消按钮
        cancelButton = New Button With {
            .Location = New Point(220, 28),
            .Size = New Size(60, 24),
            .Text = "取消",
            .Font = New Font("Microsoft YaHei", 8)
        }

        ' 添加取消按钮事件
        AddHandler cancelButton.Click, AddressOf CancelButton_Click

        ' 将控件添加到进度面板
        progressPanel.Controls.Add(progressBar)
        progressPanel.Controls.Add(progressLabel)
        progressPanel.Controls.Add(cancelButton)

        ' 将进度面板添加到主控件
        Me.Controls.Add(progressPanel)
        progressPanel.BringToFront()

        ' 居中显示进度面板
        CenterProgressPanel()
    End Sub

    Private Sub CenterProgressPanel()
        If progressPanel IsNot Nothing AndAlso Me.Width > 0 AndAlso Me.Height > 0 Then
            progressPanel.Location = New Point(
                (Me.Width - progressPanel.Width) \ 2,
                (Me.Height - progressPanel.Height) \ 2
            )
        End If
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs)
        If cancellationTokenSource IsNot Nothing Then
            cancellationTokenSource.Cancel()
            HideProgress()
        End If
    End Sub

    ' 显示进度指示器
    Public Sub ShowProgress(message As String, Optional isIndeterminate As Boolean = True)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() ShowProgress(message, isIndeterminate))
            Return
        End If

        Try
            If progressPanel IsNot Nothing Then
                progressLabel.Text = message

                If isIndeterminate Then
                    progressBar.Style = ProgressBarStyle.Marquee
                    progressBar.MarqueeAnimationSpeed = 30
                Else
                    progressBar.Style = ProgressBarStyle.Continuous
                    progressBar.Value = 0
                End If

                CenterProgressPanel()
                progressPanel.Visible = True
                progressPanel.BringToFront()

                ' 创建新的取消令牌
                cancellationTokenSource = New Threading.CancellationTokenSource()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"显示进度指示器时出错: {ex.Message}")
        End Try
    End Sub

    ' 更新进度
    Public Sub UpdateProgress(value As Integer, Optional message As String = Nothing)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() UpdateProgress(value, message))
            Return
        End If

        Try
            If progressBar IsNot Nothing Then
                progressBar.Style = ProgressBarStyle.Continuous
                progressBar.Value = Math.Max(0, Math.Min(100, value))
            End If

            If Not String.IsNullOrEmpty(message) AndAlso progressLabel IsNot Nothing Then
                progressLabel.Text = message
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"更新进度时出错: {ex.Message}")
        End Try
    End Sub

    ' 隐藏进度指示器
    Public Sub HideProgress()
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() HideProgress())
            Return
        End If

        Try
            If progressPanel IsNot Nothing Then
                progressPanel.Visible = False
            End If

            If cancellationTokenSource IsNot Nothing Then
                cancellationTokenSource.Dispose()
                cancellationTokenSource = Nothing
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"隐藏进度指示器时出错: {ex.Message}")
        End Try
    End Sub

    ' 获取取消令牌
    Public ReadOnly Property CancellationToken As Threading.CancellationToken
        Get
            Return If(cancellationTokenSource?.Token, Threading.CancellationToken.None)
        End Get
    End Property

    ' 添加用于 JavaScript 调用的方法
    <ComVisible(True)>
    Public Sub OpenBrowserLink(url As String)
        Try
            Process.Start(New ProcessStartInfo With {
                .FileName = url,
                .UseShellExecute = True
            })
        Catch ex As System.Exception
            Debug.WriteLine($"打开链接出错: {ex.Message}")
            MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。")
        End Try
    End Sub


    Private Sub Control_Resize(sender As Object, e As EventArgs)
        Try
            If Not Me.IsHandleCreated OrElse Me.Height <= 0 Then
                Return
            End If

            ' 计算并设置第一个分隔条位置
            Dim targetHeight1 As Integer = CInt(Me.Height * 0.2)
            Dim maxDistance1 As Integer = Me.Height - splitter1.Panel2MinSize
            Dim minDistance1 As Integer = splitter1.Panel1MinSize

            If 0 Then
                ' 添加调试信息
                Debug.WriteLine($"Splitter1 尺寸信息:")
                Debug.WriteLine($"  控件总高度: {Me.Height}")
                Debug.WriteLine($"  目标位置: {targetHeight1}")
                Debug.WriteLine($"  最小位置: {minDistance1}")
                Debug.WriteLine($"  最大位置: {maxDistance1}")
                Debug.WriteLine($"  Panel1MinSize: {splitter1.Panel1MinSize}")
                Debug.WriteLine($"  Panel2MinSize: {splitter1.Panel2MinSize}")
                Debug.WriteLine($"  当前SplitterDistance: {splitter1.SplitterDistance}")
            End If

            splitter1.SplitterDistance = Math.Max(minDistance1, Math.Min(targetHeight1, maxDistance1))

        Catch ex As System.Exception
            Debug.WriteLine($"Control_Resize error: {ex.Message}")
        End Try
    End Sub

    Private Sub Panel2_SizeChanged(sender As Object, e As EventArgs)
        Try
            If Not splitter2.IsHandleCreated OrElse splitter2.Height <= (splitter2.Panel1MinSize + splitter2.Panel2MinSize) Then
                Return
            End If

            ' 计算并设置第二个分隔条位置
            Dim panel2Height As Integer = splitter2.Height
            ' 确保目标高度不小于Panel1MinSize
            Dim targetHeight2 As Integer = Math.Max(
                splitter2.Panel1MinSize,
                CInt(panel2Height * 0.75)
            )
            ' 确保最大距离考虑了两个面板的最小尺寸
            Dim maxDistance2 As Integer = panel2Height - splitter2.Panel2MinSize
            Dim minDistance2 As Integer = splitter2.Panel1MinSize

            If 0 Then
                ' 添加调试信息
                Debug.WriteLine($"Splitter2 尺寸信息 (修正后):")
                Debug.WriteLine($"  Panel2总高度: {panel2Height}")
                Debug.WriteLine($"  目标位置: {targetHeight2}")
                Debug.WriteLine($"  最小位置: {minDistance2}")
                Debug.WriteLine($"  最大位置: {maxDistance2}")
                Debug.WriteLine($"  Panel1MinSize: {splitter2.Panel1MinSize}")
                Debug.WriteLine($"  Panel2MinSize: {splitter2.Panel2MinSize}")
                Debug.WriteLine($"  当前SplitterDistance: {splitter2.SplitterDistance}")
            End If

            splitter2.SplitterDistance = Math.Max(minDistance2, Math.Min(targetHeight2, maxDistance2))

        Catch ex As System.Exception
            Debug.WriteLine($"Panel2_SizeChanged error: {ex.Message}")
        End Try
    End Sub
    Private Sub Form_Load(sender As Object, e As EventArgs)
        Try
            ' 使用完整命名空间避免歧义
            System.Windows.Forms.Application.DoEvents()

            ' 设置默认的分隔比例而不是固定像素值
            splitter1.SplitterDistance = CInt(Me.Height * 0.2)
            splitter2.SplitterDistance = CInt(splitter1.Panel2.Height * 0.85)

            ' 添加分隔条移动后的事件处理
            AddHandler splitter1.SplitterMoved, AddressOf Splitter_Moved
            AddHandler splitter2.SplitterMoved, AddressOf Splitter_Moved
        Catch ex As System.Exception
            Debug.WriteLine($"设置分隔位置出错: {ex.Message}")
        End Try
    End Sub

    ' 添加 Splitter_Moved 方法定义
    Private Sub Splitter_Moved(sender As Object, e As SplitterEventArgs)
        Try
            Dim splitter As SplitContainer = DirectCast(sender, SplitContainer)
            ' 确保分隔条位置在有效范围内
            If splitter.SplitterDistance < splitter.Panel1MinSize Then
                splitter.SplitterDistance = splitter.Panel1MinSize
            ElseIf splitter.SplitterDistance > (splitter.Height - splitter.Panel2MinSize) Then
                splitter.SplitterDistance = splitter.Height - splitter.Panel2MinSize
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"调整分隔条位置出错: {ex.Message}")
        End Try
    End Sub

    ' 快速获取图标索引的函数 - 基于MAPI行数据，使用缓存优化
    Private Shared Function GetIconIndex(messageClass As String, hasAttach As Boolean, flagStatus As Integer) As Integer
        ' 使用缓存字典快速获取基础索引
        Dim baseIndex As Integer = 0
        If Not String.IsNullOrEmpty(messageClass) Then
            ' 首先尝试精确匹配
            If MessageClassBaseIndex.TryGetValue(messageClass, baseIndex) Then
                ' 找到精确匹配
            ElseIf messageClass.StartsWith("IPM.Appointment") OrElse messageClass.StartsWith("IPM.Schedule.Meeting") Then
                baseIndex = 6  ' 日历/会议基础索引
            ElseIf messageClass.StartsWith("IPM.Task") Then
                baseIndex = 12 ' 任务基础索引
            ElseIf messageClass.StartsWith("IPM.Contact") Then
                baseIndex = 18 ' 联系人基础索引
            Else
                baseIndex = 0  ' 邮件基础索引（默认）
            End If
        End If
        
        ' 计算附件偏移（+1如果有附件）
        Dim attachOffset As Integer = If(hasAttach, 1, 0)
        
        ' 计算旗标偏移（+2进行中，+4已完成）
        Dim flagOffset As Integer = 0
        Select Case flagStatus
            Case 2 ' olFlagMarked (进行中)
                flagOffset = 2
            Case 1 ' olFlagComplete (已完成)
                flagOffset = 4
            Case Else ' 无旗标或其他状态
                flagOffset = 0
        End Select
        
        Return baseIndex + attachOffset + flagOffset
    End Function

    ' 快速获取图标文本的函数
    Private Shared Function GetIconTextFast(messageClass As String, hasAttach As Boolean, flagStatus As Integer) As String
        Dim index As Integer = GetIconIndex(messageClass, hasAttach, flagStatus)
        If index >= 0 AndAlso index < IconCombinations.Length Then
            Return IconCombinations(index)
        Else
            Return "📧" ' 默认邮件图标
        End If
    End Function

    Private Function GetItemImageText(item As Object) As String
        Try
            Dim icons As New List(Of String)
            Debug.WriteLine($"GetItemImageText: 处理项目类型 {item.GetType().Name}")

            ' 检查项目类型
            If TypeOf item Is Outlook.MailItem Then
                icons.Add("✉️") '📧
                
                ' 检查附件
                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                Try
                    If mail.Attachments IsNot Nothing AndAlso mail.Attachments.Count > 0 Then
                        icons.Add("📎") ' 回形针图标表示有附件
                    End If
                Catch ex As System.Exception
                    ' 忽略附件检查错误
                End Try
                
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                icons.Add("📅")
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                icons.Add("📅") ' 会议邮件也使用日历图标，保持一致性
            Else
                icons.Add("❓")
            End If

            ' 根据任务状态添加不同的图标
            Select Case CheckItemHasTask(item)
                Case TaskStatus.InProgress
                    icons.Add("🚩") ' 红色旗标 - 未完成的任务
                Case TaskStatus.Completed
                    icons.Add("⚑")   ' 黑色旗标 - 已完成的任务
            End Select

            Return String.Join(" ", icons)
        Catch ex As System.Exception
            Debug.WriteLine($"获取图标文本出错: {ex.Message}")
            Return "❓"
        End Try
    End Function

    Private Sub SetupMailList()
        lvMails = New ListView With {
            .Dock = DockStyle.Fill,
            .View = Windows.Forms.View.Details,
            .FullRowSelect = True,
            .HideSelection = False,  ' 确保失去焦点时仍显示选中项
            .Sorting = SortOrder.Descending,
            .AllowColumnReorder = True,
            .HeaderStyle = ColumnHeaderStyle.Clickable,
            .OwnerDraw = True,  ' 启用自定义绘制
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor,
            .SmallImageList = New ImageList() With {.ImageSize = New Size(16, 15)}, ' 设置行高
            .VirtualMode = False  ' 初始禁用虚拟模式，根据需要动态启用
        }
        
        ' 启用双缓冲以减少闪烁
        Dim listViewType As Type = lvMails.GetType()
        Dim doubleBufferedProperty As Reflection.PropertyInfo = listViewType.GetProperty("DoubleBuffered", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
        If doubleBufferedProperty IsNot Nothing Then
            doubleBufferedProperty.SetValue(lvMails, True, Nothing)
        End If

        lvMails.Columns.Add("----", 50)  ' 增加宽度以适应更大的图标
        lvMails.Columns.Add("日期", 120) ' 宽度适配“yyyy-MM-dd HH:mm”
        With lvMails.Columns.Add("发件人", 100)
            .TextAlign = HorizontalAlignment.Left
        End With
        With lvMails.Columns.Add("主题", 300)
            .TextAlign = HorizontalAlignment.Left
        End With

        ' 设置文本省略模式
        'For Each column As ColumnHeader In lvMails.Columns
        '    column.Width = -2  ' 自动调整列宽以适应内容
        'Next

        ' 创建分页导航面板
        Dim paginationPanel As New Panel With {
            .Height = 25,
            .Dock = DockStyle.Bottom,
            .BackColor = currentBackColor,
            .Padding = New Padding(0, 0, 0, 0)
        }

        ' 创建分页导航控件
        Dim btnFirstPage As New Button With {
            .Text = "首页",
            .Size = New Size(50, 25),
            .Location = New Point(5, 5)
        }

        Dim btnPrevPage As New Button With {
            .Text = "上页",
            .Size = New Size(50, 25),
            .Location = New Point(60, 5)
        }

        Dim lblPageInfo As New Label With {
            .Text = "第1页/共1页",
            .Size = New Size(100, 25),
            .Location = New Point(115, 8),
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.Transparent
        }

        Dim btnNextPage As New Button With {
            .Text = "下页",
            .Size = New Size(50, 25),
            .Location = New Point(220, 5)
        }

        Dim btnLastPage As New Button With {
            .Text = "末页",
            .Size = New Size(50, 25),
            .Location = New Point(275, 5)
        }

        Dim lblItemCount As New Label With {
            .Text = "共0项",
            .Size = New Size(80, 20),
            .Location = New Point(330, 3),
            .TextAlign = ContentAlignment.MiddleLeft,
            .BackColor = Color.Transparent
        }

        ' 添加分页开关控件
        Dim chkPagination As New CheckBox With {
            .Text = "分页",
            .Size = New Size(60, 25),
            .Location = New Point(420, 5),
            .Checked = _isPaginationEnabled,
            .BackColor = Color.Transparent
        }

        ' 添加分页开关事件处理
        AddHandler chkPagination.CheckedChanged, Sub(sender, e)
                                                     IsPaginationEnabled = chkPagination.Checked
                                                 End Sub

        ' 存储分页控件引用
        paginationPanel.Tag = New With {
            .FirstPage = btnFirstPage,
            .PrevPage = btnPrevPage,
            .PageInfo = lblPageInfo,
            .NextPage = btnNextPage,
            .LastPage = btnLastPage,
            .ItemCount = lblItemCount,
            .PaginationCheckBox = chkPagination
        }

        ' 添加事件处理
        If _isPaginationEnabled Then
            AddHandler btnFirstPage.Click, Async Sub() Await LoadPageAsync(0)
            AddHandler btnPrevPage.Click, Async Sub() Await LoadPreviousPageAsync()
            AddHandler btnNextPage.Click, Async Sub() Await LoadNextPageAsync()
            AddHandler btnLastPage.Click, Async Sub() Await LoadPageAsync(totalPages - 1)
        End If

        ' 添加控件到面板
        paginationPanel.Controls.AddRange({btnFirstPage, btnPrevPage, lblPageInfo, btnNextPage, btnLastPage, lblItemCount, chkPagination})

        ' 添加到主面板
        splitter1.Panel1.Controls.Add(paginationPanel)
        splitter1.Panel1.Controls.Add(lvMails)

        ' 存储分页面板引用
        splitter1.Panel1.Tag = paginationPanel

        ' 添加绘制事件处理
        AddHandler lvMails.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler lvMails.DrawSubItem, AddressOf ListView_DrawSubItem
        
        ' 添加虚拟模式事件处理
        AddHandler lvMails.RetrieveVirtualItem, AddressOf ListView_RetrieveVirtualItem
        AddHandler lvMails.CacheVirtualItems, AddressOf ListView_CacheVirtualItems
    End Sub



    Private Sub ListView_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        e.DrawDefault = True
    End Sub

    Private Sub ListView_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        ' 使用当前项的背景色
        Dim backBrush As Brush = New SolidBrush(e.Item.BackColor)
        e.Graphics.FillRectangle(backBrush, e.Bounds)

        ' 第一列使用 emoji 字体，其他列使用默认字体
        Dim sf As New StringFormat()
        sf.Trimming = StringTrimming.EllipsisCharacter
        sf.FormatFlags = StringFormatFlags.NoWrap

        If e.ColumnIndex = 0 Then

            If e.SubItem.Text.Contains("🚩") Then
                ' 使用特殊颜色和字体
                Dim specialFont As New Font(iconFont, FontStyle.Bold)
                Dim specialBrush As Brush = Brushes.Red
                e.Graphics.DrawString(e.SubItem.Text, specialFont, specialBrush, e.Bounds, sf)
            Else
                e.Graphics.DrawString(e.SubItem.Text, iconFont, Brushes.Black, e.Bounds, sf)
            End If
        Else
            ' 根据是否高亮使用不同字体
            Dim font As Font = If(e.Item.BackColor = highlightColor, highlightFont, normalFont)
            e.Graphics.DrawString(e.SubItem.Text, font, Brushes.Black, e.Bounds, sf)
        End If
        backBrush.Dispose()
    End Sub

    ' ListView虚拟模式事件处理器
    Private Sub ListView_RetrieveVirtualItem(sender As Object, e As RetrieveVirtualItemEventArgs)
        Try
            If e.ItemIndex >= 0 AndAlso e.ItemIndex < allListViewItems.Count Then
                ' 创建虚拟项的副本
                Dim originalItem = allListViewItems(e.ItemIndex)
                Dim virtualItem As New ListViewItem(originalItem.Text)
                virtualItem.Tag = originalItem.Tag
                virtualItem.Name = originalItem.Name
                virtualItem.BackColor = originalItem.BackColor
                virtualItem.ForeColor = originalItem.ForeColor
                
                ' 复制所有子项
                For si As Integer = 1 To originalItem.SubItems.Count - 1
                    virtualItem.SubItems.Add(originalItem.SubItems(si).Text)
                Next
                
                e.Item = virtualItem
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"RetrieveVirtualItem error: {ex.Message}")
            ' 创建错误项
            e.Item = New ListViewItem("❌ 加载失败")
        End Try
    End Sub

    Private Sub ListView_CacheVirtualItems(sender As Object, e As CacheVirtualItemsEventArgs)
        ' 可选：预缓存指定范围的项目以提高性能
        Debug.WriteLine($"缓存虚拟项: {e.StartIndex} 到 {e.EndIndex}")
    End Sub


    Private Sub SetupTabPages()
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill
        }
        splitter2.Panel2.Controls.Add(tabControl)

        ' 只初始化第一个标签页
        SetupActionsTab()

        ' 延迟加载其他标签页（优化：使用BeginInvoke避免阻塞UI）
        'Task.Run(Sub()
        '            Me.BeginInvoke(Sub()
        '                              SetupTasksTab()
        '                             SetupNoteTab()
        '                            tabControl.SelectedIndex = 0
        '                       End Sub)
        '   End Sub)
    End Sub


    ' 添加检查方法
    Private Function CheckComVisibleAttribute() As Boolean
        Try
            Dim type As Type = Me.GetType()
            Dim attr As ComVisibleAttribute = DirectCast(
                Attribute.GetCustomAttribute(type, GetType(ComVisibleAttribute)),
                ComVisibleAttribute)
            Return attr IsNot Nothing AndAlso attr.Value
        Catch ex As System.Exception
            Debug.WriteLine($"检查 ComVisible 特性时出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub SetupNotesTab()
        Dim tabPage1 As New TabPage("笔记")

        ' 创建容器面板
        Dim containerPanel As New Panel With {
            .Dock = DockStyle.Fill
        }

        ' 创建按钮面板
        Dim buttonPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 40
        }

        ' 添加新建笔记按钮
        Dim btnNewNote As New Button With {
            .Text = "新建笔记",
            .Location = New Point(10, 5),
            .Size = New Size(80, 30)
        }
        AddHandler btnNewNote.Click, AddressOf btnNewNote_Click
        buttonPanel.Controls.Add(btnNewNote)  ' 确保按钮被添加到面板中

        ' 创建笔记列表视图
        Dim noteListView As New ListView With {
            .Dock = DockStyle.Fill,
            .View = Windows.Forms.View.Details,  ' Specify the namespace explicitly
            .FullRowSelect = True,
            .GridLines = True,
            .MultiSelect = False
        }

        ' 添加列
        noteListView.Columns.Add("创建日期", 120)
        noteListView.Columns.Add("标题", 200)
        noteListView.Columns.Add("操作", 100)

        ' 添加双击事件处理
        AddHandler noteListView.DoubleClick, Sub(sender, e)
                                                 If noteListView.SelectedItems.Count > 0 Then
                                                     Dim link As String = noteListView.SelectedItems(0).Tag?.ToString()
                                                     If Not String.IsNullOrEmpty(link) Then
                                                         Process.Start(New ProcessStartInfo With {
                                                        .FileName = link,
                                                        .UseShellExecute = True
                                                    })
                                                     End If
                                                 End If
                                             End Sub



        ' 按正确的顺序添加控件
        containerPanel.Controls.Add(noteListView)
        containerPanel.Controls.Add(buttonPanel)
        tabPage1.Controls.Add(containerPanel)
        tabControl.TabPages.Add(tabPage1)

        ' 保存对 ListView 的引用以便后续更新
        noteListView.Tag = "NoteList"  ' 添加标识
    End Sub

    ' 修改 GenerateHtmlContent 方法为 UpdateNoteList 方法
    Private Sub UpdateNoteList(noteList As List(Of (CreateTime As String, Title As String, Link As String)))
        ' 确保在 UI 线程上执行
        If Me.InvokeRequired Then
            Me.Invoke(Sub() UpdateNoteList(noteList))
            Return
        End If

        ' 查找笔记列表视图
        Dim noteListView As ListView = Nothing
        For Each tabPage As TabPage In tabControl.TabPages
            If tabPage.Text = "笔记" Then
                For Each control As Control In tabPage.Controls
                    If TypeOf control Is Panel Then
                        For Each subControl As Control In control.Controls
                            If TypeOf subControl Is ListView AndAlso subControl.Tag?.ToString() = "NoteList" Then
                                noteListView = DirectCast(subControl, ListView)
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If
        Next

        If noteListView Is Nothing Then Return

        noteListView.Items.Clear()

        For Each note In noteList
            Dim item As New ListViewItem(If(note.CreateTime, DateTime.Now.ToString("yyyy-MM-dd HH:mm")))
            item.SubItems.Add(If(note.Title, "无标题"))
            item.SubItems.Add("打开笔记")
            item.Tag = note.Link
            noteListView.Items.Add(item)
        Next
    End Sub

    Private Sub GetAllMailFolders(folder As Outlook.Folder, folderList As List(Of Outlook.Folder))
        Try
            ' 定义要搜索的核心文件夹名称
            Dim coreFolders As New List(Of String) From {
            "收件箱",
            "Inbox",
            "已发送邮件",
            "Sent Items",
            "Todo",
            "Doc",
            "Processed Mail",
            "Archive",
            "Doc",
            "归档",
            "todo",
            "processed mail"
        }

            ' 检查当前文件夹是否是邮件文件夹且在核心文件夹列表中
            Dim isMailItem As Boolean = False
            Me.Invoke(Sub()
                          isMailItem = (folder.DefaultItemType = Outlook.OlItemType.olMailItem)
                      End Sub)

            If isMailItem AndAlso coreFolders.Contains(folder.Name) Then
                folderList.Add(folder)
            End If

            ' 只在核心文件夹中递归搜索
            Dim subFolders As Outlook.Folders = Nothing
            Me.Invoke(Sub()
                          subFolders = folder.Folders
                      End Sub)

            If subFolders IsNot Nothing Then
                For Each subFolder As Outlook.Folder In subFolders
                    If coreFolders.Contains(subFolder.Name) Then
                        GetAllMailFolders(subFolder, folderList)
                    End If
                Next
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"处理文件夹 {folder.Name} 时出错: {ex.Message}")
        End Try
    End Sub
    ' 添加一个新的辅助方法用于递归获取所有邮件文件夹
    Private Sub GetAllMailFoldersAll(folder As Outlook.Folder, folderList As List(Of Outlook.Folder))
        Try
            Me.Invoke(Sub()
                          ' 添加当前文件夹（如果是邮件文件夹）
                          If folder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                              folderList.Add(folder)
                          End If

                          ' 递归处理子文件夹
                          For Each subFolder As Outlook.Folder In folder.Folders
                              GetAllMailFolders(subFolder, folderList)
                          Next
                      End Sub)
        Catch ex As System.Exception
            Debug.WriteLine($"处理文件夹 {folder.Name} 时出错: {ex.Message}")
        End Try
    End Sub

    ' 异步获取联系人信息的方法
    Private Async Function GetContactInfoAsync() As Task(Of String)
        Try
            ShowProgress("正在获取联系人信息...")
            Return Await Task.Run(Function()
                                      CancellationToken.ThrowIfCancellationRequested()
                                      Return GetContactInfoBackground()
                                  End Function)
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("联系人信息获取被取消")
            Return "操作已取消"
        Finally
            HideProgress()
        End Try
    End Function

    ' 在后台线程执行的联系人信息获取方法
    Private Function GetContactInfoBackground() As String
        Try
            Dim info As New StringBuilder()
            ' 性能监控
            Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim elapsedContactSearch As Long
            Dim elapsedMeetingStats As Long
            Dim elapsedMailStats As Long

            ' 在后台线程中直接访问COM对象
            Dim currentItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem Is Nothing Then Return "未选择邮件项"

            Dim senderEmail As String = String.Empty
            Dim senderName As String = String.Empty

            ' 获取发件人信息
            If TypeOf currentItem Is Outlook.MailItem Then
                Dim mail = DirectCast(currentItem, Outlook.MailItem)
                Try
                    senderEmail = mail.SenderEmailAddress
                    senderName = mail.SenderName
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常获取邮件发件人信息 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return "获取邮件发件人信息时发生COM异常"
                Catch ex As System.Exception
                    Debug.WriteLine($"获取邮件发件人信息时发生异常: {ex.Message}")
                    Return "获取邮件发件人信息时发生异常"
                End Try
            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                Dim meeting = DirectCast(currentItem, Outlook.MeetingItem)
                Try
                    senderEmail = meeting.SenderEmailAddress
                    senderName = meeting.SenderName
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常获取会议发件人信息 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return "获取会议发件人信息时发生COM异常"
                Catch ex As System.Exception
                    Debug.WriteLine($"获取会议发件人信息时发生异常: {ex.Message}")
                    Return "获取会议发件人信息时发生异常"
                End Try
            End If

            If String.IsNullOrEmpty(senderEmail) Then Return "无法获取发件人信息"

            ' 清理过期缓存
            CleanExpiredCache()

            ' 检查缓存
            If contactMailCache.ContainsKey(senderEmail) Then
                Dim cached = contactMailCache(senderEmail)
                If DateTime.Now.Subtract(cached.CacheTime).TotalMinutes < CacheExpiryMinutes Then
                    Return cached.Data
                End If
            End If

            info.AppendLine($"发件人: {senderName}")
            info.AppendLine($"邮箱: {senderEmail}")
            info.AppendLine("----------------------------------------")

            ' 搜索联系人信息 - 使用智能缓存机制
            Dim swContact = System.Diagnostics.Stopwatch.StartNew()
            Dim cachedContactInfo = GetCachedContactInfo(senderEmail)

            If cachedContactInfo.Found Then
                ' 使用缓存的联系人信息
                info.AppendLine("联系人信息:")
                If Not String.IsNullOrEmpty(cachedContactInfo.BusinessPhone) Then info.AppendLine($"工作电话: {cachedContactInfo.BusinessPhone}")
                If Not String.IsNullOrEmpty(cachedContactInfo.MobilePhone) Then info.AppendLine($"手机: {cachedContactInfo.MobilePhone}")
                If Not String.IsNullOrEmpty(cachedContactInfo.Department) Then info.AppendLine($"部门: {cachedContactInfo.Department}")
                If Not String.IsNullOrEmpty(cachedContactInfo.Company) Then info.AppendLine($"公司: {cachedContactInfo.Company}")
                info.AppendLine("----------------------------------------")
                Debug.WriteLine("使用缓存的联系人信息")
            Else
                ' 从Outlook获取联系人信息并缓存
                Try
                    Dim contacts = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
                    Dim filter = $"[Email1Address] = '{senderEmail}' OR [Email2Address] = '{senderEmail}' OR [Email3Address] = '{senderEmail}'"

                    ' 使用GetTable代替Items.Restrict获取更好性能
                    Dim contactTable = contacts.GetTable(filter)
                    ' 优化：只添加需要的列，减少数据传输
                    contactTable.Columns.RemoveAll() ' 移除默认列
                    contactTable.Columns.Add("BusinessTelephoneNumber")
                    contactTable.Columns.Add("MobileTelephoneNumber")
                    contactTable.Columns.Add("Department")
                    contactTable.Columns.Add("CompanyName")

                    Dim bt As String = ""
                    Dim mt As String = ""
                    Dim dept As String = ""
                    Dim comp As String = ""

                    If Not contactTable.EndOfTable Then
                        Dim crow = contactTable.GetNextRow()
                        bt = crow("BusinessTelephoneNumber")?.ToString()
                        mt = crow("MobileTelephoneNumber")?.ToString()
                        dept = crow("Department")?.ToString()
                        comp = crow("CompanyName")?.ToString()

                        info.AppendLine("联系人信息:")
                        If Not String.IsNullOrEmpty(bt) Then info.AppendLine($"工作电话: {bt}")
                        If Not String.IsNullOrEmpty(mt) Then info.AppendLine($"手机: {mt}")
                        If Not String.IsNullOrEmpty(dept) Then info.AppendLine($"部门: {dept}")
                        If Not String.IsNullOrEmpty(comp) Then info.AppendLine($"公司: {comp}")
                        info.AppendLine("----------------------------------------")
                    End If

                    ' 缓存联系人信息（即使为空也缓存，避免重复查询）
                    CacheContactInfo(senderEmail, bt, mt, dept, comp)

                    ' 释放COM对象
                    Runtime.InteropServices.Marshal.ReleaseComObject(contactTable)
                    Runtime.InteropServices.Marshal.ReleaseComObject(contacts)
                    Debug.WriteLine("从Outlook获取并缓存联系人信息")
                Catch ex As System.Exception
                    Debug.WriteLine($"搜索联系人信息时出错: {ex.Message}")
                    info.AppendLine("联系人信息: 搜索失败")
                    info.AppendLine("----------------------------------------")
                    ' 缓存失败结果，避免重复尝试
                    CacheContactInfo(senderEmail, "", "", "", "")
                End Try
            End If
            swContact.Stop()
            elapsedContactSearch = swContact.ElapsedMilliseconds

            ' 统计会议信息 - 使用智能缓存机制
            Dim swMeeting = System.Diagnostics.Stopwatch.StartNew()
            Dim meetingCacheKey = $"meeting_{senderEmail}"

            ' 检查会议统计缓存
            If meetingStatsCache.ContainsKey(meetingCacheKey) AndAlso
               (DateTime.Now - meetingStatsCache(meetingCacheKey).CacheTime).TotalMinutes < MeetingStatsCacheExpiryMinutes Then
                ' 使用缓存的会议统计
                Dim cachedStats = meetingStatsCache(meetingCacheKey)
                info.AppendLine($"会议统计 (近2个月):")
                info.AppendLine($"总会议数: {cachedStats.TotalMeetings}")
                info.AppendLine("按项目分类:")
                For Each kvp In cachedStats.ProjectStats.OrderByDescending(Function(x) x.Value)
                    info.AppendLine($"- {kvp.Key}: {kvp.Value}次")
                Next

                info.AppendLine(vbCrLf & "即将到来的会议:")
                For Each meeting In cachedStats.UpcomingMeetings.OrderBy(Function(x) x.MeetingDate).Take(3)
                    info.AppendLine($"- {meeting.MeetingDate:MM/dd HH:mm} {meeting.Title}")
                Next
                info.AppendLine("----------------------------------------")
                Debug.WriteLine("使用缓存的会议统计")
            Else
                ' 从Outlook获取会议统计并缓存
                Try
                    Dim calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                    Dim startDate = DateTime.Now.AddMonths(-2)
                    Dim endDate = DateTime.Now.AddMonths(1)

                    ' 构建两个过滤条件：一个获取与该联系人相关的会议（必要与会者），一个获取可选与会者
                    Dim requiredFilter = $"[Start] >= '{startDate:MM/dd/yyyy}' AND [End] <= '{endDate:MM/dd/yyyy}' AND [RequiredAttendees] LIKE '%{senderEmail}%'"
                    Dim optionalFilter = $"[Start] >= '{startDate:MM/dd/yyyy}' AND [End] <= '{endDate:MM/dd/yyyy}' AND [OptionalAttendees] LIKE '%{senderEmail}%'"

                    ' 使用Table优化会议统计
                    Dim meetingStats As New Dictionary(Of String, Integer)
                    Dim totalMeetings As Integer = 0
                    Dim upcomingMeetings As New List(Of (MeetingDate As DateTime, Title As String))

                    ' 处理必要与会者的会议
                    Dim requiredTable = calendar.GetTable(requiredFilter)
                    ' 优化：只添加需要的列，减少数据传输
                    requiredTable.Columns.RemoveAll() ' 移除默认列
                    requiredTable.Columns.Add("Subject")
                    requiredTable.Columns.Add("Start")

                    Do Until requiredTable.EndOfTable
                        Dim row = requiredTable.GetNextRow()
                        totalMeetings += 1

                        ' 获取会议主题和开始时间
                        Dim subject = If(row("Subject")?.ToString(), "")
                        Dim startObj = row("Start")

                        If Not String.IsNullOrEmpty(subject) Then
                            ' 提取项目名称
                            Dim projectName = "其他"
                            Dim match = System.Text.RegularExpressions.Regex.Match(subject, "\[(.*?)\]")
                            If match.Success Then
                                projectName = match.Groups(1).Value
                            End If

                            If meetingStats.ContainsKey(projectName) Then
                                meetingStats(projectName) += 1
                            Else
                                meetingStats.Add(projectName, 1)
                            End If

                            ' 检查是否是即将到来的会议
                            If startObj IsNot Nothing Then
                                Try
                                    Dim startTime As DateTime = DateTime.Parse(startObj.ToString())
                                    If startTime > DateTime.Now Then
                                        upcomingMeetings.Add((startTime, subject))
                                    End If
                                Catch
                                    ' 忽略日期解析错误
                                End Try
                            End If
                        End If
                    Loop

                    ' 处理可选与会者的会议
                    Dim optionalTable = calendar.GetTable(optionalFilter)
                    ' 优化：只添加需要的列，减少数据传输
                    optionalTable.Columns.RemoveAll() ' 移除默认列
                    optionalTable.Columns.Add("Subject")
                    optionalTable.Columns.Add("Start")

                    Do Until optionalTable.EndOfTable
                        Dim row = optionalTable.GetNextRow()
                        totalMeetings += 1

                        ' 获取会议主题和开始时间
                        Dim subject = If(row("Subject")?.ToString(), "")
                        Dim startObj = row("Start")

                        If Not String.IsNullOrEmpty(subject) Then
                            ' 提取项目名称
                            Dim projectName = "其他"
                            Dim match = System.Text.RegularExpressions.Regex.Match(subject, "\[(.*?)\]")
                            If match.Success Then
                                projectName = match.Groups(1).Value
                            End If

                            If meetingStats.ContainsKey(projectName) Then
                                meetingStats(projectName) += 1
                            Else
                                meetingStats.Add(projectName, 1)
                            End If

                            ' 检查是否是即将到来的会议
                            If startObj IsNot Nothing Then
                                Try
                                    Dim startTime As DateTime = DateTime.Parse(startObj.ToString())
                                    If startTime > DateTime.Now Then
                                        upcomingMeetings.Add((startTime, subject))
                                    End If
                                Catch
                                    ' 忽略日期解析错误
                                End Try
                            End If
                        End If
                    Loop

                    ' 缓存会议统计结果
                    meetingStatsCache(meetingCacheKey) = New MeetingStatsData With {
                        .TotalMeetings = totalMeetings,
                        .ProjectStats = meetingStats,
                        .UpcomingMeetings = upcomingMeetings,
                        .CacheTime = DateTime.Now
                    }

                    ' 显示会议统计
                    info.AppendLine($"会议统计 (近2个月):")
                    info.AppendLine($"总会议数: {totalMeetings}")
                    info.AppendLine("按项目分类:")
                    For Each kvp In meetingStats.OrderByDescending(Function(x) x.Value)
                        info.AppendLine($"- {kvp.Key}: {kvp.Value}次")
                    Next

                    info.AppendLine(vbCrLf & "即将到来的会议:")
                    For Each meeting In upcomingMeetings.OrderBy(Function(x) x.MeetingDate).Take(3)
                        info.AppendLine($"- {meeting.MeetingDate:MM/dd HH:mm} {meeting.Title}")
                    Next
                    info.AppendLine("----------------------------------------")

                    ' 释放COM对象
                    Runtime.InteropServices.Marshal.ReleaseComObject(requiredTable)
                    Runtime.InteropServices.Marshal.ReleaseComObject(optionalTable)
                    Runtime.InteropServices.Marshal.ReleaseComObject(calendar)
                    Debug.WriteLine("从Outlook获取并缓存会议统计")
                Catch ex As System.Exception
                    Debug.WriteLine($"统计会议信息时出错: {ex.Message}")
                    info.AppendLine("会议统计: 获取失败")
                    info.AppendLine("----------------------------------------")
                End Try
            End If

            swMeeting.Stop()
            elapsedMeetingStats = swMeeting.ElapsedMilliseconds

            ' 统计邮件往来 - 优化版本
            Dim swMail = System.Diagnostics.Stopwatch.StartNew()
            Dim mailCount As Integer = 0
            Dim recentMails As New List(Of (Received As DateTime, Subject As String))

            ' 获取优先搜索的文件夹
            Dim folders As New List(Of Outlook.Folder)
            Dim store As Outlook.Store = Globals.ThisAddIn.Application.Session.DefaultStore

            ' 获取收件箱及其指定子文件夹
            Dim inbox As Outlook.Folder = TryCast(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), Outlook.Folder)
            If inbox IsNot Nothing Then
                folders.Add(inbox)
                For Each subFolder As Outlook.Folder In inbox.Folders
                    If subFolder.Name.Equals("Doc", StringComparison.OrdinalIgnoreCase) OrElse
                       subFolder.Name.Equals("Processed Mail", StringComparison.OrdinalIgnoreCase) OrElse
                       subFolder.Name.Equals("Todo", StringComparison.OrdinalIgnoreCase) Then
                        folders.Add(subFolder)
                    End If
                Next
            End If

            ' 获取已发送邮件文件夹
            Dim sentItems As Outlook.Folder = TryCast(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail), Outlook.Folder)
            If sentItems IsNot Nothing Then
                folders.Add(sentItems)
            End If

            ' 获取归档文件夹 (假设其名称为 "Archive" 或 "归档") - 在后台线程中直接访问COM对象
            Try
                Dim rootFolders = store.GetRootFolder().Folders
                For i As Integer = 1 To rootFolders.Count
                    Dim rootFolder = rootFolders.Item(i)
                    Dim folderName = rootFolder.Name
                    If folderName.Equals("Archive", StringComparison.OrdinalIgnoreCase) OrElse
                       folderName.Equals("归档", StringComparison.OrdinalIgnoreCase) Then
                        folders.Add(rootFolder)
                        Exit For
                    End If
                Next
            Catch ex As System.Exception
                Debug.WriteLine($"获取归档文件夹时出错: {ex.Message}")
            End Try

            ' 添加时间范围限制，只搜索最近3个月的邮件
            Dim dateFilter = DateTime.Now.AddMonths(-3).ToString("MM/dd/yyyy")

            ' 只获取最近3个月的最多30封邮件，不再统计总数
            Dim tempRecentMails As New List(Of (Received As DateTime, Subject As String))
            For Each folder In folders
                Try
                    Dim mailFilter = $"[SenderEmailAddress] = '{senderEmail}' AND [ReceivedTime] >= '{dateFilter}'"
                    Dim table As Outlook.Table = folder.GetTable(mailFilter)
                    table.Columns.Add("Subject")
                    table.Columns.Add("ReceivedTime")
                    ' 使用PR_ENTRYID获取长格式EntryID
                    table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")

                    Do Until table.EndOfTable OrElse tempRecentMails.Count >= 30
                        Dim row = table.GetNextRow()
                        Try
                            Dim receivedObj = row("ReceivedTime")
                            Dim subjectObj = row("Subject")
                            Dim received As DateTime = If(receivedObj IsNot Nothing AndAlso Not String.IsNullOrEmpty(receivedObj.ToString()), DateTime.Parse(receivedObj.ToString()), DateTime.MinValue)
                            Dim subject As String = If(subjectObj IsNot Nothing, subjectObj.ToString(), "Unknown Subject")
                            tempRecentMails.Add((received, subject))
                        Catch
                            ' 忽略单个邮件获取错误
                        End Try
                    Loop
                Catch ex As System.Exception
                    Dim folderName As String = "未知文件夹"
                    Me.Invoke(Sub()
                                  folderName = folder.Name
                              End Sub)
                    Debug.WriteLine($"搜索文件夹 {folderName} 时出错: {ex.Message}")
                End Try
            Next

            ' 按时间排序并显示最近邮件，添加序号（不再生成可点击链接）
            recentMails = tempRecentMails.OrderByDescending(Function(m) m.Received).Take(30).ToList()

            swMail.Stop()
            elapsedMailStats = swMail.ElapsedMilliseconds

            info.AppendLine($"邮件往来统计:")
            info.AppendLine($"最近邮件 (最多30封):")

            For i As Integer = 0 To recentMails.Count - 1
                Dim m = recentMails(i)
                info.AppendLine($"- [{i + 1}] {m.Received:yyyy-MM-dd HH:mm} {m.Subject.Replace("[EXT]", "")}")
            Next

            ' 保存到缓存
            Dim result = info.ToString()
            contactMailCache(senderEmail) = (result, DateTime.Now)
            Debug.WriteLine($"性能统计: 联系人 {elapsedContactSearch}ms, 会议 {elapsedMeetingStats}ms, 邮件 {elapsedMailStats}ms")

            Return result  ' 添加返回语句
        Catch ex As System.Exception
            Debug.WriteLine($"获取联系人信息时出错: {ex.Message}")
            Return $"获取联系人信息时出错: {ex.Message}"
        End Try
    End Function

    ' 修改导航事件处理程序

    ' 添加打开邮件的方法
    Private Sub OpenOutlookMail(entryID As String)
        Try
            ' 使用 Application.CreateItem 方法而不是直接获取项目
            ' 这可以避免一些 COM 互操作问题
            Dim mailItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(entryID)
            If mailItem IsNot Nothing Then
                ' 使用 Try-Finally 确保资源释放
                Try
                    mailItem.Display()
                Finally
                    ' 释放 COM 对象
                    If mailItem IsNot Nothing Then
                        Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                    End If
                End Try
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"打开邮件出错: {ex.Message}")
            MessageBox.Show("无法打开邮件，可能已被删除或移动。")
        End Try
    End Sub

    Private Sub SetupTasksTab()
        Dim tabPage2 As New TabPage("任务")
        Dim taskButtonPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 40
        }

        Dim btnAddTask As New Button With {
            .Text = "新建任务",
            .Location = New Point(10, 5),
            .Size = New Size(80, 30)
        }
        AddHandler btnAddTask.Click, AddressOf BtnAddTask_Click
        taskButtonPanel.Controls.Add(btnAddTask)

        taskList = New ListView With {
            .Dock = DockStyle.Fill,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor
        }
        OutlookAddIn3.Handlers.TaskHandler.SetupTaskList(taskList)
        taskList.Columns.Add("主题", 200)
        taskList.Columns.Add("到期日", 100)
        taskList.Columns.Add("状态", 100)
        taskList.Columns.Add("完成百分比", 100)
        taskList.Columns.Add("关联邮件", 200)


        ' Add the event handler here, after taskList is initialized
        AddHandler taskList.DoubleClick, AddressOf TaskList_DoubleClick

        Dim containerPanel As New Panel With {
            .Dock = DockStyle.Fill
        }
        containerPanel.Controls.Add(taskList)
        containerPanel.Controls.Add(taskButtonPanel)
        tabPage2.Controls.Add(containerPanel)
        tabControl.TabPages.Add(tabPage2)
    End Sub

    Private Sub SetupActionsTab()
        Dim tabPage3 As New TabPage("操作")
        btnPanel = New Panel With {
            .Dock = DockStyle.Fill
        }

        ' 创建按钮面板
        Dim buttonPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 20
        }

        ' 创建ListView替代TextBox来展示联系人信息
        contactInfoList = New ListView With {
            .Dock = DockStyle.Fill,
            .View = System.Windows.Forms.View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .MultiSelect = False,
            .HeaderStyle = ColumnHeaderStyle.Clickable,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor
        }

        ' 设置ListView列
        contactInfoList.Columns.Add("类型", 60)
        contactInfoList.Columns.Add("内容", 100) ' 调整宽度为100
        contactInfoList.Columns.Add("详情", 250)

        ' 添加双击事件处理邮件链接
        AddHandler contactInfoList.DoubleClick, AddressOf ContactInfoList_DoubleClick
        ' 添加单击事件处理邮件链接
        AddHandler contactInfoList.Click, AddressOf ContactInfoList_Click

        ' 只创建按钮，不预先创建文本框
        Dim x As Integer = 10
        For i As Integer = 1 To 3
            Dim btn As New Button With {
                .Text = If(i = 1, "联系人信息", $"按钮 {i}"),
                .Location = New Point(x, 2),
                .Size = New Size(100, 15)
            }

            ' 特别处理第一个按钮 - 延迟初始化
            If i = 1 Then
                AddHandler btn.Click, Sub(s, e)
                                          GetContactInfoListHandler()
                                      End Sub
            Else
                AddHandler btn.Click, Sub(s, e)
                                          ' 显示会话信息
                                          contactInfoList.Items.Clear()
                                          Dim item1 As New ListViewItem("会话ID")
                                          item1.SubItems.Add(currentConversationId)
                                          item1.SubItems.Add("当前会话标识")
                                          contactInfoList.Items.Add(item1)

                                          Dim item2 As New ListViewItem("邮件数量")
                                          item2.SubItems.Add(lvMails.Items.Count.ToString())
                                          item2.SubItems.Add("会话中的邮件总数")
                                          contactInfoList.Items.Add(item2)

                                          Dim item3 As New ListViewItem("当前邮件")
                                          item3.SubItems.Add(currentMailEntryID)
                                          item3.SubItems.Add("当前选中的邮件ID")
                                          contactInfoList.Items.Add(item3)
                                      End Sub
            End If

            buttonPanel.Controls.Add(btn)
            x += 125
        Next

        ' 先添加按钮面板到主面板（Dock Top）
        btnPanel.Controls.Add(buttonPanel)
        ' 再添加ListView到主面板（Dock Fill）
        btnPanel.Controls.Add(contactInfoList)

        tabPage3.Controls.Add(btnPanel)
        tabControl.TabPages.Add(tabPage3)
    End Sub

    ' 新增：联系人信息列表支持与双击打开邮件
    Private Async Sub GetContactInfoListHandler()
        Try
            If contactInfoList Is Nothing Then Return

            ' 在开始收集联系人信息时立即抑制 WebView 更新
            suppressWebViewUpdate += 1

            ' 显示进度指示器
            ShowProgress("正在收集联系人信息...")

            contactInfoList.Items.Clear()
            Dim loading As New ListViewItem("状态")
            loading.SubItems.Add("正在收集联系人信息...")
            loading.SubItems.Add("")
            contactInfoList.Items.Add(loading)

            Dim result = Await Task.Run(Function() GetContactInfoData(CancellationToken))

            ' 检查是否被取消
            If CancellationToken.IsCancellationRequested Then
                Return
            End If

            If Me.InvokeRequired Then
                Me.Invoke(Sub() PopulateContactInfoList(result))
            Else
                PopulateContactInfoList(result)
            End If
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("联系人信息收集被取消")
        Catch ex As System.Exception
            Debug.WriteLine("GetContactInfoListHandler error: " & ex.Message)
        Finally
            ' 隐藏进度指示器并释放抑制计数器
            HideProgress()
            suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
        End Try
    End Sub

    ' 生成联系人信息的结构化数据
    Private Function GetContactInfoData(Optional cancellationToken As Threading.CancellationToken = Nothing) As (SenderName As String, SenderEmail As String, MeetingStats As Dictionary(Of String, Integer), Upcoming As List(Of (MeetingDate As DateTime, Title As String, EntryID As String)), MailCount As Integer, RecentMailIds As List(Of (EntryID As String, Subject As String, Received As DateTime)))
        Dim senderName As String = ""
        Dim senderEmail As String = ""
        Dim meetingStats As New Dictionary(Of String, Integer)
        Dim upcoming As New List(Of (DateTime, String, String))
        Dim mailCount As Integer = 0
        Dim recentMails As New List(Of (String, String, DateTime))
        Try
            Dim currentItem As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem Is Nothing Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)

            If TypeOf currentItem Is Outlook.MailItem Then
                Dim mail = DirectCast(currentItem, Outlook.MailItem)
                Try
                    senderEmail = mail.SenderEmailAddress
                    senderName = mail.SenderName
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常获取邮件发件人信息 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
                Catch ex As System.Exception
                    Debug.WriteLine($"获取邮件发件人信息时发生异常: {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
                End Try
            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                Dim meeting = DirectCast(currentItem, Outlook.MeetingItem)
                Try
                    senderEmail = meeting.SenderEmailAddress
                    senderName = meeting.SenderName
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常获取会议发件人信息 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
                Catch ex As System.Exception
                    Debug.WriteLine($"获取会议发件人信息时发生异常: {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
                End Try
            End If
            If String.IsNullOrEmpty(senderEmail) Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)

            ' 会议统计
            Dim calendar As Outlook.Folder = Nothing
            Dim meetings As Outlook.Items = Nothing
            Try
                calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
                Dim startDate = DateTime.Now.AddMonths(-2)
                Dim endDate = DateTime.Now.AddMonths(1)
                Dim meetingFilter = $"[Start] >= '{startDate:MM/dd/yyyy}' AND [End] <= '{endDate:MM/dd/yyyy}'"
                meetings = calendar.Items.Restrict(meetingFilter)
            Catch ex As System.Runtime.InteropServices.COMException
                Debug.WriteLine($"COM异常获取日历文件夹 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
            Catch ex As System.Exception
                Debug.WriteLine($"获取日历文件夹时发生异常: {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
            End Try

            If meetings Is Nothing Then
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
            End If
            Dim meetingsCount As Integer = meetings.Count
            For i = meetingsCount To 1 Step -1
                Dim ap As Outlook.AppointmentItem = Nothing
                Dim requiredAttendees As String = String.Empty
                Dim optionalAttendees As String = String.Empty
                Dim subject As String = String.Empty
                Dim startTime As DateTime
                Dim entryId As String = String.Empty

                Try
                    ap = DirectCast(meetings(i), Outlook.AppointmentItem)
                    If ap IsNot Nothing Then
                        requiredAttendees = ap.RequiredAttendees
                        optionalAttendees = ap.OptionalAttendees
                        subject = ap.Subject
                        startTime = ap.Start
                        entryId = ap.EntryID
                    End If
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问会议项属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Continue For
                Catch ex As System.Exception
                    Debug.WriteLine($"访问会议项属性时发生异常: {ex.Message}")
                    Continue For
                End Try

                If ap IsNot Nothing AndAlso Not String.IsNullOrEmpty(requiredAttendees) AndAlso (requiredAttendees.Contains(senderEmail) OrElse (Not String.IsNullOrEmpty(optionalAttendees) AndAlso optionalAttendees.Contains(senderEmail))) Then
                    Dim projectName = "其他"
                    Dim match = System.Text.RegularExpressions.Regex.Match(subject, "\[(.*?)\]")
                    If match.Success Then projectName = match.Groups(1).Value
                    If meetingStats.ContainsKey(projectName) Then
                        meetingStats(projectName) += 1
                    Else
                        meetingStats.Add(projectName, 1)
                    End If
                    If startTime > DateTime.Now Then
                        upcoming.Add((startTime, subject, entryId))
                    End If
                End If
            Next

            ' 邮件统计
            Dim folders As New List(Of Outlook.Folder)
            Try
                Dim store As Outlook.Store = Globals.ThisAddIn.Application.Session.DefaultStore
                If store IsNot Nothing Then
                    GetAllMailFolders(store.GetRootFolder(), folders)
                End If
            Catch ex As System.Runtime.InteropServices.COMException
                Debug.WriteLine($"COM异常获取邮件存储 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
            Catch ex As System.Exception
                Debug.WriteLine($"获取邮件存储时发生异常: {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
            End Try

            If folders.Count = 0 Then
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
            End If
            Dim dateFilter = DateTime.Now.AddMonths(-3).ToString("MM/dd/yyyy")
            Dim tasks As New List(Of Task(Of (Count As Integer, Mails As List(Of (EntryID As String, Subject As String, Received As DateTime)))))
            For Each folder In folders
                tasks.Add(Task.Run(Function()
                                       Try
                                           Dim mailFilter = $"[SenderEmailAddress] = '{senderEmail}' AND [ReceivedTime] >= '{dateFilter}'"
                                           Dim table As Outlook.Table = folder.GetTable(mailFilter)
                                           table.Columns.Add("Subject")
                                           table.Columns.Add("ReceivedTime")
                                           ' 使用PR_ENTRYID获取长格式EntryID
                                           table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")
                                           Dim folderMails As New List(Of (String, String, DateTime))
                                           Dim count As Integer = 0
                                           Dim endOfTable As Boolean
                                           Dim row As Outlook.Row
                                           Do
                                               row = table.GetNextRow()
                                               endOfTable = table.EndOfTable
                                               If row Is Nothing Then Exit Do
                                               count += 1
                                               If folderMails.Count < 50 Then
                                                   Try
                                                       Dim entryIdObj = row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")
                                                       Dim entryId As String = ConvertEntryIDToString(entryIdObj)
                                                       Dim subject As String = TryCast(row("Subject"), String)
                                                       Dim received As DateTime = DateTime.Parse(row("ReceivedTime").ToString())
                                                       folderMails.Add((entryId, If(subject, ""), received))
                                                   Catch
                                                   End Try
                                               End If
                                           Loop While Not endOfTable
                                           Return (count, folderMails)
                                       Catch ex As System.Exception
                                           Dim folderName As String = "未知文件夹"
                                           Try
                                               folderName = folder.Name
                                           Catch
                                               ' 如果无法获取文件夹名称，使用默认值
                                           End Try
                                           Debug.WriteLine($"搜索文件夹 {folderName} 时出错: {ex.Message}")
                                           Return (0, New List(Of (String, String, DateTime)))
                                       End Try
                                   End Function))
            Next
            Dim searchResults = Task.WhenAll(tasks).Result
            For Each r In searchResults
                mailCount += r.Count
                recentMails.AddRange(r.Mails)
            Next
            recentMails = recentMails.OrderByDescending(Function(m) m.Item3).Take(50).ToList()
        Catch ex As System.Exception
            Debug.WriteLine("GetContactInfoData error: " & ex.Message)
        End Try
        Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails)
    End Function

    Private Sub PopulateContactInfoList(result As (SenderName As String, SenderEmail As String, MeetingStats As Dictionary(Of String, Integer), Upcoming As List(Of (MeetingDate As DateTime, Title As String, EntryID As String)), MailCount As Integer, RecentMailIds As List(Of (EntryID As String, Subject As String, Received As DateTime))))
        ' 在填充联系人列表期间抑制 WebView 更新
        suppressWebViewUpdate += 1
        contactInfoList.BeginUpdate()
        Try
            contactInfoList.Items.Clear()
            Dim i1 As New ListViewItem("发件人")
            i1.SubItems.Add(result.SenderName)
            i1.SubItems.Add("")
            contactInfoList.Items.Add(i1)
            Dim i2 As New ListViewItem("邮箱")
            i2.SubItems.Add(result.SenderEmail)
            i2.SubItems.Add("")
            contactInfoList.Items.Add(i2)

            Dim totalMeetings = result.MeetingStats.Values.Sum()
            Dim i3 As New ListViewItem("会议(近2月)")
            i3.SubItems.Add($"总会议数: {totalMeetings}")
            i3.SubItems.Add("")
            contactInfoList.Items.Add(i3)
            For Each kv In result.MeetingStats.OrderByDescending(Function(x) x.Value)
                Dim it As New ListViewItem("项目")
                it.SubItems.Add(kv.Key)
                it.SubItems.Add($"{kv.Value}次")
                contactInfoList.Items.Add(it)
            Next
            For Each up In result.Upcoming.OrderBy(Function(x) x.MeetingDate).Take(3)
                Dim it As New ListViewItem("即将会议")
                it.SubItems.Add(up.MeetingDate.ToString("MM/dd HH:mm"))
                it.SubItems.Add(up.Title)
                it.Tag = up.EntryID ' 将EntryID存储在Tag中
                contactInfoList.Items.Add(it)
            Next

            Dim i4 As New ListViewItem("邮件往来")
            i4.SubItems.Add($"总邮件数: {result.MailCount}")
            i4.SubItems.Add("")
            contactInfoList.Items.Add(i4)

            For Each m In result.RecentMailIds
                Dim mailItem As New ListViewItem("最近邮件")
                mailItem.SubItems.Add(m.Received.ToString("yyyy-MM-dd HH:mm"))
                mailItem.SubItems.Add(m.Subject.Replace("[EXT]", ""))
                mailItem.Tag = m.EntryID
                contactInfoList.Items.Add(mailItem)
            Next
        Finally
            contactInfoList.EndUpdate()
            suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
        End Try
    End Sub

    Private Sub ContactInfoList_DoubleClick(sender As Object, e As EventArgs)
        Try
            ' 抑制模式下不响应双击
            If suppressWebViewUpdate > 0 Then Return

            If contactInfoList.SelectedItems.Count = 0 Then Return
            Dim item = contactInfoList.SelectedItems(0)
            Dim entryId = TryCast(item.Tag, String)
            If Not String.IsNullOrEmpty(entryId) Then
                ' 增加隔离标志，避免与 lvMails 联动或触发 WebView 刷新冲突
                suppressWebViewUpdate += 1
                Try
                    SafeOpenOutlookMail(entryId)
                Finally
                    suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
                End Try
            End If
        Catch ex As System.Exception
            Debug.WriteLine("ContactInfoList_DoubleClick error: " & ex.Message)
        End Try
    End Sub

    ' 保留原有：RichTextBox链接点击（若有其他地方复用）
    Private Sub OutputTextBox_LinkClicked(sender As Object, e As LinkClickedEventArgs)
        Try
            Process.Start(New ProcessStartInfo With {
                .FileName = e.LinkText,
                .UseShellExecute = True
            })
        Catch ex As System.Exception
            Debug.WriteLine($"处理链接点击时出错: {ex.Message}")
        End Try
    End Sub

    Private Sub ContactInfoList_Click(sender As Object, e As EventArgs)
        Try
            ' 允许在本窗格中点击联系人邮件时总是更新右侧 mailBrowser
            ' 抑制标志仅用于避免与外部触发的刷新串扰，不用于本地点击后的内容展示

            If contactInfoList.SelectedItems.Count = 0 Then Return
            Dim item = contactInfoList.SelectedItems(0)
            Dim entryId = TryCast(item.Tag, String)
            If Not String.IsNullOrEmpty(entryId) Then
                ' 本地点击不抬高抑制计数（保持为局部更新）
                Try
                    Dim mailItem As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(entryId)
                    Dim displayContent As String = ""
                    If TypeOf mailItem Is Outlook.MailItem Then
                        Dim mail As Outlook.MailItem = DirectCast(mailItem, Outlook.MailItem)
                        Try
                            Dim subject As String = If(String.IsNullOrEmpty(mail.Subject), "无主题", mail.Subject)
                            Dim senderName As String = If(String.IsNullOrEmpty(mail.SenderName), "未知", mail.SenderName)
                            Dim receivedTime As String = If(mail.ReceivedTime = DateTime.MinValue, "未知", mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
                            Dim htmlBody As String = If(String.IsNullOrEmpty(mail.HTMLBody), "", ReplaceTableTag(mail.HTMLBody))

                            displayContent = $"<html><body style='font-family: Arial; padding: 10px; Font-size:12px;'>" &
                                $"<h4 style='color: var(--theme-color, #0078d7);'>{subject}</h4>" &
                                $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                                $"<strong style='color: var(--theme-color, #0078d7);'>发件人:</strong> {senderName}<br/>" &
                                $"<strong style='color: var(--theme-color, #0078d7);'>时间:</strong> {receivedTime}" &
                                $"</div>" &
                                $"<div style='border-top: 1px solid var(--theme-color, #0078d7); padding-top: 10px;'>" &
                                $"<style>.hidden-table {{display: none;}} img {{display: none;}}</style>" &
                                $"{htmlBody}" &
                                $"</div>" &
                                "</body></html>"
                        Catch ex As System.Runtime.InteropServices.COMException
                            Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                            displayContent = "<html><body style='font-family: Arial; padding: 10px;'>无法访问邮件属性</body></html>"
                        Catch ex As System.Exception
                            Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                            displayContent = "<html><body style='font-family: Arial; padding: 10px;'>无法访问邮件属性</body></html>"
                        End Try
                        'displayContent = $"<h1>{mail.Subject}</h1><p><b>发件人:</b> {mail.SenderName}</p><p><b>时间:</b> {mail.ReceivedTime}</p><hr>{mail.HTMLBody}"
                    ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                        Dim appointment As Outlook.AppointmentItem = DirectCast(mailItem, Outlook.AppointmentItem)
                        Try
                            Dim subject As String = If(String.IsNullOrEmpty(appointment.Subject), "无主题", appointment.Subject)
                            Dim organizer As String = If(String.IsNullOrEmpty(appointment.Organizer), "未知", appointment.Organizer)
                            Dim startTime As String = appointment.Start.ToString("yyyy-MM-dd HH:mm:ss")
                            Dim body As String = If(String.IsNullOrEmpty(appointment.Body), "", ReplaceTableTag(appointment.Body))

                            displayContent = $"<html><body style='font-family: Arial; padding: 10px; Font-size:12px;'>" &
                                $"<h4 style='color: var(--theme-color, #0078d7);'>{subject}</h4>" &
                                $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                                $"<strong style='color: var(--theme-color, #0078d7);'>组织者:</strong> {organizer}<br/>" &
                                $"<strong style='color: var(--theme-color, #0078d7);'>时间:</strong> {startTime}" &
                                $"</div>" &
                                $"<div style='border-top: 1px solid var(--theme-color, #0078d7); padding-top: 10px;'>" &
                                $"<style>.hidden-table {{display: none;}} img {{display: none;}}</style>" &
                                $"{body}" &
                                $"</div>" &
                                "</body></html>"
                        Catch ex As System.Runtime.InteropServices.COMException
                            Debug.WriteLine($"COM异常访问会议属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                            displayContent = "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议属性</body></html>"
                        Catch ex As System.Exception
                            Debug.WriteLine($"访问会议属性时发生异常: {ex.Message}")
                            displayContent = "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议属性</body></html>"
                        End Try
                        'displayContent = $"<h4>{appointment.Subject}</h4><p><b>组织者:</b> {appointment.Organizer}</p><p><b>时间:</b> {appointment.Start}</p><hr>{appointment.Body}"
                    End If
                    ' 本地点击：始终更新当前窗格的 WebView
                    mailBrowser.DocumentText = displayContent
                    'Else
                    '    Debug.WriteLine("无法获取邮件项或邮件项不是MailItem/AppointmentItem类型。")
                    'End If
                Catch ex As System.Exception
                    Debug.WriteLine("获取邮件HTML内容时出错: " & ex.Message)
                Finally
                    ' 本地点击不再修改抑制计数
                End Try
            End If
        Catch ex As System.Exception
            Debug.WriteLine("ContactInfoList_Click error: " & ex.Message)
        End Try
    End Sub

    Private Sub SafeOpenOutlookMail(entryID As String)
        Try
            Debug.WriteLine($"尝试快速打开邮件，EntryID: {If(entryID?.Length > 10, entryID.Substring(0, 10) & "...", "null")}")

            ' 检查EntryID是否有效
            If String.IsNullOrEmpty(entryID) Then
                Debug.WriteLine("EntryID为空")
                Return
            End If

            ' 抑制 WebView 更新以避免打开邮件时触发额外刷新
            Dim wasSupressed = IsWebViewUpdateSuppressed
            If Not wasSupressed Then
                suppressWebViewUpdate += 1
                Debug.WriteLine("已抑制 WebView 更新以提升邮件打开速度")
            End If

            Try
                ' 使用优化的快速打开方法（支持 StoreID）
                ' TODO: 如果在 Flag 任务中有 StoreID 信息，可以传入第二个参数进一步提升性能
                Dim success = OutlookAddIn3.Utils.OutlookUtils.FastOpenMailItem(entryID)

                If success Then
                    Debug.WriteLine("快速邮件打开成功")
                Else
                    Debug.WriteLine("快速邮件打开失败，尝试兜底方法")

                    ' 兜底：使用原有方法
                    Dim mailItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(entryID)
                    If mailItem IsNot Nothing Then
                        Try
                            If TypeOf mailItem Is Outlook.MailItem Then
                                DirectCast(mailItem, Outlook.MailItem).Display(False)
                            ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                                DirectCast(mailItem, Outlook.AppointmentItem).Display(False)
                            ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                                DirectCast(mailItem, Outlook.MeetingItem).Display(False)
                            ElseIf TypeOf mailItem Is Outlook.TaskItem Then
                                DirectCast(mailItem, Outlook.TaskItem).Display(False)
                            Else
                                CallByName(mailItem, "Display", CallType.Method, False)
                            End If
                            Debug.WriteLine("兜底方法邮件打开成功")
                        Finally
                            OutlookAddIn3.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                        End Try
                    End If
                End If
            Finally
                ' 延迟恢复 WebView 更新（避免邮件打开过程中的干扰）
                If Not wasSupressed Then
                    Task.Run(Async Function()
                                 Await Task.Delay(500) ' 等待邮件窗口完全打开
                                 Try
                                     If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
                                         Me.BeginInvoke(Sub() suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1))
                                     End If
                                 Catch ex As System.Exception
                                     Debug.WriteLine($"恢复 WebView 更新时出错: {ex.Message}")
                                 End Try
                                 Return Nothing
                             End Function)
                    Debug.WriteLine("已安排延迟恢复 WebView 更新")
                End If
            End Try

        Catch ex As System.Exception
            Debug.WriteLine($"安全打开邮件时出错: {ex.Message}")
        End Try
    End Sub

    ' 将异步逻辑移到单独的方法中
    ' 将异步逻辑移到单独的方法中
    Private Async Function GetContactInfoHandler(outputTextBox As Control) As Task(Of String)
        Dim info As String = String.Empty
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub()
                              If TypeOf outputTextBox Is TextBox Then
                                  DirectCast(outputTextBox, TextBox).Text = "正在收集联系人信息..."
                              ElseIf TypeOf outputTextBox Is RichTextBox Then
                                  DirectCast(outputTextBox, RichTextBox).Text = "正在收集联系人信息..."
                              End If
                          End Sub)
            Else
                If TypeOf outputTextBox Is TextBox Then
                    DirectCast(outputTextBox, TextBox).Text = "正在收集联系人信息..."
                ElseIf TypeOf outputTextBox Is RichTextBox Then
                    DirectCast(outputTextBox, RichTextBox).Text = "正在收集联系人信息..."
                End If
            End If

            ' 在后台线程中执行耗时的Outlook操作
            info = Await GetContactInfoAsync()

            If Me.InvokeRequired Then
                Me.Invoke(Sub()
                              If Not String.IsNullOrEmpty(info) Then
                                  If TypeOf outputTextBox Is TextBox Then
                                      DirectCast(outputTextBox, TextBox).Text = info
                                  ElseIf TypeOf outputTextBox Is RichTextBox Then
                                      DirectCast(outputTextBox, RichTextBox).Text = info
                                  End If
                              Else
                                  If TypeOf outputTextBox Is TextBox Then
                                      DirectCast(outputTextBox, TextBox).Text = "未能获取联系人信息"
                                  ElseIf TypeOf outputTextBox Is RichTextBox Then
                                      DirectCast(outputTextBox, RichTextBox).Text = "未能获取联系人信息"
                                  End If
                              End If
                          End Sub)
            Else
                If Not String.IsNullOrEmpty(info) Then
                    If TypeOf outputTextBox Is TextBox Then
                        DirectCast(outputTextBox, TextBox).Text = info
                    ElseIf TypeOf outputTextBox Is RichTextBox Then
                        DirectCast(outputTextBox, RichTextBox).Text = info
                    End If
                Else
                    If TypeOf outputTextBox Is TextBox Then
                        DirectCast(outputTextBox, TextBox).Text = "未能获取联系人信息"
                    ElseIf TypeOf outputTextBox Is RichTextBox Then
                        DirectCast(outputTextBox, RichTextBox).Text = "未能获取联系人信息"
                    End If
                End If
            End If
        Catch ex As System.Exception
            If Me.InvokeRequired Then
                Me.Invoke(Sub()
                              If TypeOf outputTextBox Is TextBox Then
                                  DirectCast(outputTextBox, TextBox).Text = $"获取联系人信息时出错: {ex.Message}"
                              ElseIf TypeOf outputTextBox Is RichTextBox Then
                                  DirectCast(outputTextBox, RichTextBox).Text = $"获取联系人信息时出错: {ex.Message}"
                              End If
                          End Sub)
            Else
                If TypeOf outputTextBox Is TextBox Then
                    DirectCast(outputTextBox, TextBox).Text = $"获取联系人信息时出错: {ex.Message}"
                ElseIf TypeOf outputTextBox Is RichTextBox Then
                    DirectCast(outputTextBox, RichTextBox).Text = $"获取联系人信息时出错: {ex.Message}"
                End If
            End If
            Debug.WriteLine($"获取联系人信息时出错: {ex.Message}")
            Return $"获取联系人信息时出错: {ex.Message}"
        End Try
        Return info
    End Function

    Private Function IsNetworkAvailable() As Boolean
        Try
            Return System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()
        Catch ex As System.Exception
            Debug.WriteLine($"检查网络连接出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Async Function CheckWolaiRecordAsync(conversationId As String) As Task(Of String)
        Try
            ShowProgress("正在检查Wolai记录...")
            CancellationToken.ThrowIfCancellationRequested()
            Dim noteList As New List(Of (CreateTime As String, Title As String, Link As String))
            ' 首先检查所有相关邮件的属性
            Try
                ' 获取当前会话的所有邮件

                Dim currentItem As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                Dim conversation As Outlook.Conversation = Nothing

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                End If


                If conversation IsNot Nothing Then
                    Dim table As Outlook.Table = conversation.GetTable()
                    ' 优化：只添加需要的列，减少数据传输
                    table.Columns.RemoveAll() ' 移除默认列
                    ' 使用PR_ENTRYID获取长格式EntryID
                    table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")

                    ' 遍历会话中的所有项目
                    Do Until table.EndOfTable
                        Dim item As Object = Nothing  ' Declare item at the beginning of the loop
                        Try
                            Dim row As Outlook.Row = table.GetNextRow()
                            Dim entryIdStr As String = ConvertEntryIDToString(row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"))
                            item = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(entryIdStr)

                            ' 检查所有支持 UserProperties 的项目类型
                            If TypeOf item Is Outlook.MailItem OrElse
                            TypeOf item Is Outlook.AppointmentItem OrElse
                            TypeOf item Is Outlook.MeetingItem Then

                                Try
                                    Dim userProps = CallByName(item, "UserProperties", CallType.Get)
                                    Dim wolaiProp = userProps.Find("WolaiNoteLink")
                                    Dim createTimeProp = userProps.Find("WolaiNoteCreateTime")

                                    If wolaiProp IsNot Nothing Then
                                        Dim wolaiLink = wolaiProp.Value.ToString()
                                        Dim itemSubject = CallByName(item, "Subject", CallType.Get)
                                        Dim createTime = If(createTimeProp IsNot Nothing,
                                                            createTimeProp.Value.ToString(),
                                                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                                        Debug.WriteLine($"从项目属性中找到 Wolai 链接: {wolaiLink}")

                                        ' 避免重复添加相同的链接
                                        If Not noteList.Any(Function(n) n.Link = wolaiLink) Then
                                            noteList.Add((createTime, itemSubject, wolaiLink))
                                        End If
                                    End If
                                Catch ex As System.Exception
                                    Debug.WriteLine($"检查项目属性时出错: {ex.Message}")
                                End Try
                            End If
                        Catch ex As System.Exception
                            Debug.WriteLine($"处理项目是否存在 wolai 链接时出错: {ex.Message}")
                            Continue Do
                        Finally
                            If item IsNot Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(item)
                            End If
                        End Try
                    Loop
                    ' #todo: task,  meeting, 是否能刷出来对应note? 只要能有list(属于conversation)的: appointment, mail 可以.  
                Else

                    ' 检查所有支持 UserProperties 的项目类型
                    If TypeOf currentItem Is Outlook.TaskItem Then

                        Try
                            Dim userProps = CallByName(currentItem, "UserProperties", CallType.Get)
                            Dim wolaiProp = userProps.Find("WolaiNoteLink")
                            Dim createTimeProp = userProps.Find("WolaiNoteCreateTime")

                            If wolaiProp IsNot Nothing Then
                                Dim wolaiLink = wolaiProp.Value.ToString()
                                Dim itemSubject = CallByName(currentItem, "Subject", CallType.Get)
                                Dim createTime = If(createTimeProp IsNot Nothing,
                                                            createTimeProp.Value.ToString(),
                                                            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                                Debug.WriteLine($"从项目属性中找到 Wolai 链接: {wolaiLink}")

                                ' 避免重复添加相同的链接
                                If Not noteList.Any(Function(n) n.Link = wolaiLink) Then
                                    noteList.Add((createTime, itemSubject, wolaiLink))
                                End If
                            End If
                        Catch ex As System.Exception
                            Debug.WriteLine($"检查项目属性时出错: {ex.Message}")
                        End Try
                    End If
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"检查邮件属性时出错??: {ex.Message}")
            End Try

            ' 如果邮件属性中没有找到，且网络可用，则进行网络查询
            'If Not IsNetworkAvailable() Then
            '    Debug.WriteLine("网络不可用，跳过网络查询")
            UpdateNoteList(noteList)
            Return String.Empty
            'End If

            Using client As New HttpClient()
                ' 获取 token
                Dim tokenData As New JObject()
                tokenData.Add("", "2NdHab5WdUG995izevb69b")
                tokenData.Add("appSecret", "ffa888d4ebd73bae77a77abebcacf80001654b3f19d4ffbbcc3c41cbe0bed645")

                Dim tokenContent = New StringContent(tokenData.ToString(), Encoding.UTF8, "application/json")
                Dim tokenResponse = Await client.PostAsync("https://openapi.wolai.com/v1/token", tokenContent)

                If Not tokenResponse.IsSuccessStatusCode Then
                    Debug.WriteLine("获取令牌失败")
                    Return String.Empty
                End If

                Dim tokenResult = Await tokenResponse.Content.ReadAsStringAsync()
                Dim tokenJson = JObject.Parse(tokenResult)
                Dim appToken = tokenJson.SelectToken("data.app_token")?.ToString()

                If String.IsNullOrEmpty(appToken) Then
                    Debug.WriteLine("获取令牌为空")
                    Return String.Empty
                End If

                ' 查询数据
                client.DefaultRequestHeaders.Clear()
                client.DefaultRequestHeaders.Add("Authorization", appToken)

                ' 构建查询参数
                Dim queryData As New JObject()
                queryData.Add("filter", New JObject From {
                    {"property", "ConvID"},
                    {"value", conversationId},
                    {"type", "text"},
                    {"operator", "equals"}
                })

                Dim queryContent = New StringContent(queryData.ToString(), Encoding.UTF8, "application/json")
                Dim queryResponse = Await client.PostAsync("https://openapi.wolai.com/v1/databases/pLEYWMtYy4xFRzTyLEewrX/query", queryContent)

                If queryResponse.IsSuccessStatusCode Then
                    Dim responseContent = Await queryResponse.Content.ReadAsStringAsync()
                    Dim responseJson = JObject.Parse(responseContent)
                    Dim rows = responseJson.SelectToken("data")

                    If rows IsNot Nothing AndAlso rows.HasValues Then

                        For Each row In rows
                            Dim pageId = row.ToString().Split("/"c).Last()
                            Dim wolaiLink = $"https://www.wolai.com/{pageId}"
                            Dim title = row.Parent.Parent("Title")?.ToString()
                            Dim createTime = row.Parent.Parent("Created Time")?.ToString()
                            ' 避免重复添加
                            If Not noteList.Any(Function(n) n.Link = wolaiLink) Then
                                noteList.Add((createTime, title, wolaiLink))
                            End If
                        Next

                        UpdateNoteList(noteList)
                        Return String.Empty
                    End If
                End If

                UpdateNoteList(noteList)  ' Update ListView even if no results
                Return String.Empty
            End Using
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("Wolai记录检查被取消")
            Return "操作已取消"
        Catch ex As System.Exception
            Debug.WriteLine($"CheckWolaiRecord 执行出错: {ex.Message}")
            Return String.Empty
        Finally
            HideProgress()
        End Try
    End Function

    Private Function GenerateHtmlContent(noteList As List(Of (CreateTime As String, Title As String, Link As String))) As String
        Dim htmlContent As New StringBuilder()
        htmlContent.AppendLine("<html><body style='font-family: Arial; padding: 10px; font-size: 12px;'>")
        'htmlContent.AppendLine("<h3 style='font-size: 14px; margin: 0 0 10px 0;'>已存在的笔记记录：</h3>")
        htmlContent.AppendLine("<table style='width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 12px;'>")
        htmlContent.AppendLine("<tr style='background-color: #f2f2f2;'>")
        htmlContent.AppendLine("<th style='padding: 4px; border: 1px solid #ddd; text-align: left; font-size: 12px;'>创建日期</th>")
        htmlContent.AppendLine("<th style='padding: 4px; border: 1px solid #ddd; text-align: left; font-size: 12px;'>标题</th>")
        htmlContent.AppendLine("<th style='padding: 4px; border: 1px solid #ddd; text-align: left; font-size: 12px;'>操作</th>")
        htmlContent.AppendLine("</tr>")

        For Each note In noteList
            htmlContent.AppendLine("<tr>")
            htmlContent.AppendLine($"<td style='padding: 4px; border: 1px solid #ddd; font-size: 12px;'>{If(note.CreateTime, DateTime.Now.ToString("yyyy-MM-dd HH:mm"))}</td>")
            htmlContent.AppendLine($"<td style='padding: 4px; border: 1px solid #ddd; font-size: 12px;'>{If(note.Title, "无标题")}</td>")
            htmlContent.AppendLine($"<td style='padding: 4px; border: 1px solid #ddd; font-size: 12px;'>")
            htmlContent.AppendLine($"<a href='{note.Link}' target='_blank' onclick='window.open(this.href); return false;' style='font-size: 12px;'>打开笔记</a>")
            htmlContent.AppendLine("</td>")
            htmlContent.AppendLine("</tr>")
        Next

        htmlContent.AppendLine("</table>")
        htmlContent.AppendLine($"<div style='margin-top: 10px; font-size: 12px;'><a href='https://www.wolai.com/autolab/pLEYWMtYy4xFRzTyLEewrX' target='_blank' onclick='window.open(this.href); return false;'>所有笔记</a></div>")
        htmlContent.AppendLine("</body></html>")

        Return htmlContent.ToString()
    End Function

    Private Async Function SaveToWolaiAsync(conversationId As String, conversationTitle As String) As Task(Of Boolean)
        Try
            ShowProgress("正在保存到Wolai...")
            CancellationToken.ThrowIfCancellationRequested()
            Using client As New HttpClient()
                ' 获取 token
                Dim tokenData As New JObject()
                tokenData.Add("appId", "2NdHab5WdUG995izevb69b")
                tokenData.Add("appSecret", "ffa888d4ebd73bae77a77abebcacf80001654b3f19d4ffbbcc3c41cbe0bed645")

                Dim tokenContent = New StringContent(tokenData.ToString(), Encoding.UTF8, "application/json")
                Dim tokenResponse = Await client.PostAsync("https://openapi.wolai.com/v1/token", tokenContent)

                If Not tokenResponse.IsSuccessStatusCode Then
                    MessageBox.Show("获取令牌失败")
                    Return False
                End If

                Dim tokenResult = Await tokenResponse.Content.ReadAsStringAsync()
                Dim tokenJson = JObject.Parse(tokenResult)
                Dim appToken = tokenJson.SelectToken("data.app_token")?.ToString()

                If String.IsNullOrEmpty(appToken) Then
                    MessageBox.Show("获取令牌失败")
                    Return False
                End If

                ' 保存数据
                client.DefaultRequestHeaders.Clear()
                client.DefaultRequestHeaders.Add("Authorization", appToken)

                Dim saveData As New JObject()
                Dim rows As New JArray()
                Dim row As New JObject()
                row.Add("Title", conversationTitle)
                row.Add("URL", "undefined")
                row.Add("ConvID", conversationId)
                rows.Add(row)
                saveData.Add("rows", rows)

                Dim saveContent = New StringContent(saveData.ToString(), Encoding.UTF8, "application/json")
                Dim saveResponse = Await client.PostAsync("https://openapi.wolai.com/v1/databases/pLEYWMtYy4xFRzTyLEewrX/rows", saveContent)

                If saveResponse.IsSuccessStatusCode Then
                    'MessageBox.Show("保存成功")
                    Dim responseContent = Await saveResponse.Content.ReadAsStringAsync()
                    Dim responseJson = JObject.Parse(responseContent)

                    ' 从响应中获取 page_id
                    Dim pageUrl = responseJson.SelectToken("data[0]")?.ToString()
                    Dim pageId = If(Not String.IsNullOrEmpty(pageUrl),
                                  pageUrl.Split("/"c).Last(),
                                  Nothing)

                    If Not String.IsNullOrEmpty(pageId) Then
                        ' 构建 Wolai 页面链接（使用 page_id）
                        Dim wolaiLink = $"https://www.wolai.com/{pageId}"

                        ' 保存链接到邮件属性
                        Try
                            Dim item As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                            If item IsNot Nothing Then
                                ' 检查是否支持 UserProperties
                                If TypeOf item Is Outlook.MailItem OrElse
                                TypeOf item Is Outlook.AppointmentItem OrElse
                                TypeOf item Is Outlook.MeetingItem Then

                                    ' 尝试添加属性
                                    Try
                                        Dim userProps = CallByName(item, "UserProperties", CallType.Get)

                                        ' Link
                                        userProps.Add("WolaiNoteLink", Outlook.OlUserPropertyType.olText, True, Outlook.OlFormatText.olFormatTextText)
                                        userProps("WolaiNoteLink").Value = wolaiLink

                                        ' 添加创建时间字段
                                        userProps.Add("WolaiNoteCreateTime", Outlook.OlUserPropertyType.olText, True, Outlook.OlFormatText.olFormatTextText)
                                        userProps("WolaiNoteCreateTime").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

                                        CallByName(item, "Save", CallType.Method)
                                        Debug.WriteLine($"已保存 Wolai 链接到项目属性: {wolaiLink}")
                                    Catch ex As System.Exception
                                        Debug.WriteLine($"添加属性时出错: {ex.Message}")
                                    End Try
                                End If
                            End If
                        Catch ex As System.Exception
                            Debug.WriteLine($"保存链接到项目属性时出错: {ex.Message}")
                        End Try

                        ' Update the ListView with the new note
                        Dim noteList As New List(Of (CreateTime As String, Title As String, Link As String)) From {
                            (DateTime.Now.ToString("yyyy-MM-dd HH:mm"), conversationTitle, wolaiLink)
                        }
                        UpdateNoteList(noteList)

                        'MessageBox.Show($"保存成功！笔记链接：{wolaiLink}")
                        Debug.WriteLine($"创建记录成功，page_id: {pageId}")
                        Return True
                    Else
                        MessageBox.Show("保存成功，但未能获取记录链接")
                        Debug.WriteLine($"API 响应内容: {responseContent}")
                    End If
                    Return True
                Else
                    Dim errorResult = Await saveResponse.Content.ReadAsStringAsync()
                    MessageBox.Show($"保存失败: {errorResult}")
                    Return False
                End If
                Return True  ' Add appropriate return value
            End Using

        Catch ex As System.OperationCanceledException
            Debug.WriteLine("保存到Wolai被取消")
            MessageBox.Show("操作已取消")
            Return False
        Catch ex As System.Exception
            Debug.WriteLine($"SaveToWolai 执行出错: {ex.Message}")
            MessageBox.Show($"保存失败: {ex.Message}")
            Return False
        Finally
            HideProgress()
        End Try
    End Function

    <System.Runtime.InteropServices.ComVisible(True)>
    Public Sub OpenLink(url As String)
        Try
            Process.Start(New ProcessStartInfo With {
                .FileName = url,
                .UseShellExecute = True
            })
        Catch ex As System.Exception
            Debug.WriteLine($"打开链接出错: {ex.Message}")
            MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。")
        End Try
    End Sub

    Private Async Sub btnNewNote_Click(sender As Object, e As EventArgs)
        Try
            ' 在后台线程中获取邮件主题，避免阻塞UI
            Dim subject As String = Await Task.Run(Function()
                                                       Try
                                                           Dim mailItem As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                                                           If mailItem IsNot Nothing Then
                                                               ' 根据不同类型获取主题
                                                               Try
                                                                   If TypeOf mailItem Is Outlook.MailItem Then
                                                                       Return DirectCast(mailItem, Outlook.MailItem).Subject
                                                                   ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                                                                       Return DirectCast(mailItem, Outlook.AppointmentItem).Subject
                                                                   ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                                                                       Return DirectCast(mailItem, Outlook.MeetingItem).Subject
                                                                   ElseIf TypeOf mailItem Is Outlook.TaskItem Then
                                                                       Return DirectCast(mailItem, Outlook.TaskItem).Subject
                                                                   End If
                                                               Catch ex As System.Runtime.InteropServices.COMException
                                                                   Debug.WriteLine($"COM异常访问项目主题 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                                                                   Return "无法访问主题"
                                                               Catch ex As System.Exception
                                                                   Debug.WriteLine($"访问项目主题时发生异常: {ex.Message}")
                                                                   Return "无法访问主题"
                                                               End Try
                                                           End If
                                                           Return ""
                                                       Catch ex As System.Exception
                                                           Debug.WriteLine($"获取邮件主题时出错: {ex.Message}")
                                                           Return ""
                                                       End Try
                                                   End Function)

            Await SaveToWolaiAsync(currentConversationId, subject)
        Catch ex As System.Exception
            Debug.WriteLine($"btnNewNote_Click error: {ex.Message}")
            MessageBox.Show($"创建笔记时出错: {ex.Message}")
        End Try
    End Sub

    Private Sub BindEvents()
        AddHandler lvMails.SelectedIndexChanged, AddressOf lvMails_SelectedIndexChanged
        AddHandler lvMails.ColumnClick, AddressOf lvMails_ColumnClick
        AddHandler lvMails.DoubleClick, AddressOf lvMails_DoubleClick
        If mailBrowser IsNot Nothing Then
            AddHandler mailBrowser.Navigating, AddressOf MailBrowser_Navigating
        End If
    End Sub

    Private Sub MailBrowser_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
        Try
            If e.Url Is Nothing Then Return
            Dim urlStr As String = e.Url.ToString()
            If urlStr.StartsWith("about:") Then Return

            ' 统一拦截，防止 WebBrowser 直接导航
            e.Cancel = True

            ' 优先处理 Outlook 协议，提取 entityID/storeID 并快速打开
            Dim scheme As String = e.Url.Scheme
            If Not String.IsNullOrEmpty(scheme) AndAlso (scheme.Equals("outlook", StringComparison.OrdinalIgnoreCase) _
                                                         OrElse scheme.Equals("ms-outlook", StringComparison.OrdinalIgnoreCase)) Then
                Dim entityId As String = Nothing
                Dim storeId As String = Nothing

                ' 解析查询参数（大小写不敏感）
                Dim qIndex As Integer = urlStr.IndexOf("?"c)
                If qIndex >= 0 AndAlso qIndex < urlStr.Length - 1 Then
                    Dim query As String = urlStr.Substring(qIndex + 1)
                    For Each kv In query.Split("&"c)
                        Dim parts = kv.Split("="c)
                        If parts.Length >= 2 Then
                            Dim key = parts(0)
                            Dim val = String.Join("=", parts.Skip(1)) ' 允许值中包含 '='
                            If key.Equals("entityid", StringComparison.OrdinalIgnoreCase) Then
                                entityId = Uri.UnescapeDataString(val)
                            ElseIf key.Equals("storeid", StringComparison.OrdinalIgnoreCase) Then
                                storeId = Uri.UnescapeDataString(val)
                            End If
                        End If
                    Next
                End If

                If Not String.IsNullOrEmpty(entityId) Then
                    If Not OutlookAddIn3.Utils.OutlookUtils.FastOpenMailItem(entityId, storeId) Then
                        ' 兜底：仍然交给系统处理
                        MailHandler.OpenLink(urlStr)
                    End If
                Else
                    ' 未能解析 entityID，回退到系统打开
                    MailHandler.OpenLink(urlStr)
                End If
            Else
                ' 普通 http/https 等链接，走系统默认浏览器
                MailHandler.OpenLink(urlStr)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"MailBrowser_Navigating error: {ex.Message}")
        End Try
    End Sub

    ' 添加类级别的防重复调用变量
    Private isUpdatingMailList As Boolean = False
    Private lastUpdateTime As DateTime = DateTime.MinValue
    Private Const UpdateThreshold As Integer = 500 ' 毫秒

    Public Async Sub UpdateMailList(conversationId As String, mailEntryID As String)
        Try
            ' 防重复调用检查
            If isUpdatingMailList Then
                Debug.WriteLine("UpdateMailList: 已有更新操作正在进行中，跳过")
                Return
            End If

            ' 时间间隔检查（避免短时间内重复调用）
            Dim now = DateTime.Now
            If (now - lastUpdateTime).TotalMilliseconds < UpdateThreshold AndAlso
               String.Equals(mailEntryID, currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                Debug.WriteLine($"UpdateMailList: 跳过重复更新，时间间隔: {(now - lastUpdateTime).TotalMilliseconds}ms")
                Return
            End If

            isUpdatingMailList = True
            lastUpdateTime = now

            ' 调试信息（仅在需要时启用）
            'Debug.WriteLine($"UpdateMailList 被调用，调用堆栈: {Environment.StackTrace}")

            If String.IsNullOrEmpty(mailEntryID) Then
                lvMails?.Items.Clear()
                Try
                    If suppressWebViewUpdate = 0 Then
                        mailBrowser.DocumentText = "<html><body style='font-family: Segoe UI; padding: 20px; color: #666;'><div>请选择一封邮件</div></body></html>"
                    End If
                Catch
                End Try
                Return
            End If

            ' 记录开始时间，用于性能分析
            Dim startTime = DateTime.Now
            Debug.WriteLine($"开始更新邮件列表: {startTime}")

            ' 列表将重建，清空EntryID比较缓存
            entryIdCompareCache.Clear()
            entryIdCacheExpireTime = DateTime.Now.AddMinutes(CacheExpireMinutes)

            ' 检查是否需要重新加载列表
            Dim needReload As Boolean = True
            If lvMails.Items.Count > 0 AndAlso Not String.IsNullOrEmpty(conversationId) AndAlso
           String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                needReload = False
            End If

            ' 单独处理无会话的邮件
            If Not String.IsNullOrEmpty(mailEntryID) AndAlso String.IsNullOrEmpty(conversationId) Then
                Debug.WriteLine($"处理无会话邮件，强制重新加载({mailEntryID})")

                ' 异步加载列表（将当前单封邮件加入列表）
                Await LoadConversationMailsAsync(mailEntryID)
                
                ' 加载完成后再设置currentMailEntryID
                currentMailEntryID = mailEntryID

                ' 自动加载 WebView 内容
                If Me.IsHandleCreated Then
                    Me.BeginInvoke(Sub() LoadMailContentDeferred(mailEntryID))
                End If

                Debug.WriteLine($"处理无会话邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Return
            End If

            If needReload Then
                ' 异步加载会话邮件，完全不阻塞主窗口
                Await LoadConversationMailsAsync(mailEntryID)
                currentMailEntryID = mailEntryID

                ' 更新当前会话ID并检查笔记
                If Not String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                    currentConversationId = conversationId
                    'Await CheckWolaiRecordAsync(currentConversationId)
                End If
            Else
                ' 只更新高亮
                UpdateHighlightByEntryID(currentMailEntryID, mailEntryID)
                currentMailEntryID = mailEntryID
            End If
            Debug.WriteLine($"完成更新邮件列表，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateMailList error: {ex.Message}")
        Finally
            isUpdatingMailList = False
        End Try
    End Sub

    Public Async Sub UpdateMailListOld(conversationId As String, mailEntryID As String)

        ' 添加堆栈跟踪日志，查看谁调用了这个方法
        Debug.WriteLine($"UpdateMailList 被调用，调用堆栈: {Environment.StackTrace}")
        Try
            If String.IsNullOrEmpty(mailEntryID) Then
                lvMails?.Items.Clear()
                Return
            End If

            ' 记录开始时间，用于性能分析
            Dim startTime = DateTime.Now
            Debug.WriteLine($"开始更新邮件列表: {startTime}")

            If mailEntryID = currentMailEntryID Then
                Debug.WriteLine($"跳过重复更新，时间间隔: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Return
            End If

            ' 检查是否需要重新加载列表
            Dim needReload As Boolean = True
            If lvMails.Items.Count > 0 AndAlso Not String.IsNullOrEmpty(conversationId) AndAlso
               String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                needReload = False
            End If

            ' 单独处理无会话的邮件
            If Not String.IsNullOrEmpty(mailEntryID) AndAlso String.IsNullOrEmpty(conversationId) Then
                currentMailEntryID = mailEntryID
                Debug.WriteLine($"处理无会话邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Return
            End If

            If needReload Then
                ' 暂时移除事件处理器，避免重复触发
                'RemoveHandler lvMails.SelectedIndexChanged, AddressOf lvMails_SelectedIndexChanged
                ' 使用异步方法加载会话邮件
                Await LoadConversationMailsAsync(mailEntryID)
                'LoadConversationMails(mailEntryID)
                ' 重新添加事件处理器
                'AddHandler lvMails.SelectedIndexChanged, AddressOf lvMails_SelectedIndexChanged
                ' 更新当前会话ID并检查笔记
                If Not String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                    currentConversationId = conversationId
                    'Await CheckWolaiRecordAsync(currentConversationId)
                End If


            Else
                ' 只更新高亮
                UpdateHighlightByEntryID(currentMailEntryID, mailEntryID)
            End If
            currentMailEntryID = mailEntryID
            Debug.WriteLine($"完成更新邮件列表，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateMailList error: {ex.Message}")
        End Try

    End Sub

    Private Function GetIndexByEntryID(entryID As String) As Integer
        Dim normalizedEntryID As String = entryID.Trim()
        Return mailItems.FindIndex(Function(x) String.Equals(x.EntryID, normalizedEntryID, StringComparison.OrdinalIgnoreCase))
    End Function


    ' 虚拟化ListView核心方法
    Private Sub EnableVirtualMode(totalItems As Integer)
        If Not _isPaginationEnabled Then
            isVirtualMode = False
            totalPages = 1
            currentPage = 0
            lvMails.VirtualMode = False
            Debug.WriteLine($"分页开关关闭：强制禁用虚拟模式，总项目={totalItems}")
            Return
        End If

        If totalItems > PageSize Then
            isVirtualMode = True
            totalPages = Math.Ceiling(totalItems / PageSize)
            currentPage = 0
            
            ' 启用ListView的虚拟模式
            lvMails.VirtualMode = True
            lvMails.VirtualListSize = totalItems
            
            Debug.WriteLine($"启用虚拟模式: 总项目={totalItems}, 总页数={totalPages}, 页大小={PageSize}")
        Else
            isVirtualMode = False
            totalPages = 1
            currentPage = 0
            lvMails.VirtualMode = False
            Debug.WriteLine($"禁用虚拟模式: 总项目={totalItems}")
        End If
    End Sub

    Private Sub LoadPage(pageIndex As Integer)
        If isLoadingPage OrElse pageIndex < 0 OrElse pageIndex >= totalPages Then
            Return
        End If

        isLoadingPage = True
        currentPage = pageIndex

        Try
            suppressWebViewUpdate += 1
            lvMails.BeginUpdate()
            lvMails.Items.Clear()
            mailItems.Clear()

            Dim startIndex As Integer = pageIndex * PageSize
            Dim endIndex As Integer = Math.Min(startIndex + PageSize - 1, allListViewItems.Count - 1)

            For i As Integer = startIndex To endIndex
                If i < allListViewItems.Count Then
                    ' 创建 ListViewItem 的副本以避免重复添加异常
                    Dim originalItem = allListViewItems(i)
                    Dim itemCopy As New ListViewItem(originalItem.Text)
                    itemCopy.Tag = originalItem.Tag

                    ' 复制除第一列外的所有子项
                    For si As Integer = 1 To originalItem.SubItems.Count - 1
                        itemCopy.SubItems.Add(originalItem.SubItems(si).Text)
                    Next

                    ' 复制其他属性（样式与图像）
                    itemCopy.BackColor = originalItem.BackColor
                    itemCopy.ForeColor = originalItem.ForeColor
                    itemCopy.Font = originalItem.Font
                    itemCopy.ImageKey = originalItem.ImageKey
                    itemCopy.ImageIndex = originalItem.ImageIndex
                    itemCopy.UseItemStyleForSubItems = originalItem.UseItemStyleForSubItems

                    lvMails.Items.Add(itemCopy)
                    If i < allMailItems.Count Then
                        mailItems.Add(allMailItems(i))
                    End If
                End If
            Next

            ' 分页完成后重设高亮并滚动到可见
            If Not String.IsNullOrEmpty(currentHighlightEntryID) Then
                UpdateHighlightByEntryID(String.Empty, currentHighlightEntryID)
            ElseIf Not String.IsNullOrEmpty(currentMailEntryID) Then
                UpdateHighlightByEntryID(String.Empty, currentMailEntryID)
            End If

            Debug.WriteLine($"加载第{pageIndex + 1}页: 显示项目{startIndex + 1}-{endIndex + 1}")
        Finally
            Try
                lvMails.EndUpdate()
            Finally
                suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
            End Try
            isLoadingPage = False
            UpdatePaginationUI()
        End Try
    End Sub

    Private Sub LoadNextPage()
        If isVirtualMode AndAlso currentPage < totalPages - 1 Then
            LoadPage(currentPage + 1)
        End If
    End Sub

    Private Sub LoadPreviousPage()
        If isVirtualMode AndAlso currentPage > 0 Then
            LoadPage(currentPage - 1)
        End If
    End Sub

    ' 异步版本的分页方法（优化：使用BeginInvoke避免阻塞UI）
    Private Async Function LoadPageAsync(pageIndex As Integer) As Task
        Try
            ShowProgress("正在加载页面...")
            Dim tcs As New TaskCompletionSource(Of Boolean)()
            Await Task.Run(Sub()
                               CancellationToken.ThrowIfCancellationRequested()
                               ' 使用BeginInvoke避免阻塞UI线程
                               Me.BeginInvoke(Sub()
                                                  Try
                                                      LoadPage(pageIndex)
                                                  Finally
                                                      tcs.SetResult(True)
                                                  End Try
                                              End Sub)
                           End Sub)
            Await tcs.Task
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("页面加载被取消")
        Finally
            HideProgress()
        End Try
    End Function

    Private Async Function LoadNextPageAsync() As Task
        Try
            ShowProgress("正在加载下一页...")
            Await Task.Run(Sub()
                               CancellationToken.ThrowIfCancellationRequested()
                               ' 使用BeginInvoke避免阻塞UI线程
                               Me.BeginInvoke(Sub() LoadNextPage())
                           End Sub)
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("下一页加载被取消")
        Finally
            HideProgress()
        End Try
    End Function

    Private Async Function LoadPreviousPageAsync() As Task
        Try
            ShowProgress("正在加载上一页...")
            Await Task.Run(Sub()
                               CancellationToken.ThrowIfCancellationRequested()
                               ' 使用BeginInvoke避免阻塞UI线程
                               Me.BeginInvoke(Sub() LoadPreviousPage())
                           End Sub)
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("上一页加载被取消")
        Finally
            HideProgress()
        End Try
    End Function

    ' 更新分页状态显示
    Private Sub UpdatePaginationUI()
        Try
            Dim paginationPanel As Panel = TryCast(splitter1?.Panel1?.Tag, Panel)
            If paginationPanel IsNot Nothing AndAlso paginationPanel.Tag IsNot Nothing Then
                Dim controls = paginationPanel.Tag

                ' 更新页面信息
                Dim lblPageInfo As Label = controls.PageInfo
                Dim lblItemCount As Label = controls.ItemCount
                Dim btnFirstPage As Button = controls.FirstPage
                Dim btnPrevPage As Button = controls.PrevPage
                Dim btnNextPage As Button = controls.NextPage
                Dim btnLastPage As Button = controls.LastPage

                If Not _isPaginationEnabled Then
                    lblPageInfo.Text = "第1页/共1页"
                    lblItemCount.Text = $"共{allListViewItems.Count}项"
                    ' 隐藏分页按钮但保持面板可见，以便显示CheckBox
                    btnFirstPage.Visible = False
                    btnPrevPage.Visible = False
                    lblPageInfo.Visible = False
                    btnNextPage.Visible = False
                    btnLastPage.Visible = False
                    paginationPanel.Visible = True
                ElseIf isVirtualMode Then
                    lblPageInfo.Text = $"第{currentPage + 1}页/共{totalPages}页"
                    lblItemCount.Text = $"共{allListViewItems.Count}项"

                    ' 显示所有分页控件
                    btnFirstPage.Visible = True
                    btnPrevPage.Visible = True
                    lblPageInfo.Visible = True
                    btnNextPage.Visible = True
                    btnLastPage.Visible = True

                    ' 更新按钮状态
                    btnFirstPage.Enabled = currentPage > 0
                    btnPrevPage.Enabled = currentPage > 0
                    btnNextPage.Enabled = currentPage < totalPages - 1
                    btnLastPage.Enabled = currentPage < totalPages - 1

                    paginationPanel.Visible = True
                Else
                    lblPageInfo.Text = "第1页/共1页"
                    lblItemCount.Text = $"共{allListViewItems.Count}项"
                    ' 根据邮件数量决定是否显示分页按钮
                    Dim shouldShowPagination = allListViewItems.Count > PageSize
                    btnFirstPage.Visible = shouldShowPagination
                    btnPrevPage.Visible = shouldShowPagination
                    lblPageInfo.Visible = shouldShowPagination
                    btnNextPage.Visible = shouldShowPagination
                    btnLastPage.Visible = shouldShowPagination
                    paginationPanel.Visible = True
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"UpdatePaginationUI error: {ex.Message}")
        End Try
    End Sub

    ' 批量获取邮件属性，减少COM调用频率
    Private Function GetMailItemPropertiesBatch(mailItems As List(Of Object)) As List(Of MailItemProperties)
        Dim properties As New List(Of MailItemProperties)(mailItems.Count)
        Dim sw As New Stopwatch()
        sw.Start()
        Dim cacheHits As Integer = 0
        Dim comCalls As Integer = 0

        ' COM对象不是线程安全的，使用优化的串行处理
        ' 预分配容量提高性能
        properties.Capacity = mailItems.Count

        ' 批量处理，减少异常处理开销
        For i As Integer = 0 To mailItems.Count - 1
            Dim mailItem = mailItems(i)
            Dim props As New MailItemProperties()
            Dim entryID As String = Nothing

            Try
                If mailItem IsNot Nothing Then
                    ' 先获取EntryID用于缓存查找
                    Select Case True
                        Case TypeOf mailItem Is Outlook.MailItem
                            entryID = DirectCast(mailItem, Outlook.MailItem).EntryID
                        Case TypeOf mailItem Is Outlook.AppointmentItem
                            entryID = DirectCast(mailItem, Outlook.AppointmentItem).EntryID
                        Case TypeOf mailItem Is Outlook.MeetingItem
                            entryID = DirectCast(mailItem, Outlook.MeetingItem).EntryID
                    End Select

                    ' 检查缓存
                    If Not String.IsNullOrEmpty(entryID) Then
                        SyncLock mailPropertiesCache
                            If mailPropertiesCache.ContainsKey(entryID) Then
                                Dim cacheEntry = mailPropertiesCache(entryID)
                                If (DateTime.Now - cacheEntry.CacheTime).TotalMinutes < MailPropertiesCacheExpiryMinutes Then
                                    ' 缓存命中
                                    props = cacheEntry.Properties
                                    cacheHits += 1
                                    properties.Add(props)
                                    Continue For
                                Else
                                    ' 缓存过期，移除
                                    mailPropertiesCache.Remove(entryID)
                                End If
                            End If
                        End SyncLock
                    End If

                    ' 缓存未命中，执行COM调用
                    comCalls += 1
                    Select Case True
                        Case TypeOf mailItem Is Outlook.MailItem
                            Dim mail As Outlook.MailItem = DirectCast(mailItem, Outlook.MailItem)
                            ' 一次性读取所有属性，减少COM调用
                            props.EntryID = mail.EntryID
                            props.ReceivedTime = mail.ReceivedTime
                            props.SenderName = mail.SenderName
                            props.Subject = mail.Subject
                            props.MessageClass = mail.MessageClass
                            props.CreationTime = mail.CreationTime
                            props.IsValid = True

                        Case TypeOf mailItem Is Outlook.AppointmentItem
                            Dim appt As Outlook.AppointmentItem = DirectCast(mailItem, Outlook.AppointmentItem)
                            props.EntryID = appt.EntryID
                            props.ReceivedTime = appt.Start
                            props.SenderName = appt.Organizer
                            props.Subject = appt.Subject
                            props.MessageClass = appt.MessageClass
                            props.CreationTime = appt.CreationTime
                            props.IsValid = True

                        Case TypeOf mailItem Is Outlook.MeetingItem
                            Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                            props.EntryID = meeting.EntryID
                            props.ReceivedTime = meeting.CreationTime
                            props.SenderName = meeting.SenderName
                            props.Subject = meeting.Subject
                            props.MessageClass = meeting.MessageClass
                            props.CreationTime = meeting.CreationTime
                            props.IsValid = True
                    End Select

                    ' 将结果存入缓存
                    If props.IsValid AndAlso Not String.IsNullOrEmpty(props.EntryID) Then
                        SyncLock mailPropertiesCache
                            ' 限制缓存大小，防止内存泄漏
                            If mailPropertiesCache.Count >= 500 Then
                                ' 清理过期缓存
                                Dim expiredKeys As New List(Of String)
                                For Each kvp In mailPropertiesCache
                                    If (DateTime.Now - kvp.Value.CacheTime).TotalMinutes >= MailPropertiesCacheExpiryMinutes Then
                                        expiredKeys.Add(kvp.Key)
                                    End If
                                Next
                                For Each key In expiredKeys
                                    mailPropertiesCache.Remove(key)
                                Next

                                ' 如果清理后仍然过多，移除最旧的条目
                                If mailPropertiesCache.Count >= 500 Then
                                    Dim oldestKey As String = Nothing
                                    Dim oldestTime As DateTime = DateTime.MaxValue
                                    For Each kvp In mailPropertiesCache
                                        If kvp.Value.CacheTime < oldestTime Then
                                            oldestTime = kvp.Value.CacheTime
                                            oldestKey = kvp.Key
                                        End If
                                    Next
                                    If oldestKey IsNot Nothing Then
                                        mailPropertiesCache.Remove(oldestKey)
                                    End If
                                End If
                            End If

                            mailPropertiesCache(props.EntryID) = (props, DateTime.Now)
                        End SyncLock
                    End If
                End If
            Catch ex As System.Runtime.InteropServices.COMException
                ' 简化异常处理，减少字符串操作
                props.IsValid = False
                props.EntryID = "无法访问"
                props.SenderName = "无法访问"
                props.Subject = "无法访问"
                props.ReceivedTime = DateTime.MinValue
            Catch ex As System.Exception
                props.IsValid = False
                props.EntryID = "无法访问"
                props.SenderName = "无法访问"
                props.Subject = "无法访问"
                props.ReceivedTime = DateTime.MinValue
            End Try

            properties.Add(props)
        Next

        ' 优化完成：移除了并行处理，使用线程安全的串行处理

        sw.Stop()
        Debug.WriteLine($"批量获取 {mailItems.Count} 封邮件属性耗时: {sw.ElapsedMilliseconds}ms, 缓存命中: {cacheHits}, COM调用: {comCalls}, 缓存命中率: {If(mailItems.Count > 0, Math.Round(cacheHits * 100.0 / mailItems.Count, 1), 0)}%")
        Return properties
    End Function

    ' 新的异步方法，完全在后台线程执行耗时操作
    Private Async Function LoadConversationMailsAsync(currentMailEntryID As String) As Task
        ' 使用长格式EntryID进行比较
        If String.IsNullOrEmpty(currentMailEntryID) Then
            Return
        End If
        
        ' 立即更新实例变量，避免过期检查失败
        Me.currentMailEntryID = currentMailEntryID

        Try
            ' 快速检查：如果是同一个会话且列表已加载，直接更新高亮即可
            Dim quickConversationId As String = String.Empty
            Try
                Dim quickItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If quickItem IsNot Nothing Then
                    If TypeOf quickItem Is Outlook.MailItem Then
                        quickConversationId = DirectCast(quickItem, Outlook.MailItem).ConversationID
                    ElseIf TypeOf quickItem Is Outlook.AppointmentItem Then
                        quickConversationId = DirectCast(quickItem, Outlook.AppointmentItem).ConversationID
                    End If
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"快速获取会话ID失败: {ex.Message}")
            End Try

            ' 如果会话ID相同且列表已有内容，只更新高亮，不重新构建列表
            If Not String.IsNullOrEmpty(quickConversationId) AndAlso
               String.Equals(quickConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) AndAlso
               lvMails.Items.Count > 0 Then
                Debug.WriteLine($"会话ID未变化({quickConversationId})，跳过列表重建，仅更新高亮")
                ' 更新类级别的currentMailEntryID，然后更新高亮
                Dim oldEntryID As String = Me.currentMailEntryID
                Me.currentMailEntryID = currentMailEntryID
                UpdateHighlightByEntryID(oldEntryID, currentMailEntryID)
                Return
            End If

            ' 显示进度指示器
            ShowProgress("正在加载会话邮件...")

            Dim startTime = DateTime.Now
            Debug.WriteLine($"开始异步加载会话邮件: {startTime}")

            ' 在UI线程中显示加载状态（使用BeginInvoke避免阻塞）
            If Me.InvokeRequired Then
                Me.BeginInvoke(Sub()
                                   lvMails.BeginUpdate()
                                   lvMails.Items.Clear()
                                   ' 可以添加一个"正在加载..."的提示项
                                   Dim loadingItem As New ListViewItem("正在加载会话邮件...")
                                   loadingItem.SubItems.Add("")
                                   loadingItem.SubItems.Add("")
                                   loadingItem.SubItems.Add("")
                                   lvMails.Items.Add(loadingItem)
                                   lvMails.EndUpdate()
                               End Sub)
            Else
                lvMails.BeginUpdate()
                lvMails.Items.Clear()
                Dim loadingItem As New ListViewItem("正在加载会话邮件...")
                loadingItem.SubItems.Add("")
                loadingItem.SubItems.Add("")
                loadingItem.SubItems.Add("")
                lvMails.Items.Add(loadingItem)
                lvMails.EndUpdate()
            End If

            ' 在后台线程中执行耗时的Outlook操作
            Await Task.Run(Sub()
                               ' 检查取消令牌
                               CancellationToken.ThrowIfCancellationRequested()
                               LoadConversationMailsBackground(currentMailEntryID, startTime)
                           End Sub)
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("会话邮件加载被取消")
        Finally
            ' 隐藏进度指示器
            HideProgress()
        End Try
    End Function

    ' 后台线程执行的邮件加载逻辑
    Private Sub LoadConversationMailsBackground(currentMailEntryID As String, startTime As DateTime)
        Dim currentItem As Object = Nothing
        Dim conversation As Outlook.Conversation = Nothing
        Dim table As Outlook.Table = Nothing
        Dim allItems As New List(Of ListViewItem)()
        Dim tempMailItems As New List(Of (Index As Integer, EntryID As String))()

        ' 首先检查缓存
        Dim conversationId As String = String.Empty
        Try
            currentItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem IsNot Nothing Then
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversationId = DirectCast(currentItem, Outlook.MailItem).ConversationID
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversationId = DirectCast(currentItem, Outlook.AppointmentItem).ConversationID
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"获取会话ID失败: {ex.Message}")
        End Try

        ' 如果会话ID相同，只需要更新高亮
        If Not String.IsNullOrEmpty(conversationId) AndAlso
           String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
            Debug.WriteLine($"会话ID未变化({conversationId})，只更新高亮")
            If Me.IsHandleCreated Then
                Me.BeginInvoke(Sub()
                                   Dim oldEntryID As String = Me.currentMailEntryID
                                   Me.currentMailEntryID = currentMailEntryID
                                   UpdateHighlightByEntryID(oldEntryID, currentMailEntryID)
                               End Sub)
            End If
            Return
        End If

        ' 无会话邮件强制重新加载，不进行EntryID比较
        Debug.WriteLine($"处理邮件: 会话ID={If(String.IsNullOrEmpty(conversationId), "无", conversationId)}, EntryID={currentMailEntryID}")

        ' 检查会话缓存
        If Not String.IsNullOrEmpty(conversationId) AndAlso conversationMailsCache.ContainsKey(conversationId) Then
            Dim cachedData = conversationMailsCache(conversationId)
            If (DateTime.Now - cachedData.CacheTime).TotalMinutes < ConversationCacheExpiryMinutes Then
                Debug.WriteLine($"使用缓存的会话邮件数据: {cachedData.ListViewItems.Count} 封邮件")

                ' 深度克隆缓存的 ListViewItem 对象，避免跨实例引用
                allItems = New List(Of ListViewItem)(cachedData.ListViewItems.Count)
                For Each originalItem As ListViewItem In cachedData.ListViewItems
                    Dim itemCopy As New ListViewItem(originalItem.Text)
                    itemCopy.Tag = originalItem.Tag
                    itemCopy.Name = originalItem.Name
                    For si As Integer = 1 To originalItem.SubItems.Count - 1
                        itemCopy.SubItems.Add(originalItem.SubItems(si).Text)
                    Next
                    itemCopy.BackColor = originalItem.BackColor
                    itemCopy.ForeColor = originalItem.ForeColor
                    itemCopy.Font = originalItem.Font
                    itemCopy.ImageKey = originalItem.ImageKey
                    itemCopy.ImageIndex = originalItem.ImageIndex
                    itemCopy.UseItemStyleForSubItems = originalItem.UseItemStyleForSubItems
                    allItems.Add(itemCopy)
                Next
                tempMailItems = New List(Of (Index As Integer, EntryID As String))(cachedData.MailItems)

                ' 直接跳到UI更新部分
                GoTo UpdateUI
            Else
                ' 缓存过期，移除
                conversationMailsCache.Remove(conversationId)
            End If
        End If

        Try
            Try
                currentItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem Is Nothing Then
                    Throw New System.Exception("无法获取邮件项")
                End If

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                End If

                If conversation Is Nothing Then
                    ' 处理没有会话的单个邮件 - 使用批量属性获取
                    Debug.WriteLine($"处理无会话邮件，类型: {currentItem.GetType().Name}")
                    Debug.WriteLine($"邮件MessageClass: {If(currentItem.MessageClass, "未知")}")
                    Debug.WriteLine($"邮件EntryID: {currentMailEntryID}")

                    ' 预分配单邮件容量
                    allItems = New List(Of ListViewItem)(1)
                    tempMailItems = New List(Of (Index As Integer, EntryID As String))(1)

                    Try
                        Dim singleItemList As New List(Of Object) From {currentItem}
                        Dim propertiesList As List(Of MailItemProperties) = GetMailItemPropertiesBatch(singleItemList)

                        If propertiesList Is Nothing OrElse propertiesList.Count = 0 Then
                            Debug.WriteLine("GetMailItemPropertiesBatch 返回空结果")
                            Throw New System.Exception("无法获取邮件属性")
                        End If

                        Dim props As MailItemProperties = propertiesList(0)
                        Debug.WriteLine($"邮件属性获取结果: IsValid={props.IsValid}, Subject={props.Subject}")

                        Dim entryId As String = GetPermanentEntryID(currentItem)
                        Debug.WriteLine($"EntryID: {If(String.IsNullOrEmpty(entryId), "空", "已获取")}")

                        Dim lvi As New ListViewItem(GetItemImageText(currentItem)) With {
                            .Tag = entryId,
                            .Name = "0"
                        }

                        With lvi.SubItems
                            If props.IsValid Then
                                .Add(props.ReceivedTime.ToString("yyyy-MM-dd HH:mm"))
                                .Add(props.SenderName)
                                .Add(props.Subject)
                            Else
                                .Add("无法访问")
                                .Add("无法访问")
                                .Add("无法访问")
                            End If
                        End With

                        allItems.Add(lvi)
                        tempMailItems.Add((0, entryId))

                        Debug.WriteLine($"处理单个邮件完成，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                        Debug.WriteLine($"创建的ListView项目: Text='{lvi.Text}', Tag='{lvi.Tag}', SubItems数量={lvi.SubItems.Count}")
                        For i As Integer = 0 To lvi.SubItems.Count - 1
                            Debug.WriteLine($"  SubItem[{i}]: '{lvi.SubItems(i).Text}'")
                        Next
                    Catch singleEx As System.Exception
                        Debug.WriteLine($"处理无会话邮件时出错: {singleEx.Message}")
                        ' 创建一个错误显示项
                        Dim errorItem As New ListViewItem($"❌ 加载失败") With {
                            .Tag = currentMailEntryID,
                            .Name = "0"
                        }
                        errorItem.SubItems.Add(DateTime.Now.ToString("yyyy-MM-dd HH:mm"))
                        errorItem.SubItems.Add("系统")
                        errorItem.SubItems.Add($"无法加载邮件: {singleEx.Message}")

                        allItems.Add(errorItem)
                        tempMailItems.Add((0, currentMailEntryID))
                    End Try
                Else
                    ' 首先检查会话中的邮件数量
                    Dim conversationItemCount As Integer = 0
                    Try
                        Dim tempTable As Outlook.Table = conversation.GetTable()
                        Try
                            ' 快速计算会话邮件数量
                            Do Until tempTable.EndOfTable
                                Dim row As Outlook.Row = tempTable.GetNextRow()
                                conversationItemCount += 1
                                If row IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                End If
                            Loop
                        Finally
                            If tempTable IsNot Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(tempTable)
                            End If
                        End Try
                    Catch ex As System.Exception
                        Debug.WriteLine($"计算会话邮件数量失败: {ex.Message}")
                        conversationItemCount = 1 ' 默认按单邮件处理
                    End Try

                    ' 预分配allItems和tempMailItems容量，减少动态扩容开销
                    allItems = New List(Of ListViewItem)(Math.Max(conversationItemCount, 10))
                    tempMailItems = New List(Of (Index As Integer, EntryID As String))(Math.Max(conversationItemCount, 10))
                    Debug.WriteLine($"预分配列表容量: {Math.Max(conversationItemCount, 10)}")

                    If conversationItemCount <= 1 Then
                        ' 会话中只有1封邮件，按单邮件处理，避免双路径
                        Debug.WriteLine($"会话邮件数量={conversationItemCount}，按单邮件处理")

                        Dim stepTimer As New Stopwatch()
                        stepTimer.Start()

                        ' 直接从currentItem获取属性，避免GetMailItemPropertiesBatch调用
                        Dim entryId As String = ""
                        Dim subject As String = "无主题"
                        Dim senderName As String = "未知发件人"
                        Dim receivedTime As DateTime = DateTime.MinValue
                        Dim messageClass As String = ""
                        
                        Try
                            ' 直接访问邮件属性，减少COM调用
                            entryId = GetPermanentEntryID(currentItem)
                            
                            ' 安全获取邮件属性
                            Try
                                subject = If(currentItem.Subject, "无主题")
                            Catch
                                subject = "无法访问"
                            End Try
                            
                            Try
                                If TypeOf currentItem Is Outlook.MailItem Then
                                    senderName = If(DirectCast(currentItem, Outlook.MailItem).SenderName, "未知发件人")
                                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                                    senderName = If(DirectCast(currentItem, Outlook.AppointmentItem).Organizer, "未知组织者")
                                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                                    senderName = If(DirectCast(currentItem, Outlook.MeetingItem).SenderName, "未知发件人")
                                Else
                                    senderName = "未知发件人"
                                End If
                            Catch
                                senderName = "无法访问"
                            End Try
                            
                            Try
                                If TypeOf currentItem Is Outlook.MailItem Then
                                    receivedTime = DirectCast(currentItem, Outlook.MailItem).ReceivedTime
                                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                                    receivedTime = DirectCast(currentItem, Outlook.AppointmentItem).Start
                                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                                    receivedTime = DirectCast(currentItem, Outlook.MeetingItem).ReceivedTime
                                Else
                                    receivedTime = DateTime.MinValue
                                End If
                            Catch
                                receivedTime = DateTime.MinValue
                            End Try
                            
                            Try
                                messageClass = If(currentItem.MessageClass, "")
                            Catch
                                messageClass = ""
                            End Try
                            
                        Catch ex As System.Exception
                            Debug.WriteLine($"获取邮件属性失败: {ex.Message}")
                        End Try
                        
                        Debug.WriteLine($"直接获取属性耗时: {stepTimer.ElapsedMilliseconds}ms")

                        stepTimer.Restart()
                        ' 组装图标：类型 + 附件 + 旗标
                        Dim icons As New List(Of String)
                        If Not String.IsNullOrEmpty(messageClass) Then
                            If messageClass.Contains("IPM.Appointment") OrElse messageClass.Contains("IPM.Schedule.Meeting") Then
                                icons.Add("📅")
                            ElseIf messageClass.Contains("IPM.Task") Then
                                icons.Add("📋")
                            ElseIf messageClass.Contains("IPM.Contact") Then
                                icons.Add("👤")
                            Else
                                icons.Add("📧")
                            End If
                        Else
                            icons.Add("📧")
                        End If
                        ' 附件
                        Try
                            If currentItem IsNot Nothing Then
                                Dim mailForAttach = TryCast(currentItem, Outlook.MailItem)
                                If mailForAttach IsNot Nothing AndAlso mailForAttach.Attachments IsNot Nothing AndAlso mailForAttach.Attachments.Count > 0 Then
                                    icons.Add("📎")
                                End If
                            End If
                        Catch
                        End Try
                        ' 旗标
                        Try
                            Dim status = CheckItemHasTask(currentItem)
                            If status = TaskStatus.InProgress Then
                                icons.Add("🚩")
                            ElseIf status = TaskStatus.Completed Then
                                icons.Add("⚑")
                            End If
                        Catch
                        End Try

                        Dim iconText As String = String.Join(" ", icons)
                        Debug.WriteLine($"获取图标文本耗时: {stepTimer.ElapsedMilliseconds}ms")

                        Dim lvi As New ListViewItem(iconText) With {
                            .Tag = entryId,
                            .Name = "0"
                        }

                        With lvi.SubItems
                            .Add(If(receivedTime <> DateTime.MinValue, receivedTime.ToString("yyyy-MM-dd HH:mm"), "无时间"))
                            .Add(senderName)
                            .Add(subject)
                        End With

                        allItems.Add(lvi)
                        tempMailItems.Add((0, entryId))

                        Debug.WriteLine($"处理会话单邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                    Else
                        ' 会话中有多封邮件，进行会话处理
                        Debug.WriteLine($"会话邮件数量={conversationItemCount}，进行会话批量处理")
                        ' 使用批量处理方式加载会话邮件
                        table = conversation.GetTable()
                        ' 优化：只添加需要的列，减少数据传输和内存占用
                        table.Columns.RemoveAll() ' 移除默认列
                        Try
                            ' 只添加必需的列，避免重复
                            ' 使用PR_ENTRYID获取长格式EntryID而不是默认的短格式
                            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102") ' PR_ENTRYID (长格式)
                            table.Columns.Add("Subject")
                            table.Columns.Add("SenderName")
                            table.Columns.Add("ReceivedTime")
                            table.Columns.Add("MessageClass")
                            table.Columns.Add("CreationTime")
                            ' 添加附件和旗标状态列以优化性能
                            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B") ' PR_HASATTACH
                            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x10900003") ' PR_FLAG_STATUS

                            ' 预分配容量，提高性能
                            Dim currentIndex As Integer = 0
                            Dim batchSize As Integer = 0

                            ' 直接使用Table数据创建ListView项目，避免重复COM调用
                            Do Until table.EndOfTable
                                Dim row As Outlook.Row = table.GetNextRow()
                                Try
                                    ' 直接从Table行数据获取属性，避免SafeGetItemFromID调用
                                    ' 从PR_ENTRYID列获取长格式EntryID
                                    Dim entryId As String = If(row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102") IsNot Nothing, ConvertEntryIDToString(row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")), "")
                                    Dim subject As String = If(row("Subject") IsNot Nothing, row("Subject").ToString(), "无主题")
                                    Dim senderName As String = If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "未知发件人")
                                    Dim messageClass As String = If(row("MessageClass") IsNot Nothing, row("MessageClass").ToString(), "")
                                    
                                    ' 安全获取时间属性
                                    Dim receivedTime As DateTime = DateTime.MinValue
                                    Try
                                        If row("ReceivedTime") IsNot Nothing Then
                                            receivedTime = Convert.ToDateTime(row("ReceivedTime"))
                                        End If
                                    Catch
                                        receivedTime = DateTime.MinValue
                                    End Try

                                    ' 直接基于MAPI行数据生成图标，避免COM调用以提升性能
                                    Dim hasAttach As Boolean = False
                                    Dim flagStatus As Integer = 0
                                    
                                    ' 获取附件状态
                                    Try
                                        If row("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B") IsNot Nothing Then
                                            hasAttach = Convert.ToBoolean(row("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"))
                                        End If
                                    Catch
                                        hasAttach = False
                                    End Try
                                    
                                    ' 获取旗标状态
                                    Try
                                        If row("http://schemas.microsoft.com/mapi/proptag/0x10900003") IsNot Nothing Then
                                            flagStatus = Convert.ToInt32(row("http://schemas.microsoft.com/mapi/proptag/0x10900003"))
                                        End If
                                    Catch
                                        flagStatus = 0
                                    End Try
                                    
                                    ' 使用快速图标生成函数
                                    Dim iconText As String = GetIconTextFast(messageClass, hasAttach, flagStatus)

                                    ' 创建 ListViewItem，直接使用Table数据
                                    Dim lvi As New ListViewItem(iconText) With {
                                        .Tag = entryId,
                                        .Name = currentIndex.ToString()
                                    }

                                    ' 直接使用Table数据添加列，无需额外COM调用
                                    With lvi.SubItems
                                        .Add(If(receivedTime <> DateTime.MinValue, receivedTime.ToString("yyyy-MM-dd HH:mm"), "无时间"))
                                        .Add(senderName)
                                        .Add(subject)
                                    End With

                                    ' 添加到临时列表
                                    allItems.Add(lvi)
                                    tempMailItems.Add((currentIndex, entryId))
                                    currentIndex += 1
                                    batchSize += 1
                                    
                                Finally
                                    If row IsNot Nothing Then
                                        Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                    End If
                                End Try
                            Loop

                            Debug.WriteLine($"优化后收集了 {batchSize} 封邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms (无额外COM调用)")
                        Finally
                            If table IsNot Nothing Then
                                Runtime.InteropServices.Marshal.ReleaseComObject(table)
                            End If
                        End Try
                    End If
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"处理邮件时出错: {ex.Message}")
                ' 在UI线程中显示错误信息（使用BeginInvoke避免阻塞）
                Me.BeginInvoke(Sub()
                                   lvMails.BeginUpdate()
                                   lvMails.Items.Clear()
                                   Dim errorItem As New ListViewItem($"加载失败: {ex.Message}")
                                   errorItem.SubItems.Add("")
                                   errorItem.SubItems.Add("")
                                   errorItem.SubItems.Add("")
                                   lvMails.Items.Add(errorItem)
                                   lvMails.EndUpdate()
                               End Sub)
            End Try
        Finally
            ' 释放 COM 对象
            If conversation IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(conversation)
            End If
            If currentItem IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(currentItem)
            End If
        End Try

UpdateUI:
        ' 优化缓存策略：只缓存合理大小的会话，减少内存占用
        If Not String.IsNullOrEmpty(conversationId) AndAlso allItems.Count > 0 AndAlso allItems.Count <= 50 Then
            Dim swCache As New Stopwatch()
            swCache.Start()

            ' 使用更高效的克隆方式，只复制必要的属性
            Dim cacheItems As New List(Of ListViewItem)(allItems.Count)
            For Each originalItem As ListViewItem In allItems
                Dim itemCopy As New ListViewItem(originalItem.Text)
                itemCopy.Tag = originalItem.Tag
                itemCopy.Name = originalItem.Name

                ' 批量添加子项，减少逐个添加的开销
                If originalItem.SubItems.Count > 1 Then
                    Dim subItemTexts(originalItem.SubItems.Count - 2) As String
                    For si As Integer = 1 To originalItem.SubItems.Count - 1
                        subItemTexts(si - 1) = originalItem.SubItems(si).Text
                    Next
                    itemCopy.SubItems.AddRange(subItemTexts)
                End If

                ' 只复制关键的显示属性
                itemCopy.BackColor = originalItem.BackColor
                itemCopy.ImageKey = originalItem.ImageKey
                cacheItems.Add(itemCopy)
            Next

            ' 检查缓存大小，实施LRU清理策略
            SyncLock conversationMailsCache
                If conversationMailsCache.Count >= 20 Then
                    ' 找到最旧的缓存项并移除
                    Dim oldestKey As String = Nothing
                    Dim oldestTime As DateTime = DateTime.MaxValue
                    For Each kvp In conversationMailsCache
                        If kvp.Value.CacheTime < oldestTime Then
                            oldestTime = kvp.Value.CacheTime
                            oldestKey = kvp.Key
                        End If
                    Next
                    If oldestKey IsNot Nothing Then
                        conversationMailsCache.Remove(oldestKey)
                        Debug.WriteLine($"缓存已满，移除最旧项: {oldestKey}")
                    End If
                End If

                conversationMailsCache(conversationId) = (New List(Of (Index As Integer, EntryID As String))(tempMailItems), cacheItems, DateTime.Now)
            End SyncLock

            swCache.Stop()
            Debug.WriteLine($"缓存会话邮件数据: {cacheItems.Count} 封邮件，耗时: {swCache.ElapsedMilliseconds}ms，当前缓存项: {conversationMailsCache.Count}")
        ElseIf allItems.Count > 50 Then
            Debug.WriteLine($"会话邮件数量过多({allItems.Count}封)，跳过缓存以节省内存")
        End If

        ' 在UI线程中更新界面（使用BeginInvoke避免阻塞）
        suppressWebViewUpdate += 1
        Me.BeginInvoke(Sub()
                           Try
                               ' 检查是否被取消或邮件ID已改变
                               If CancellationToken.IsCancellationRequested OrElse 
                                  Not String.Equals(currentMailEntryID, Me.currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                                   Debug.WriteLine($"后台任务已过期，跳过UI更新: 期望{currentMailEntryID}, 当前{Me.currentMailEntryID}")
                                   Return
                               End If

                               ' 对邮件按时间降序排序（最新邮件在前）
                               allItems.Sort(New ListViewItemComparer(1, SortOrder.Descending))

                               ' 存储完整数据到虚拟化变量
                               allMailItems = New List(Of (Index As Integer, EntryID As String))(tempMailItems)
                               allListViewItems = New List(Of ListViewItem)(allItems)

                               ' 启用虚拟模式检查
                               EnableVirtualMode(allItems.Count)

                               If isVirtualMode Then
                                   ' 虚拟模式：清空ListView，依赖RetrieveVirtualItem事件
                                   lvMails.BeginUpdate()
                                   lvMails.Items.Clear()
                                   mailItems.Clear()
                                   
                                   ' 设置虚拟列表大小，触发RetrieveVirtualItem事件
                                   lvMails.VirtualListSize = allItems.Count
                                   lvMails.EndUpdate()
                                   
                                   Debug.WriteLine($"虚拟模式启用: 总项目={allItems.Count}，依赖RetrieveVirtualItem事件显示")
                               Else
                                   ' 非虚拟模式：优化的快速加载
                                   lvMails.BeginUpdate()
                                   lvMails.Items.Clear()
                                   mailItems.Clear()

                                   If allItems.Count > 0 Then
                                       ' 优化：直接添加原始项目，避免深度克隆
                                       ' 对于少量邮件（通常是单邮件），克隆开销远大于收益
                                       If allItems.Count <= 5 Then
                                           ' 少量邮件：直接使用原始项目，避免克隆开销
                                           lvMails.Items.AddRange(allItems.ToArray())
                                       Else
                                           ' 多量邮件：使用轻量级克隆，只复制必要属性
                                           Dim clones(allItems.Count - 1) As ListViewItem
                                           For i As Integer = 0 To allItems.Count - 1
                                               Dim originalItem As ListViewItem = allItems(i)
                                               Dim itemCopy As New ListViewItem(originalItem.Text) With {
                                                   .Tag = originalItem.Tag,
                                                   .Name = originalItem.Name
                                               }
                                               ' 批量添加子项，减少逐个添加开销
                                               If originalItem.SubItems.Count > 1 Then
                                                   Dim subTexts(originalItem.SubItems.Count - 2) As String
                                                   For si As Integer = 1 To originalItem.SubItems.Count - 1
                                                       subTexts(si - 1) = originalItem.SubItems(si).Text
                                                   Next
                                                   itemCopy.SubItems.AddRange(subTexts)
                                               End If
                                               clones(i) = itemCopy
                                           Next
                                           lvMails.Items.AddRange(clones)
                                       End If
                                       mailItems = tempMailItems
                                   End If

                                   lvMails.EndUpdate()
                               End If

                               ' 设置排序
                               lvMails.Sorting = SortOrder.Descending
                               lvMails.ListViewItemSorter = New ListViewItemComparer(1, SortOrder.Descending)
                               lvMails.Sort()

                               ' 设置高亮并确保可见（使用参数中的currentMailEntryID，避免被其他操作覆盖）
                               Me.currentMailEntryID = currentMailEntryID
                               UpdateHighlightByEntryID(String.Empty, currentMailEntryID)

                               ' 更新分页UI
                               UpdatePaginationUI()

                               ' 隐藏进度指示器
                               HideProgress()

                               Debug.WriteLine($"完成异步加载会话邮件，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                           Finally
                               ' 确保EndUpdate被调用
                               If Not isVirtualMode Then
                                   Try
                                       lvMails.EndUpdate()
                                   Catch
                                       ' 忽略重复EndUpdate错误
                                   End Try
                               End If
                               suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)

                               ' 如果抑制已解除且有选中项，更新web内容
                               If suppressWebViewUpdate = 0 AndAlso lvMails.SelectedItems.Count > 0 Then
                                   Dim selectedItem = lvMails.SelectedItems(0)
                                   If selectedItem.Tag IsNot Nothing Then
                                       Dim entryID = ConvertEntryIDToString(selectedItem.Tag)
                                       LoadMailContentDeferred(entryID)
                                   End If
                               End If
                           End Try
                       End Sub)
    End Sub

    ' 保留原有的同步方法作为备用
    Private Sub LoadConversationMails(currentMailEntryID As String)
        If String.IsNullOrEmpty(currentMailEntryID) Then
            Return
        End If

        Dim startTime = DateTime.Now
        Debug.WriteLine($"开始加载会话邮件: {startTime}")

        lvMails.BeginUpdate()
        Dim currentItem As Object = Nothing
        Dim conversation As Outlook.Conversation = Nothing
        Dim table As Outlook.Table = Nothing

        Try
            lvMails.Items.Clear()
            mailItems.Clear()

            Try
                currentItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem Is Nothing Then
                    Throw New System.Exception("无法获取邮件项")
                End If

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                End If

                If conversation Is Nothing Then
                    ' 处理没有会话的单个邮件
                    Dim entryId As String = GetPermanentEntryID(currentItem)
                    Dim lvi As New ListViewItem(GetItemImageText(currentItem)) With {
                    .Tag = entryId,
                    .Name = "0"
                }

                    With lvi.SubItems
                        Try
                            If TypeOf currentItem Is Outlook.MailItem Then
                                Dim mail As Outlook.MailItem = DirectCast(currentItem, Outlook.MailItem)
                                .Add(mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm"))
                                .Add(mail.SenderName)
                                .Add(mail.Subject)
                            ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                                Dim appt As Outlook.AppointmentItem = DirectCast(currentItem, Outlook.AppointmentItem)
                                .Add(appt.Start.ToString("yyyy-MM-dd HH:mm"))
                                .Add(appt.Organizer)
                                .Add(appt.Subject)
                            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                                Dim meeting As Outlook.MeetingItem = DirectCast(currentItem, Outlook.MeetingItem)
                                .Add(meeting.ReceivedTime.ToString("yyyy-MM-dd HH:mm"))
                                .Add(meeting.SenderName)
                                .Add(meeting.Subject)
                            End If
                        Catch ex As System.Runtime.InteropServices.COMException
                            Debug.WriteLine($"COM异常访问项目属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                            .Add("无法访问")
                            .Add("无法访问")
                            .Add("无法访问")
                        Catch ex As System.Exception
                            Debug.WriteLine($"访问项目属性时发生异常: {ex.Message}")
                            .Add("无法访问")
                            .Add("无法访问")
                            .Add("无法访问")
                        End Try
                    End With

                    lvMails.Items.Add(lvi)
                    mailItems.Add((0, entryId))

                    Debug.WriteLine($"处理单个邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Else
                    ' 使用批量处理方式加载会话邮件
                    table = conversation.GetTable()
                    Try
                        ' 优化：只添加需要的列，减少数据传输
                        table.Columns.RemoveAll() ' 移除默认列
                        ' 使用PR_ENTRYID获取长格式EntryID
                        table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")
                        table.Columns.Add("SentOn")
                        table.Columns.Add("ReceivedTime")
                        table.Columns.Add("SenderName")
                        table.Columns.Add("Subject")
                        table.Columns.Add("MessageClass")
                        ' 添加附件和旗标列以支持快速图标生成
                        table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B") ' PR_HASATTACH
                        table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x10900003") ' PR_FLAG_STATUS

                        ' 预分配容量，提高性能
                        Dim allItems As New List(Of ListViewItem)(100)
                        Dim tempMailItems As New List(Of (Index As Integer, EntryID As String))(100)
                        Dim currentIndex As Integer = 0
                        Dim batchSize As Integer = 0

                        ' 一次性收集所有数据
                        Do Until table.EndOfTable
                            Dim row As Outlook.Row = table.GetNextRow()
                            Try
                                ' 直接使用Table提供的长格式EntryID，避免额外的COM调用
                                Dim entryId As String = ConvertEntryIDToString(row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"))
                                Dim messageClass As String = If(row("MessageClass") IsNot Nothing, row("MessageClass").ToString(), "")
                                
                                ' 直接基于MAPI行数据生成图标，避免COM调用以提升性能
                                Dim hasAttach As Boolean = False
                                Dim flagStatus As Integer = 0
                                
                                ' 获取附件状态
                                Try
                                    If row("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B") IsNot Nothing Then
                                        hasAttach = Convert.ToBoolean(row("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"))
                                    End If
                                Catch
                                    hasAttach = False
                                End Try
                                
                                ' 获取旗标状态
                                Try
                                    If row("http://schemas.microsoft.com/mapi/proptag/0x10900003") IsNot Nothing Then
                                        flagStatus = Convert.ToInt32(row("http://schemas.microsoft.com/mapi/proptag/0x10900003"))
                                    End If
                                Catch
                                    flagStatus = 0
                                End Try
                                
                                ' 使用快速图标生成函数
                                Dim iconText As String = GetIconTextFast(messageClass, hasAttach, flagStatus)
                                
                                ' 创建 ListViewItem，使用长格式EntryID
                                Dim lvi As New ListViewItem(iconText) With {
                                .Tag = entryId,
                                .Name = currentIndex.ToString()
                            }

                                ' 添加所有列，直接使用Table数据
                                With lvi.SubItems
                                    .Add(If(row("ReceivedTime") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row("ReceivedTime").ToString()),
                                    DateTime.Parse(row("ReceivedTime").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                    "Unknown Date"))
                                    .Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                                    .Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))
                                End With

                                ' 添加到临时列表
                                allItems.Add(lvi)
                                tempMailItems.Add((currentIndex, entryId))
                                currentIndex += 1
                                batchSize += 1
                            Finally
                                If row IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                End If
                            End Try
                        Loop

                        Debug.WriteLine($"收集了 {batchSize} 封邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")

                        ' 一次性添加所有项目
                        Try
                            suppressWebViewUpdate += 1
                            lvMails.Items.Clear()
                            mailItems.Clear()
                            Dim clones2 As New List(Of ListViewItem)(allItems.Count)
                            For Each originalItem As ListViewItem In allItems
                                Dim itemCopy As New ListViewItem(originalItem.Text)
                                itemCopy.Tag = originalItem.Tag
                                For si As Integer = 1 To originalItem.SubItems.Count - 1
                                    itemCopy.SubItems.Add(originalItem.SubItems(si).Text)
                                Next
                                itemCopy.BackColor = originalItem.BackColor
                                itemCopy.ForeColor = originalItem.ForeColor
                                itemCopy.Font = originalItem.Font
                                itemCopy.ImageKey = originalItem.ImageKey
                                itemCopy.ImageIndex = originalItem.ImageIndex
                                itemCopy.UseItemStyleForSubItems = originalItem.UseItemStyleForSubItems
                                clones2.Add(itemCopy)
                            Next
                            lvMails.Items.AddRange(clones2.ToArray())
                            mailItems = tempMailItems
                        Finally
                            suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
                        End Try

                        ' 设置排序
                        lvMails.Sorting = SortOrder.Descending
                        lvMails.ListViewItemSorter = New ListViewItemComparer(1, SortOrder.Descending)
                        lvMails.Sort()

                        ' 设置高亮并确保可见
                        UpdateHighlightByEntryID(String.Empty, currentMailEntryID)

                        Debug.WriteLine($"完成加载会话邮件，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                    Finally
                        If table IsNot Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(table)
                        End If
                    End Try
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"处理邮件时出错: {ex.Message}")
                ' 避免向用户显示不必要的错误消息
                ' MessageBox.Show($"处理邮件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        Finally
            lvMails.EndUpdate()

            ' 释放 COM 对象
            If conversation IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(conversation)
            End If
            If currentItem IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(currentItem)
            End If
        End Try
    End Sub

    ' 在listview_Mailist添加构造列表
    Private Sub LoadConversationMailsOld(currentMailEntryID As String)
        If String.IsNullOrEmpty(currentMailEntryID) Then
            Return
        End If

        lvMails.BeginUpdate()
        Dim currentItem As Object = Nothing
        Dim conversation As Outlook.Conversation = Nothing
        Dim table As Outlook.Table = Nothing
        Try
            lvMails.Items.Clear()
            mailItems.Clear()

            Try
                currentItem = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem Is Nothing Then
                    Throw New System.Exception("无法获取邮件项")
                End If

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                End If

                If conversation Is Nothing Then
                    'Throw New System.Exception("无法获取会话信息")
                    '# 不要优化这个分支. 没有会话类型的Item. 后续还需观察有哪些需要特殊处理. 
                Else

                    table = conversation.GetTable()
                    Try
                        ' 优化：只添加需要的列，减少数据传输
                        table.Columns.RemoveAll() ' 移除默认列
                        ' 使用PR_ENTRYID获取长格式EntryID
                        table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")
                        table.Columns.Add("SentOn")
                        table.Columns.Add("ReceivedTime")
                        table.Columns.Add("SenderName")
                        table.Columns.Add("Subject")
                        table.Columns.Add("MessageClass")

                        Dim allItems As New List(Of ListViewItem)
                        Dim tempMailItems As New List(Of (Index As Integer, EntryID As String))
                        Dim currentIndex As Integer = 0

                        ' 一次性收集所有数据
                        Do Until table.EndOfTable
                            Dim row As Outlook.Row = table.GetNextRow()
                            Try
                                ' 直接使用Table提供的长格式EntryID，避免额外的COM调用
                                Dim entryId As String = ConvertEntryIDToString(row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"))
                                Dim messageClass As String = If(row("MessageClass") IsNot Nothing, row("MessageClass").ToString(), "")
                                
                                ' 直接基于MAPI行数据生成图标，避免COM调用以提升性能
                                Dim hasAttach As Boolean = False
                                Dim flagStatus As Integer = 0
                                
                                ' 获取附件状态
                                Try
                                    If row("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B") IsNot Nothing Then
                                        hasAttach = Convert.ToBoolean(row("http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"))
                                    End If
                                Catch
                                    hasAttach = False
                                End Try
                                
                                ' 获取旗标状态
                                Try
                                    If row("http://schemas.microsoft.com/mapi/proptag/0x10900003") IsNot Nothing Then
                                        flagStatus = Convert.ToInt32(row("http://schemas.microsoft.com/mapi/proptag/0x10900003"))
                                    End If
                                Catch
                                    flagStatus = 0
                                End Try
                                
                                ' 使用快速图标生成函数
                                Dim iconText As String = GetIconTextFast(messageClass, hasAttach, flagStatus)

                                ' 创建 ListViewItem，使用长格式EntryID
                                Dim lvi As New ListViewItem(iconText) With {
                                .Tag = entryId,
                                .Name = currentIndex.ToString()
                            }

                                ' 添加所有列，直接使用Table数据
                                With lvi.SubItems
                                    .Add(If(row("ReceivedTime") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row("ReceivedTime").ToString()),
                                    DateTime.Parse(row("ReceivedTime").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                    "Unknown Date"))
                                    .Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                                    .Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))
                                End With

                                ' 添加到临时列表
                                allItems.Add(lvi)
                                tempMailItems.Add((currentIndex, entryId))
                                currentIndex += 1
                            Finally
                                If row IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                End If
                            End Try
                        Loop

                        ' 一次性添加所有项目
                        lvMails.Items.Clear()
                        mailItems.Clear()
                        Dim clones3 As New List(Of ListViewItem)(allItems.Count)
                        For Each originalItem As ListViewItem In allItems
                            Dim itemCopy As New ListViewItem(originalItem.Text)
                            itemCopy.Tag = originalItem.Tag
                            For si As Integer = 1 To originalItem.SubItems.Count - 1
                                itemCopy.SubItems.Add(originalItem.SubItems(si).Text)
                            Next
                            itemCopy.BackColor = originalItem.BackColor
                            itemCopy.ForeColor = originalItem.ForeColor
                            itemCopy.Font = originalItem.Font
                            itemCopy.ImageKey = originalItem.ImageKey
                            itemCopy.ImageIndex = originalItem.ImageIndex
                            itemCopy.UseItemStyleForSubItems = originalItem.UseItemStyleForSubItems
                            clones3.Add(itemCopy)
                        Next
                        lvMails.Items.AddRange(clones3.ToArray())
                        mailItems = tempMailItems

                        ' 设置排序
                        lvMails.Sorting = SortOrder.Descending
                        lvMails.ListViewItemSorter = New ListViewItemComparer(1, SortOrder.Descending)
                        lvMails.Sort()

                        ' 设置高亮并确保可见
                        UpdateHighlightByEntryID(String.Empty, currentMailEntryID)

                    Finally
                        If table IsNot Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(table)
                        End If
                    End Try
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"Failed to process mail item: {ex.Message}")
                MessageBox.Show($"处理邮件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try

        Catch ex As System.Exception
            Debug.WriteLine($"LoadConversationMails error: {ex.Message}")
            MessageBox.Show("加载邮件时出错，请尝试重启 Outlook。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            lvMails.EndUpdate()

            ' 按顺序释放 COM 对象
            ' 按顺序释放所有 COM 对象
            If table IsNot Nothing Then
                Try
                    Runtime.InteropServices.Marshal.ReleaseComObject(table)
                Catch ex As System.Exception
                    Debug.WriteLine($"释放 table 对象时出错: {ex.Message}")
                End Try
                table = Nothing
            End If
            If conversation IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(conversation)
            End If
            If currentItem IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(currentItem)
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Private Enum TaskStatus
        None = 0
        InProgress = 1
        Completed = 2
    End Enum


    Private Function CheckItemHasTask(item As Object) As TaskStatus
        Try
            If TypeOf item Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)

                Try
                    ' 只检查是否被标记为任务，移除耗时的UserProperties遍历
                    If mail.IsMarkedAsTask Then
                        ' 使用 FlagStatus 检查任务是否完成
                        If mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete Then
                            Return TaskStatus.Completed
                        Else
                            Return TaskStatus.InProgress
                        End If
                    End If
                Catch ex As System.Runtime.InteropServices.COMException
                    ' COM异常时直接返回None，避免日志输出影响性能
                    Return TaskStatus.None
                Catch ex As System.Exception
                    Return TaskStatus.None
                End Try
            End If

            Return TaskStatus.None
        Catch ex As System.Exception
            Return TaskStatus.None
        End Try
    End Function

    Public Sub New()
        ' 这个调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 之后添加任何初始化代码
        defaultFont = SystemFonts.DefaultFont
        iconFont = New Font("Segoe UI Emoji", 9, FontStyle.Regular)  ' 使用 Segoe UI Emoji 字体以获得更好的 emoji 显示效果
        'iconFont = New Font("Segoe UI Emoji", 12, FontStyle.Regular)
        'iconFont = New Font(defaultFont, FontStyle.Regular)
        normalFont = New Font(defaultFont, FontStyle.Regular)
        highlightFont = New Font(defaultFont, FontStyle.Bold)  ' 使用 defaultFont 作为基础字体

        ' 最后设置控件
        SetupControls()
    End Sub

    ''' <summary>
    ''' 将ListView项目的Tag转换为EntryID字符串
    ''' </summary>
    ''' <param name="tag">ListView项目的Tag对象</param>
    ''' <returns>EntryID字符串</returns>
    Private Function ConvertEntryIDToString(tag As Object) As String
        Try
            If tag Is Nothing Then
                Return String.Empty
            End If
            
            ' 如果Tag是字节数组（长格式EntryID的二进制数据）
            If TypeOf tag Is Byte() Then
                Dim bytes As Byte() = DirectCast(tag, Byte())
                ' 将字节数组转换为十六进制字符串
                Return BitConverter.ToString(bytes).Replace("-", "")
            End If
            
            ' 如果Tag是字符串，直接返回
            Return tag.ToString()
        Catch ex As System.Exception
            Debug.WriteLine($"ConvertEntryIDToString error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' 将字符串EntryID转换为十六进制格式以便与字节数组格式进行比较
    ''' </summary>
    ''' <param name="entryId">字符串格式的EntryID</param>
    ''' <returns>十六进制格式的EntryID字符串</returns>
    Private Function ConvertStringToHexFormat(entryId As String) As String
        Try
            If String.IsNullOrEmpty(entryId) Then
                Return String.Empty
            End If
            
            ' 如果已经是十六进制格式（只包含0-9和A-F），直接返回
            If System.Text.RegularExpressions.Regex.IsMatch(entryId, "^[0-9A-Fa-f]+$") Then
                Return entryId.ToUpper()
            End If
            
            ' 如果是Base64格式的EntryID，先转换为字节数组再转换为十六进制
            Try
                Dim bytes As Byte() = Convert.FromBase64String(entryId)
                Return BitConverter.ToString(bytes).Replace("-", "")
            Catch
                ' 如果不是Base64格式，尝试将字符串转换为字节数组
                Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(entryId)
                Return BitConverter.ToString(bytes).Replace("-", "")
            End Try
        Catch ex As System.Exception
            Debug.WriteLine($"ConvertStringToHexFormat error: {ex.Message}")
            Return entryId ' 转换失败时返回原始字符串
        End Try
    End Function

    Private Sub UpdateHighlightByEntryID(oldEntryID As String, newEntryID As String)
        If Me.InvokeRequired Then
            Me.Invoke(New Action(Of String, String)(AddressOf UpdateHighlightByEntryID), oldEntryID, newEntryID)
        Else
            Try
                lvMails.BeginUpdate()

                ' 优化：只处理需要变化的项目，避免遍历所有项目
                Dim oldItem As ListViewItem = Nothing
                Dim newItem As ListViewItem = Nothing

                ' 如果oldEntryID为空，需要清除所有高亮项目
                If String.IsNullOrEmpty(oldEntryID) Then
                    ' 清除所有选中和高亮项目
                    For Each item As ListViewItem In lvMails.Items
                        If item.Selected OrElse item.BackColor = highlightColor Then
                            SetItemHighlight(item, False)
                        End If
                    Next
                Else
                    ' 查找特定的旧项目进行清除
                For Each item As ListViewItem In lvMails.Items
                        If item.Tag IsNot Nothing Then
                            ' 取缓存的规范化ItemEntryID（避免重复Convert）
                            Dim rawTag = item.Tag
                            Dim cacheKey As String = If(TypeOf rawTag Is String, DirectCast(rawTag, String), ConvertEntryIDToString(rawTag))
                            Dim itemEntryID As String = String.Empty
                            If Not entryIdCompareCache.TryGetValue(cacheKey, itemEntryID) Then
                                itemEntryID = ConvertEntryIDToString(rawTag)
                                entryIdCompareCache(cacheKey) = itemEntryID
                            End If
                            ' 尝试使用CompareEntryIDs进行MAPI级别的比较，如果失败回退到字符串比较
                            Dim isMatchedOld As Boolean = False
                            Dim normalizedOldEntryID As String = ConvertStringToHexFormat(oldEntryID.Trim())
                            Try
                                isMatchedOld = Globals.ThisAddIn.Application.Session.CompareEntryIDs(itemEntryID, normalizedOldEntryID)
                            Catch ex As System.Exception
                                Debug.WriteLine($"UpdateHighlightByEntryID: CompareEntryIDs(Old)失败: {ex.Message}, 回退到字符串比较")
                                Dim shortOldEntryID As String = OutlookAddIn3.Utils.OutlookUtils.GetShortEntryID(normalizedOldEntryID)
                                isMatchedOld = String.Equals(itemEntryID, normalizedOldEntryID, StringComparison.OrdinalIgnoreCase) _
                                               OrElse String.Equals(itemEntryID, shortOldEntryID, StringComparison.OrdinalIgnoreCase)
                            End Try
                            If isMatchedOld Then
                                oldItem = item
                                Exit For
                            End If
                        End If
                    Next
                End If

                ' 查找需要设置高亮的新项目
                If Not String.IsNullOrEmpty(newEntryID) Then
                    Debug.WriteLine($"UpdateHighlightByEntryID: 查找EntryID={newEntryID.Trim()}")
                    Dim normalizedNewEntryID As String = ConvertStringToHexFormat(newEntryID.Trim())
                    Dim shortNewEntryID As String = OutlookAddIn3.Utils.OutlookUtils.GetShortEntryID(normalizedNewEntryID)
                    Debug.WriteLine($"UpdateHighlightByEntryID: 规范化后(长)={normalizedNewEntryID}, 转换短格式={shortNewEntryID}")
                    
                    For Each item As ListViewItem In lvMails.Items
                        If item.Tag IsNot Nothing Then
                            ' 取缓存的规范化ItemEntryID（避免重复Convert）
                            Dim rawTag = item.Tag
                            Dim cacheKey As String = If(TypeOf rawTag Is String, DirectCast(rawTag, String), ConvertEntryIDToString(rawTag))
                            Dim itemEntryID As String = String.Empty
                            If Not entryIdCompareCache.TryGetValue(cacheKey, itemEntryID) Then
                                itemEntryID = ConvertEntryIDToString(rawTag)
                                entryIdCompareCache(cacheKey) = itemEntryID
                            End If
                            Debug.WriteLine($"UpdateHighlightByEntryID: 比较项目EntryID={itemEntryID} (Tag类型: {item.Tag.GetType().Name}, 原始Tag长度: {If(TypeOf rawTag Is String, DirectCast(rawTag, String).Length, If(TypeOf rawTag Is Byte(), DirectCast(rawTag, Byte()).Length, 0))})")
                            ' 尝试使用CompareEntryIDs进行MAPI级别的比较，如果失败回退到字符串比较
                            Dim isMatched As Boolean = False
                            Try
                                ' 使用Outlook Session的CompareEntryIDs方法进行精确比较
                                isMatched = Globals.ThisAddIn.Application.Session.CompareEntryIDs(itemEntryID, normalizedNewEntryID)
                                Debug.WriteLine($"UpdateHighlightByEntryID: CompareEntryIDs成功，结果={isMatched}")
                            Catch ex As System.Exception
                                ' 如果MAPI比较失败，使用字符串比较作为回退
                                Debug.WriteLine($"UpdateHighlightByEntryID: CompareEntryIDs失败: {ex.Message}, 回退到字符串比较")
                                shortNewEntryID = OutlookAddIn3.Utils.OutlookUtils.GetShortEntryID(normalizedNewEntryID)
                                isMatched = String.Equals(itemEntryID, normalizedNewEntryID, StringComparison.OrdinalIgnoreCase) _
                                           OrElse String.Equals(itemEntryID, shortNewEntryID, StringComparison.OrdinalIgnoreCase)
                                Debug.WriteLine($"UpdateHighlightByEntryID: 字符串比较结果={isMatched} (长格式匹配={String.Equals(itemEntryID, normalizedNewEntryID, StringComparison.OrdinalIgnoreCase)}, 短格式匹配={String.Equals(itemEntryID, shortNewEntryID, StringComparison.OrdinalIgnoreCase)})")
                            End Try
                            
                            If isMatched Then
                                newItem = item
                                Debug.WriteLine($"UpdateHighlightByEntryID: 找到匹配项目")
                                Exit For
                            End If
                        End If
                    Next
                    If newItem Is Nothing Then
                        Debug.WriteLine($"UpdateHighlightByEntryID: 未找到匹配的EntryID={newEntryID.Trim()}")
                    End If
                End If

                ' 只更新需要变化的项目，避免对同一项目重复操作
                If oldItem IsNot Nothing AndAlso newItem IsNot oldItem Then
                    SetItemHighlight(oldItem, False)
                End If

                If newItem IsNot Nothing Then
                    SetItemHighlight(newItem, True)
                    newItem.EnsureVisible()
                    currentHighlightEntryID = newEntryID
                End If

            Finally
                Try
                    lvMails.EndUpdate()
                Catch
                    ' 忽略重复EndUpdate错误
                End Try
            End Try
        End If
    End Sub


    Private Sub SetItemHighlight(item As ListViewItem, isHighlighted As Boolean)
        If isHighlighted Then
            item.BackColor = highlightColor
            item.Font = highlightFont
            item.Selected = True
        Else
            item.BackColor = SystemColors.Window
            item.Font = normalFont
            item.Selected = False  ' 确保取消选中状态
        End If
    End Sub
    Private Function GetPermanentEntryID(item As Object) As String
        Try
            Dim longEntryID As String = String.Empty
            If TypeOf item Is Outlook.MailItem Then
                longEntryID = DirectCast(item, Outlook.MailItem).EntryID
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                longEntryID = DirectCast(item, Outlook.AppointmentItem).EntryID
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                longEntryID = DirectCast(item, Outlook.MeetingItem).EntryID
            End If
            
            ' 统一返回长格式EntryID
            If Not String.IsNullOrEmpty(longEntryID) Then
                Return longEntryID
            End If
            Return String.Empty
        Catch ex As System.Exception
            Debug.WriteLine($"GetPermanentEntryID error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 添加键盘事件处理，支持分页导航（优化：改为异步调用，避免UI线程卡顿）
    Private Async Sub lvMails_KeyDown(sender As Object, e As KeyEventArgs) Handles lvMails.KeyDown
        Try
            If isVirtualMode Then
                Select Case e.KeyCode
                    Case Keys.PageDown
                        If e.Control Then
                            Await LoadNextPageAsync()
                            e.Handled = True
                        End If
                    Case Keys.PageUp
                        If e.Control Then
                            Await LoadPreviousPageAsync()
                            e.Handled = True
                        End If
                    Case Keys.Home
                        If e.Control Then
                            Await LoadPageAsync(0)
                            e.Handled = True
                        End If
                    Case Keys.End
                        If e.Control Then
                            Await LoadPageAsync(totalPages - 1)
                            e.Handled = True
                        End If
                End Select
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_KeyDown error: {ex.Message}")
        End Try
    End Sub

    ' 添加鼠标滚轮事件处理，支持自动分页（优化：改为异步调用，避免UI线程卡顿）
    Private Async Sub lvMails_MouseWheel(sender As Object, e As MouseEventArgs) Handles lvMails.MouseWheel
        Try
            If isVirtualMode AndAlso Control.ModifierKeys = Keys.Control Then
                If e.Delta > 0 Then
                    Await LoadPreviousPageAsync()
                ElseIf e.Delta < 0 Then
                    Await LoadNextPageAsync()
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_MouseWheel error: {ex.Message}")
        End Try
    End Sub

    Private Sub lvMails_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then Return

            Dim mailId As String = ConvertEntryIDToString(lvMails.SelectedItems(0).Tag)
            If String.IsNullOrEmpty(mailId) Then Return

            ' 始终更新高亮，不受suppressWebViewUpdate影响
            If Not mailId.Equals(currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                Dim oldMailId As String = currentMailEntryID
                currentMailEntryID = mailId
                UpdateHighlightByEntryID(oldMailId, mailId)

                ' 只有在非抑制模式下才加载WebView内容
                If suppressWebViewUpdate = 0 Then
                    ' 使用 BeginInvoke 在事件回调结束后加载邮件内容
                    Me.BeginInvoke(New Action(Of String)(AddressOf LoadMailContentDeferred), mailId)
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_SelectedIndexChanged error: {ex.Message}")
        End Try
    End Sub

    ' 异步加载邮件内容的方法
    Private Async Sub LoadMailContentAsync(mailId As String)
        Try
            ' 显示进度指示器
            ShowProgress("正在加载邮件内容...")

            ' 获取HTML内容并显示在中间区域的WebBrowser中
            Dim html As String = Await Task.Run(Function()
                                                    ' 检查取消令牌
                                                    CancellationToken.ThrowIfCancellationRequested()
                                                    Return MailHandler.DisplayMailContent(mailId)
                                                End Function)

            ' 检查是否被取消
            If CancellationToken.IsCancellationRequested Then
                Return
            End If

            ' 抑制期间不更新 WebView
            If suppressWebViewUpdate > 0 Then
                Debug.WriteLine($"WebView更新被抑制，跳过 LoadMailContentAsync: {mailId}")
            ElseIf mailBrowser IsNot Nothing AndAlso mailBrowser.IsHandleCreated Then
                mailBrowser.DocumentText = html
            End If
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("邮件内容加载被取消")
        Catch ex As System.Exception
            Debug.WriteLine($"LoadMailContentAsync error: {ex.Message}")
        Finally
            ' 隐藏进度指示器
            HideProgress()
        End Try
    End Sub

    ' 延迟加载邮件内容的方法，避免在事件回调中直接访问 Outlook 对象导致 COMException
    Private Async Sub LoadMailContentDeferred(mailId As String)
        Try
            ' 抑制期间不进行 WebView 更新，避免联系人信息列表构造时触发刷新
            If suppressWebViewUpdate > 0 Then
                Debug.WriteLine($"WebView更新被抑制，延迟重试 LoadMailContentDeferred: {mailId}")
                Await Task.Delay(100)
                If suppressWebViewUpdate = 0 AndAlso Me.IsHandleCreated Then
                    Me.BeginInvoke(Sub() LoadMailContentDeferred(mailId))
                End If
                Return
            End If

            Dim html As String = Await Task.Run(Function() MailHandler.DisplayMailContent(mailId))
            If mailBrowser IsNot Nothing AndAlso mailBrowser.IsHandleCreated AndAlso suppressWebViewUpdate = 0 Then
                mailBrowser.DocumentText = html
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"LoadMailContentDeferred error: {ex.Message}")
        End Try
    End Sub

    Private Class ListViewItemComparer
        Implements System.Collections.IComparer
        Implements System.Collections.Generic.IComparer(Of ListViewItem)

        Private columnIndex As Integer
        Private sortOrder As SortOrder

        Public Sub New(column As Integer, order As SortOrder)
            columnIndex = column
            sortOrder = order
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements System.Collections.IComparer.Compare
            Return Compare(DirectCast(x, ListViewItem), DirectCast(y, ListViewItem))
        End Function

        Public Function Compare(x As ListViewItem, y As ListViewItem) As Integer Implements System.Collections.Generic.IComparer(Of ListViewItem).Compare
            Dim result As Integer
            If columnIndex = 1 Then ' 日期列
                Dim dateX As DateTime
                Dim dateY As DateTime
                If DateTime.TryParse(x.SubItems(columnIndex).Text, dateX) AndAlso
                   DateTime.TryParse(y.SubItems(columnIndex).Text, dateY) Then
                    result = DateTime.Compare(dateX, dateY)
                Else
                    result = String.Compare(x.SubItems(columnIndex).Text,
                                         y.SubItems(columnIndex).Text)
                End If
            Else
                result = String.Compare(x.SubItems(columnIndex).Text,
                                     y.SubItems(columnIndex).Text)
            End If

            Return If(sortOrder = SortOrder.Ascending, result, -result)
        End Function
    End Class




    ' 此方法已被替换为上面的lvMails_ColumnClick方法
    'Private Sub lvMails_ColumnClick(sender As Object, e As ColumnClickEventArgs)
    '    Try
    '        Dim lv As ListView = DirectCast(sender, ListView)
    '
    '        ' 切换排序方向
    '        lv.Sorting = If(lv.Sorting = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)

    '        ' 使用自定义排序器
    '        lv.ListViewItemSorter = New MailThreadPane.ListViewItemComparer(e.Column, lv.Sorting)
    '        lv.Sort()
    '
    '        ' 更新高亮
    '        If Not String.IsNullOrEmpty(currentMailEntryID) Then
    '            UpdateHighlightByEntryID(String.Empty, currentMailEntryID)
    '        End If
    '
    '    Catch ex As System.Exception
    '        Debug.WriteLine("lvMails_ColumnClick error: " & ex.Message)
    '    End Try
    'End Sub

    Private Sub lvMails_DoubleClick(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count > 0 Then
                Dim selectedItem As ListViewItem = lvMails.SelectedItems(0)
                Dim mailId As String = ConvertEntryIDToString(selectedItem.Tag)
                If Not String.IsNullOrEmpty(mailId) Then
                    ' 优先使用快速打开（可进一步传StoreID优化）
                    If Not OutlookAddIn3.Utils.OutlookUtils.FastOpenMailItem(mailId) Then
                        ' 兜底：GetItemFromID + Display
                        Dim mailItem As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(mailId)
                        If mailItem IsNot Nothing Then
                            Try
                                mailItem.Display()
                            Finally
                                OutlookAddIn3.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                            End Try
                        End If
                    End If
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine("lvMails_DoubleClick error: " & ex.Message)
        End Try
    End Sub

    Private Sub TaskList_DoubleClick(sender As Object, e As EventArgs)
        Try
            If taskList.SelectedItems.Count > 0 Then
                Dim selectedItem As ListViewItem = taskList.SelectedItems(0)
                If selectedItem.Tag IsNot Nothing Then
                    Dim entryId As String = ""

                    ' 检查 Tag 类型，获取相应的 EntryID
                    If TypeOf selectedItem.Tag Is OutlookAddIn3.Models.TaskInfo Then
                        Dim taskInfo As OutlookAddIn3.Models.TaskInfo = DirectCast(selectedItem.Tag, OutlookAddIn3.Models.TaskInfo)
                        ' 优先使用 TaskEntryID，如果为空则使用 MailEntryID
                        entryId = If(Not String.IsNullOrEmpty(taskInfo.TaskEntryID), taskInfo.TaskEntryID, taskInfo.MailEntryID)
                    Else
                        ' 兜底：将 Tag 作为 EntryID 字符串处理
                        entryId = ConvertEntryIDToString(selectedItem.Tag)
                    End If

                    If Not String.IsNullOrEmpty(entryId) Then
                        ' 优先使用快速打开（传入 StoreID 可进一步优化）
                        Dim storeId As String = Nothing
                        If TypeOf selectedItem.Tag Is OutlookAddIn3.Models.TaskInfo Then
                            storeId = DirectCast(selectedItem.Tag, OutlookAddIn3.Models.TaskInfo).StoreID
                        End If
                        If Not OutlookAddIn3.Utils.OutlookUtils.FastOpenMailItem(entryId, storeId) Then
                            ' 兜底：GetItemFromID + Display
                            Dim taskItem As Object = OutlookAddIn3.Utils.OutlookUtils.SafeGetItemFromID(entryId, storeId)
                            If taskItem IsNot Nothing Then
                                Try
                                    taskItem.Display()
                                Finally
                                    OutlookAddIn3.Utils.OutlookUtils.SafeReleaseComObject(taskItem)
                                End Try
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine("TaskList_DoubleClick error: " & ex.Message)
        End Try
    End Sub
    Private Async Sub BtnAddTask_Click(sender As Object, e As EventArgs)
        Try
            If String.IsNullOrEmpty(currentConversationId) Then
                MessageBox.Show("请先选择一封邮件")
                Return
            End If

            ' 在后台线程中创建任务，避免阻塞UI
            Await Task.Run(Sub()
                               OutlookAddIn3.Handlers.TaskHandler.CreateNewTask(currentConversationId, currentMailEntryID)
                           End Sub)
        Catch ex As System.Exception
            Debug.WriteLine("BtnAddTask_Click error: " & ex.Message)
            MessageBox.Show("创建任务时出错: " & ex.Message)
        End Try
    End Sub

    Private Sub lvMails_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvMails.ColumnClick
        Try
            ' 列排序逻辑
            Dim column As Integer = e.Column
            If column = currentSortColumn Then
                ' 如果点击的是当前排序列，则反转排序方向
                currentSortOrder = Not currentSortOrder
            Else
                ' 如果点击的是新列，则设置为升序
                currentSortColumn = column
                currentSortOrder = True
            End If

            ' 应用排序
            lvMails.ListViewItemSorter = New ListViewItemComparer(column, currentSortOrder)
        Catch ex As System.Exception
            Debug.WriteLine("lvMails_ColumnClick error: " & ex.Message)
        End Try
    End Sub

    Public Shared Function ReplaceTableTag(mailItemHTML As String) As String
        Dim oldTableTag As String
        Dim newTableTag As String

        ' 定义要查找和替换的字符串
        oldTableTag = "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" width=""100%"">"
        newTableTag = "<table class=""hidden-table"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""left"" width=""100%"">"

        ' 检查是否包含旧的表格标签
        If InStr(mailItemHTML, oldTableTag) > 0 Then
            ' 替换第一个匹配的表格标签
            Return Replace(mailItemHTML, oldTableTag, newTableTag, 1, 1)
            ' 输出或处理替换后的HTML
            'Debug.Print resultHTML
        Else
            ' 如果没有找到，输出原始HTML
            'Debug.Print "未找到匹配的表格标签，原始HTML保持不变。"
            'Debug.Print mailItemHTML
            Return mailItemHTML
        End If
    End Function
End Class
