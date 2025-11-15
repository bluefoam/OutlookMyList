Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports System.Threading.Tasks
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Linq
Imports OutlookMyList.Utils


<ComVisible(True)>
Public Class MailThreadPane
    Inherits UserControl





    ' 添加类级别的字体缓存
    Private ReadOnly iconFont As Font
    Private Shadows ReadOnly defaultFont As Font
    Private ReadOnly highlightFont As Font
    Private ReadOnly normalFont As Font
    ' 高亮颜色将在ApplyTheme中动态设置

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

    ' 全局主题变量 - 缓存当前主题状态，避免每次重新获取
    Public Shared globalThemeBackgroundColor As String = "#ffffff"
    Public Shared globalThemeForegroundColor As String = "#000000"
    Public Shared globalThemeAccentColor As String = "#0078d7"
    Public Shared globalThemeLastUpdate As DateTime = DateTime.MinValue

    ' 主题应用标志 - 用于处理异步控件创建的时序问题
    Private needsThemeApplication As Boolean = False

    ' 判断是否为黑色主题
    Private Function IsDarkTheme(backgroundColor As Color) As Boolean
        ' 计算亮度 - 如果背景色较暗，则认为是黑色主题
        Dim brightness As Double = (0.299 * backgroundColor.R + 0.587 * backgroundColor.G + 0.114 * backgroundColor.B) / 255
        Return brightness < 0.5
    End Function

    ' 高亮颜色 - 根据主题动态调整
    Private highlightColor As Color

    ' 抑制在列表构造/填充时触发 WebView 刷新或加载的标志
    Private suppressWebViewUpdate As Integer = 0 ' 使用计数器以支持嵌套调用

    ' 标志：当前是否正在显示邮件内容（而不是默认的"请选择一封邮件"内容）
    Private isDisplayingMailContent As Boolean = False

    ' 存储当前的会话分组数据，用于会话节点点击时获取最新邮件
    Private currentConversationGroups As Dictionary(Of String, List(Of (EntryID As String, Subject As String, Received As DateTime)))

    ' 暴露抑制状态以供外部检查
    Public ReadOnly Property IsWebViewUpdateSuppressed As Boolean
        Get
            Return suppressWebViewUpdate > 0
        End Get
    End Property

    ' 事件绑定状态
    Private eventsBound As Boolean = False

    ' 确保事件绑定完成的方法
    Public Sub EnsureEventsBound()
        If Not eventsBound Then
            BindEvents()
            eventsBound = True
            Debug.WriteLine("EnsureEventsBound: 事件绑定完成")
        End If
    End Sub

    ' 重置WebView更新抑制状态的公共方法
    Public Sub ResetWebViewUpdateSuppression()
        Dim oldValue As Integer = suppressWebViewUpdate
        suppressWebViewUpdate = 0
        Debug.WriteLine($"ResetWebViewUpdateSuppression: suppressWebViewUpdate从 {oldValue} 重置为 {suppressWebViewUpdate}")
    End Sub

    ' 分页功能开关的私有字段
    Private _isPaginationEnabled As Boolean = False

    ' 分页状态改变事件
    Public Event PaginationEnabledChanged(enabled As Boolean)

    ' 按钮显示状态跟踪变量
    Private button1Visible As Boolean = False
    Private button2Visible As Boolean = False
    Private button3Visible As Boolean = False

    ' 会话信息按钮引用，用于更新邮件数量显示
    Private conversationInfoButton As Button

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
                                AddItemWithTheme(lvMails, clonedItem)
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
            Debug.WriteLine($"=== ApplyTheme 开始 ===")
            Debug.WriteLine($"背景色: {backgroundColor} (R:{backgroundColor.R}, G:{backgroundColor.G}, B:{backgroundColor.B})")
            Debug.WriteLine($"前景色: {foregroundColor} (R:{foregroundColor.R}, G:{foregroundColor.G}, B:{foregroundColor.B})")

            ' 更新全局主题变量
            globalThemeBackgroundColor = $"#{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}"
            globalThemeForegroundColor = $"#{foregroundColor.R:X2}{foregroundColor.G:X2}{foregroundColor.B:X2}"
            globalThemeLastUpdate = DateTime.Now
            Debug.WriteLine($"全局主题变量已更新: 背景={globalThemeBackgroundColor}, 前景={globalThemeForegroundColor}")

            ' 保存当前主题颜色
            currentBackColor = backgroundColor
            currentForeColor = foregroundColor

            ' 动态计算高亮颜色 - 根据主题背景色调整
              If IsDarkTheme(backgroundColor) Then
                  ' 黑色主题：使用更亮更明显的强调色
                  highlightColor = Color.FromArgb(100, 149, 237) ' 矢车菊蓝，在暗色主题下更明显
              Else
                  ' 浅色主题：使用更鲜明的蓝色
                  highlightColor = Color.FromArgb(65, 105, 225) ' 皇家蓝，在浅色主题下更明显
              End If

            ' 应用到控件
            Me.BackColor = backgroundColor
            Debug.WriteLine($"主面板背景色已设置: {Me.BackColor}")

            ' 应用到ListView
            If lvMails IsNot Nothing Then
                lvMails.BackColor = backgroundColor
                lvMails.ForeColor = foregroundColor
                ' 强制刷新ListView以确保主题正确应用和鼠标状态正常
                lvMails.Refresh()
                Debug.WriteLine($"ListView主题已应用: 背景={lvMails.BackColor}, 前景={lvMails.ForeColor}")
            Else
                Debug.WriteLine("警告: lvMails 为 Nothing")
            End If

            ' 应用到任务列表 - 如果taskList还没有创建，设置标志以便后续应用
            If taskList IsNot Nothing Then
                taskList.BackColor = backgroundColor
                taskList.ForeColor = foregroundColor
                taskList.Refresh()
                Debug.WriteLine("taskList主题已应用")
            Else
                Debug.WriteLine("taskList尚未创建，主题将在创建后应用")
                needsThemeApplication = True
            End If

            ' 应用到邮件历史列表
            If mailHistoryList IsNot Nothing Then
                mailHistoryList.BackColor = backgroundColor
                mailHistoryList.ForeColor = foregroundColor
                mailHistoryList.Refresh()
            End If

            ' 应用到待办邮件列表
            If pendingMailList IsNot Nothing Then
                pendingMailList.BackColor = backgroundColor
                pendingMailList.ForeColor = foregroundColor
                pendingMailList.Refresh()
                Debug.WriteLine("pendingMailList主题已应用")
            Else
                Debug.WriteLine("pendingMailList尚未创建，主题将在创建后应用")
                needsThemeApplication = True
            End If

            ' 应用到来往邮件TreeView - 添加空值检查防止NullReferenceException
            If tabControl IsNot Nothing AndAlso tabControl.TabPages.ContainsKey("来往邮件") Then
                Dim mailInteractionTabPage As TabPage = tabControl.TabPages("来往邮件")
                If mailInteractionTabPage IsNot Nothing Then
                    Dim mailInteractionTreeView As TreeView = FindTreeViewInControl(mailInteractionTabPage)
                    If mailInteractionTreeView IsNot Nothing Then
                        mailInteractionTreeView.BackColor = backgroundColor
                        mailInteractionTreeView.ForeColor = foregroundColor
                        Debug.WriteLine("mailInteractionTreeView主题已应用")
                    Else
                        Debug.WriteLine("警告: 未在来往邮件页面中找到TreeView控件")
                    End If
                Else
                    Debug.WriteLine("警告: 来往邮件TabPage为Nothing")
                End If
            Else
                Debug.WriteLine("警告: tabControl为Nothing或不包含来往邮件页面，延迟主题应用")
                ' 设置标志，在SetupTabPages完成后重新应用主题
                needsThemeApplication = True
            End If

            ' 应用到联系人信息TreeView
            If contactInfoTree IsNot Nothing Then
                contactInfoTree.BackColor = backgroundColor
                contactInfoTree.ForeColor = foregroundColor
                Debug.WriteLine("contactInfoTree主题已应用")
            End If

            ' 应用到分隔控件 - 保持分割条的视觉界限效果
            If splitter1 IsNot Nothing Then
                ' 设置面板为主题颜色
                splitter1.Panel1.BackColor = backgroundColor
                splitter1.Panel2.BackColor = backgroundColor
                ' 保持分割条的深灰色以增强视觉界限（不随主题变化）
                ' splitter1.BackColor 保持为 Color.FromArgb(70, 70, 70)
            End If

            If splitter2 IsNot Nothing Then
                ' 设置面板为主题颜色
                splitter2.Panel1.BackColor = backgroundColor
                splitter2.Panel2.BackColor = backgroundColor
                ' 保持分割条的深灰色以增强视觉界限（不随主题变化）
                ' splitter2.BackColor 保持为 Color.FromArgb(70, 70, 70)
            End If

            ' 应用到WebBrowser
            If mailBrowser IsNot Nothing Then
                ' 设置WebBrowser控件本身的背景色
                Try
                    mailBrowser.BackColor = backgroundColor
                Catch ex As System.Exception
                    Debug.WriteLine($"设置WebBrowser背景色失败: {ex.Message}")
                End Try

                ' 更新WebBrowser的CSS样式变量
                UpdateWebBrowserTheme(backgroundColor, foregroundColor)
                Debug.WriteLine("WebBrowser主题已更新")
            Else
                Debug.WriteLine("警告: mailBrowser 为 Nothing")
            End If

            ' 应用到TabControl
            If tabControl IsNot Nothing Then
                tabControl.BackColor = backgroundColor
                tabControl.ForeColor = foregroundColor
                Debug.WriteLine($"TabControl主题已应用: 背景={tabControl.BackColor}, 前景={tabControl.ForeColor}")

                ' 应用到所有TabPage
                For Each tabPage As TabPage In tabControl.TabPages
                    tabPage.BackColor = backgroundColor
                    tabPage.ForeColor = foregroundColor
                    Debug.WriteLine($"TabPage '{tabPage.Text}' 主题已应用: 背景={tabPage.BackColor}")

                    ' 递归应用主题到TabPage中的所有控件
                    ApplyThemeToControlsRecursive(tabPage, backgroundColor, foregroundColor)
                Next

                ' 延迟重绘TabControl，避免与DrawItem冲突
                ' tabControl.Invalidate()
                ' tabControl.Refresh()
            Else
                Debug.WriteLine("警告: tabControl 为 Nothing")
            End If

            ' 应用到按钮面板
            If btnPanel IsNot Nothing Then
                btnPanel.BackColor = backgroundColor

                ' 应用到按钮面板中的所有控件
                For Each ctrl As Control In btnPanel.Controls
                    If TypeOf ctrl Is Button Then
                        ' 为按钮应用主题颜色
                        Dim btn As Button = DirectCast(ctrl, Button)
                        btn.BackColor = backgroundColor
                        btn.ForeColor = foregroundColor
                        btn.FlatStyle = FlatStyle.Flat
                        btn.FlatAppearance.BorderColor = foregroundColor
                        btn.FlatAppearance.BorderSize = 1
                    Else
                        ctrl.BackColor = backgroundColor
                        ctrl.ForeColor = foregroundColor
                    End If
                Next
            End If

            ' 应用到分页面板
            Dim paginationPanel As Panel = TryCast(splitter1?.Panel1?.Tag, Panel)
            If paginationPanel IsNot Nothing Then
                paginationPanel.BackColor = backgroundColor
                Debug.WriteLine($"分页面板主题已应用: 背景={paginationPanel.BackColor}")

                ' 应用到分页面板中的所有控件
                For Each ctrl As Control In paginationPanel.Controls
                    If TypeOf ctrl Is Button Then
                        ' 为分页按钮应用主题颜色
                        Dim btn As Button = DirectCast(ctrl, Button)
                        btn.BackColor = backgroundColor
                        btn.ForeColor = foregroundColor
                        btn.FlatStyle = FlatStyle.Flat
                        btn.FlatAppearance.BorderColor = foregroundColor
                        btn.FlatAppearance.BorderSize = 1
                        Debug.WriteLine($"分页按钮 '{btn.Text}' 主题已应用: 背景={btn.BackColor}, 前景={btn.ForeColor}")
                    ElseIf TypeOf ctrl Is Label Then
                        ctrl.ForeColor = foregroundColor
                        Debug.WriteLine($"分页标签 '{ctrl.Text}' 前景色已设置: {ctrl.ForeColor}")
                    Else
                        ctrl.BackColor = backgroundColor
                        ctrl.ForeColor = foregroundColor
                    End If
                Next
            Else
                Debug.WriteLine("警告: 分页面板未找到")
            End If

            ' 应用主题到所有现有的ListView项目
            ApplyThemeToAllListViewItems()

            ' 温和的重绘，避免与DrawItem冲突
            Me.BeginInvoke(Sub()
                               Try
                                   ' 延迟刷新ListView以避免绘制冲突
                                   If lvMails IsNot Nothing Then lvMails.Invalidate()
                                   If taskList IsNot Nothing Then taskList.Invalidate()
                                   If mailHistoryList IsNot Nothing Then mailHistoryList.Invalidate()
                                   If pendingMailList IsNot Nothing Then pendingMailList.Invalidate()
                                   If tabControl IsNot Nothing Then tabControl.Invalidate()
                               Catch ex As System.Exception
                                   Debug.WriteLine($"延迟刷新异常: {ex.Message}")
                               End Try
                           End Sub)
            Debug.WriteLine($"=== ApplyTheme 完成 ===")
            Debug.WriteLine($"最终主题: 背景色={backgroundColor}, 前景色={foregroundColor}")
            Debug.WriteLine($"主题类型: {(If(IsDarkTheme(backgroundColor), "深色主题", "浅色主题"))}")

            ' 添加主题切换提示
            If IsDarkTheme(backgroundColor) Then
                Debug.WriteLine("当前为黑色主题模式 - 所有控件已应用深色配色")
            Else
                Debug.WriteLine("当前为浅色主题模式 - 所有控件已应用标准配色")
            End If
        Catch ex As System.Exception
            Debug.WriteLine("ApplyTheme error: " & ex.Message)
        End Try
    End Sub

    ' 递归应用主题到控件及其子控件
    Private Sub ApplyThemeToControlsRecursive(parent As Control, backgroundColor As Color, foregroundColor As Color)
        Try
            For Each ctrl As Control In parent.Controls
                If TypeOf ctrl Is Button Then
                    ' 为按钮应用主题颜色
                    Dim btn As Button = DirectCast(ctrl, Button)
                    btn.BackColor = backgroundColor
                    btn.ForeColor = foregroundColor
                    btn.FlatStyle = FlatStyle.Flat
                    btn.FlatAppearance.BorderColor = foregroundColor
                    btn.FlatAppearance.BorderSize = 1
                    Debug.WriteLine($"按钮 '{btn.Text}' 主题已应用: 背景={btn.BackColor}, 前景={btn.ForeColor}")
                ElseIf TypeOf ctrl Is Panel Then
                    ' 为面板应用主题颜色
                    ctrl.BackColor = backgroundColor
                    ctrl.ForeColor = foregroundColor
                    ' 递归处理面板中的控件
                    ApplyThemeToControlsRecursive(ctrl, backgroundColor, foregroundColor)
                ElseIf TypeOf ctrl Is Label Then
                    ' 为标签应用前景色
                    ctrl.ForeColor = foregroundColor
                ElseIf Not (TypeOf ctrl Is ListView) Then
                    ' 为其他控件应用主题颜色（除了ListView，因为它们已经单独处理了）
                    Try
                        ctrl.BackColor = backgroundColor
                        ctrl.ForeColor = foregroundColor
                    Catch
                        ' 忽略某些控件可能不支持颜色设置的错误
                    End Try

                    ' 如果控件有子控件，递归处理
                    If ctrl.HasChildren Then
                        ApplyThemeToControlsRecursive(ctrl, backgroundColor, foregroundColor)
                    End If
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"ApplyThemeToControlsRecursive error: {ex.Message}")
        End Try
    End Sub

    ' 更新WebBrowser的主题
    Public Sub UpdateWebBrowserTheme(backgroundColor As Color, foregroundColor As Color)
        Try
            currentBackColor = backgroundColor
            currentForeColor = foregroundColor
            Dim backHexLog As String = $"#{currentBackColor.R:X2}{currentBackColor.G:X2}{currentBackColor.B:X2}"
            Dim foreHexLog As String = $"#{currentForeColor.R:X2}{currentForeColor.G:X2}{currentForeColor.B:X2}"
            Debug.WriteLine($"[ThemeDebug] UpdateWebBrowserTheme: back={backHexLog}, fore={foreHexLog}")
            Debug.WriteLine($"[ThemeDebug] UpdateWebBrowserTheme: readyState={If(mailBrowser Is Nothing, "<null>", mailBrowser.ReadyState.ToString())}, documentIsNull={If(mailBrowser Is Nothing OrElse mailBrowser.Document Is Nothing, True, False)}, isDisplayingMailContent={isDisplayingMailContent}")
            If mailBrowser IsNot Nothing AndAlso mailBrowser.IsHandleCreated Then
                ' 构建CSS样式
                Dim bgColorHex As String = $"#{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}"
                Dim fgColorHex As String = $"#{foregroundColor.R:X2}{foregroundColor.G:X2}{foregroundColor.B:X2}"

                ' 计算辅助颜色
                Dim accentColorHex As String = "#0078d7"
                Dim borderColorHex As String = $"#{Math.Min(255, backgroundColor.R + 40):X2}{Math.Min(255, backgroundColor.G + 40):X2}{Math.Min(255, backgroundColor.B + 40):X2}"

                ' 创建增强的主题样式，使用更强的CSS选择器和!important声明
                Dim themeStyle As String = $"
                <style type=""text/css"">
                    /* 强制覆盖所有元素的背景和文字颜色 */
                    *, *:before, *:after {{
                        background-color: {bgColorHex} !important;
                        color: {fgColorHex} !important;
                        border-color: {borderColorHex} !important;
                    }}
                    
                    html, body {{
                        background-color: {bgColorHex} !important;
                        color: {fgColorHex} !important;
                        font-family: 'Segoe UI', Tahoma, Arial, sans-serif !important;
                        margin: 0 !important;
                        padding: 10px !important;
                        line-height: 1.4 !important;
                        /* IE滚动条样式 */
                        scrollbar-base-color: {borderColorHex} !important;
                        scrollbar-face-color: {borderColorHex} !important;
                        scrollbar-track-color: {bgColorHex} !important;
                        scrollbar-arrow-color: {fgColorHex} !important;
                        scrollbar-shadow-color: {borderColorHex} !important;
                        scrollbar-darkshadow-color: {bgColorHex} !important;
                        scrollbar-highlight-color: {borderColorHex} !important;
                        scrollbar-3dlight-color: {bgColorHex} !important;
                    }}
                    
                    /* 滚动条容器样式 */
                    body {{
                        overflow-y: auto !important;
                        overflow-x: hidden !important;
                    }}
                    
                    /* 滚动条样式 - IE兼容性 */
                    html {{
                        scrollbar-base-color: {borderColorHex} !important;
                        scrollbar-face-color: {borderColorHex} !important;
                        scrollbar-track-color: {bgColorHex} !important;
                        scrollbar-arrow-color: {fgColorHex} !important;
                        scrollbar-shadow-color: {borderColorHex} !important;
                        scrollbar-darkshadow-color: {bgColorHex} !important;
                        scrollbar-highlight-color: {borderColorHex} !important;
                        scrollbar-3dlight-color: {bgColorHex} !important;
                    }}
                    
                    /* 滚动条样式 - Webkit兼容性（备用） */
                    ::-webkit-scrollbar {{
                        width: 8px !important;
                        height: 8px !important;
                    }}
                    
                    ::-webkit-scrollbar-track {{
                        background: {bgColorHex} !important;
                    }}
                    
                    ::-webkit-scrollbar-thumb {{
                        background: {borderColorHex} !important;
                        border-radius: 4px !important;
                    }}
                    
                    ::-webkit-scrollbar-thumb:hover {{
                        background: {accentColorHex} !important;
                    }}
                    
                    /* 标题样式 */
                    h1, h2, h3, h4, h5, h6 {{
                        color: {accentColorHex} !important;
                        background-color: transparent !important;
                        margin-top: 0 !important;
                        margin-bottom: 10px !important;
                    }}
                    
                    /* 强调文本 */
                    strong, b, em, i {{
                        color: {accentColorHex} !important;
                        background-color: transparent !important;
                    }}
                    
                    /* 段落和文本 */
                    p, div, span, td, th, li {{
                        background-color: transparent !important;
                        color: {fgColorHex} !important;
                    }}
                    
                    /* 表格样式 */
                    table {{
                        background-color: transparent !important;
                        border-color: {borderColorHex} !important;
                    }}
                    
                    /* 链接样式 */
                    a, a:visited, a:hover, a:active {{
                        color: {accentColorHex} !important;
                        background-color: transparent !important;
                    }}
                    
                    /* 移除Outlook特定的样式 */
                    .MsoNormal, .MsoPlainText {{
                        background-color: transparent !important;
                        color: {fgColorHex} !important;
                    }}
                    
                    /* 处理内联样式的覆盖 */
                    [style*=""background""] {{
                        background-color: {bgColorHex} !important;
                    }}
                    
                    [style*=""color""] {{
                        color: {fgColorHex} !important;
                    }}
                </style>"

                ' 如果当前显示的是默认内容，更新为带主题的默认内容
                If Not isDisplayingMailContent AndAlso Not String.IsNullOrEmpty(mailBrowser.DocumentText) AndAlso
                   mailBrowser.DocumentText.Contains("请选择一封邮件") Then
                    Debug.WriteLine("UpdateWebBrowserTheme: 更新默认内容主题")
                    mailBrowser.DocumentText = GetThemedDefaultContent()
                ElseIf isDisplayingMailContent AndAlso Not String.IsNullOrEmpty(mailBrowser.DocumentText) Then
                    Debug.WriteLine("UpdateWebBrowserTheme: 当前正在显示邮件内容，强制应用主题色")
                    ' 即使正在显示邮件内容，也要确保WebBrowser控件本身的主题色正确
                    Try
                        mailBrowser.BackColor = backgroundColor

                        ' 如果文档已加载，尝试注入主题样式
                        Dim doc = mailBrowser.Document
                        If doc IsNot Nothing Then
                            bgColorHex = $"#{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}"
                            fgColorHex = $"#{foregroundColor.R:X2}{foregroundColor.G:X2}{foregroundColor.B:X2}"

                            ' 使用JavaScript强制应用样式到整个文档
                            Dim jsScript As String = $"
                                (function() {{
                                    try {{
                                        document.body.style.setProperty('background-color', '#{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}', 'important');
                                        document.body.style.setProperty('color', '#{foregroundColor.R:X2}{foregroundColor.G:X2}{foregroundColor.B:X2}', 'important');
                                        document.documentElement.style.setProperty('background-color', '#{backgroundColor.R:X2}{backgroundColor.G:X2}{backgroundColor.B:X2}', 'important');
                                    }} catch(e) {{}}
                                }})();
                            "

                            Try
                                doc.InvokeScript("eval", New Object() {jsScript})
                                Debug.WriteLine($"UpdateWebBrowserTheme: 已强制应用主题色到当前文档")
                                Dim bgComputed As Object = doc.InvokeScript("eval", New Object() {"(function(){try{var el=document.body;if(!el) return 'no-body';var c=(window.getComputedStyle?window.getComputedStyle(el):el.currentStyle);var val=(c? (c.backgroundColor||'') : ''); if(!val||val==='transparent'){val=(el.style?el.style.backgroundColor:'')||val;} return val||'';}catch(e){return 'err:'+e.message;}})();"})
                                Dim textLen As Object = doc.InvokeScript("eval", New Object() {"(function(){try{var b=document.body;return b&&b.innerText?b.innerText.length:0;}catch(e){return -1;}})();"})
                                Debug.WriteLine($"[ThemeDebug] UpdateWebBrowserTheme: computedBodyBg={If(bgComputed, Nothing)}, innerTextLen={If(textLen, Nothing)}")
                            Catch scriptEx As Exception
                                Debug.WriteLine($"UpdateWebBrowserTheme: 应用脚本失败: {scriptEx.Message}")
                            End Try
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"UpdateWebBrowserTheme: 应用主题色时出错: {ex.Message}")
                    End Try
                    ' ThemeMonitor 定时器暂时移除
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateWebBrowserTheme error: {ex.Message}")
        End Try
    End Sub

    ' 获取当前主题颜色（从ThisAddIn获取）
    Private Function GetCurrentThemeColors() As (backgroundColor As Color, foregroundColor As Color)
        Try
            ' 直接从ThisAddIn获取主题颜色
            If Globals.ThisAddIn IsNot Nothing Then
                Return Globals.ThisAddIn.GetCurrentThemeColors()
            Else
                Debug.WriteLine("ThisAddIn为空，使用默认颜色")
                Return (SystemColors.Window, SystemColors.WindowText)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"获取主题颜色失败: {ex.Message}")
            ' 返回默认颜色
            Return (SystemColors.Window, SystemColors.WindowText)
        End Try
    End Function

    ' 生成带主题的默认WebBrowser内容
    Private Function GetThemedDefaultContent() As String
        ' 使用全局主题变量，避免每次重新获取导致的冲突
        Debug.WriteLine($"GetThemedDefaultContent: 使用全局主题变量 - 背景: {globalThemeBackgroundColor}, 前景: {globalThemeForegroundColor}")

        ' 重置邮件内容显示标志
        isDisplayingMailContent = False
        Debug.WriteLine($"GetThemedDefaultContent: 重置 isDisplayingMailContent = {isDisplayingMailContent}")

        Return $"<html><head><style>
            body {{
                background-color: {globalThemeBackgroundColor} !important;
                color: {globalThemeForegroundColor} !important;
                font-family: Segoe UI, Arial, sans-serif;
                margin: 0;
                padding: 20px;
                text-align: center;
            }}
        </style></head><body><div>请选择一封邮件</div></body></html>"
    End Function


    Private WithEvents lvMails As ListView
    Private WithEvents taskList As ListView
    Private WithEvents mailHistoryList As ListView
    Private WithEvents pendingMailList As ListView  ' 专门用于待办邮件标签页
    Private WithEvents contactInfoTree As TreeView
    Public WithEvents mailBrowser As WebBrowser
    Private splitter1, splitter2 As SplitContainer
    Private tabControl As TabControl
    Private btnPanel As Panel
    Private mailContextMenu As ContextMenuStrip
    Private treeContextMenu As ContextMenuStrip

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

    ' 全局缓存开关访问 - 安全读取
    Private Shared Function IsCacheEnabled() As Boolean
        Try
            Return Globals.ThisAddIn IsNot Nothing AndAlso Globals.ThisAddIn.CacheEnabled
        Catch
            Return True ' 若不可用，默认启用缓存以保持行为一致
        End Try
    End Function

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
            If Not IsCacheEnabled() Then Return
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
        If Not IsCacheEnabled() Then
            Return ("", "", "", "", False)
        End If
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
        If Not IsCacheEnabled() Then Return
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
                                            EnsureEventsBound()
                                        End Sub)
                     Else
                         ' 如果控件尚未完成初始化，等待控件句柄创建完成
                         AddHandler Me.HandleCreated, Sub(s, e)
                                                          Task.Run(Async Function()
                                                                       Await Task.Delay(50)
                                                                       Me.BeginInvoke(Sub()
                                                                                          SetupTabPages()
                                                                                          BindEvents()
                                                                                          EnsureEventsBound()
                                                                                      End Sub)
                                                                   End Function)
                                                      End Sub
                     End If
                 End Function)
    End Sub

    Private Sub InitializeSplitContainers()
        ' 创建第一个分隔控件 - 增强视觉界限
        splitter1 = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .Panel1MinSize = 100,
            .Panel2MinSize = 150,
            .SplitterWidth = 2,
            .BackColor = currentBackColor, 'Color.FromArgb(70, 70, 70),  ' 深灰色分割条增强视觉界限
            .BorderStyle = BorderStyle.FixedSingle    ' 添加边框增强界限
        }
        ' 明确设置面板颜色以避免继承分割条颜色
        splitter1.Panel1.BackColor = currentBackColor
        splitter1.Panel2.BackColor = currentBackColor

        ' 创建第二个分隔控件 - 增强视觉界限
        splitter2 = New SplitContainer With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .Panel1MinSize = 100,
            .Panel2MinSize = 30,  ' 减小最小尺寸让侧边栏更窄
            .SplitterWidth = 2,
            .BackColor = currentBackColor, 'Color.FromArgb(70, 70, 70),  ' 深灰色分割条增强视觉界限
            .BorderStyle = BorderStyle.FixedSingle    ' 添加边框增强界限
        }
        ' 明确设置面板颜色以避免继承分割条颜色
        splitter2.Panel1.BackColor = currentBackColor
        splitter2.Panel2.BackColor = currentBackColor

        ' 设置初始分割位置，让Panel2（邮件内容侧边栏）占更小的比例
        ' 设置Panel2占整个高度的30%，Panel1占70%
        splitter2.SplitterDistance = CInt(splitter2.Height * 0.7)

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

        ' 延迟设置默认的主题化内容，确保主题完全初始化
        Try
            ' 使用Timer延迟设置，确保ThisAddIn的主题已经完全初始化
            Dim themeTimer As New System.Windows.Forms.Timer()
            themeTimer.Interval = 100 ' 延迟100毫秒
            AddHandler themeTimer.Tick, Sub(sender, e)
                                            Try
                                                themeTimer.Stop()
                                                themeTimer.Dispose()
                                                mailBrowser.DocumentText = GetThemedDefaultContent()
                                                Debug.WriteLine("延迟设置mailBrowser默认主题内容成功")
                                            Catch ex2 As System.Exception
                                                Debug.WriteLine($"延迟设置mailBrowser默认内容失败: {ex2.Message}")
                                            End Try
                                        End Sub
            themeTimer.Start()
            Debug.WriteLine("启动延迟设置mailBrowser主题内容的定时器")
        Catch ex As System.Exception
            Debug.WriteLine($"设置mailBrowser默认内容定时器失败: {ex.Message}")
        End Try

        ' 然后添加第一个分隔控件到窗体
        Me.Controls.Add(splitter1)

        ' 添加尺寸改变事件处理
        AddHandler Me.SizeChanged, AddressOf Control_Resize
        AddHandler splitter1.Panel2.SizeChanged, AddressOf Panel2_SizeChanged
    End Sub

    Private Sub SetupProgressIndicator()
        ' 创建进度标签
        progressLabel = New Label With {
            .Dock = DockStyle.Bottom,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Visible = False,
            .Height = 20 ' 设置一个与分页栏相似的高度
        }

        ' 将进度标签添加到主控件
        Me.Controls.Add(progressLabel)
        progressLabel.BringToFront()
    End Sub

    ' 显示进度指示器
    Public Sub ShowProgress(message As String, Optional isIndeterminate As Boolean = True)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() ShowProgress(message, isIndeterminate))
            Return
        End If

        Try
            If progressLabel IsNot Nothing Then
                progressLabel.Text = message
                progressLabel.Visible = True
                progressLabel.BackColor = currentBackColor
                progressLabel.ForeColor = currentForeColor
                progressLabel.Refresh()

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
            If progressLabel IsNot Nothing Then
                progressLabel.Visible = False
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
            If ErrorNotificationSettings.Instance.ShowErrorDialogs Then
                MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
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
            If TypeOf item Is Microsoft.Office.Interop.Outlook.MailItem Then
                icons.Add("✉️") '📧

                ' 检查附件
                Dim mail As Microsoft.Office.Interop.Outlook.MailItem = DirectCast(item, Microsoft.Office.Interop.Outlook.MailItem)
                Try
                    If mail.Attachments IsNot Nothing AndAlso mail.Attachments.Count > 0 Then
                        icons.Add("📎") ' 回形针图标表示有附件
                    End If
                Catch ex As System.Exception
                    ' 忽略附件检查错误
                End Try

            ElseIf TypeOf item Is Microsoft.Office.Interop.Outlook.AppointmentItem Then
                icons.Add("📅")
            ElseIf TypeOf item Is Microsoft.Office.Interop.Outlook.MeetingItem Then
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
            .OwnerDraw = True,  ' 使用自定义绘制以支持主题颜色
            .SmallImageList = New ImageList() With {.ImageSize = New Size(16, 15)}, ' 设置行高
            .VirtualMode = False  ' 初始禁用虚拟模式，根据需要动态启用
        }
        ' 启用拖拽合并
        lvMails.AllowDrop = True
        ' 不在这里设置颜色，等待ApplyTheme方法调用

        ' 创建右键菜单
        SetupContextMenu()

        ' 启用双缓冲以减少闪烁
        Dim listViewType As Type = lvMails.GetType()
        Dim doubleBufferedProperty As Reflection.PropertyInfo = listViewType.GetProperty("DoubleBuffered", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
        If doubleBufferedProperty IsNot Nothing Then
            doubleBufferedProperty.SetValue(lvMails, True, Nothing)
        End If

        With lvMails.Columns.Add("----", 50)  ' 增加宽度以适应更大的图标
            '.BackColor = Color.Transparent
        End With

        With lvMails.Columns.Add("日期", 120) ' 宽度适配“yyyy-MM-dd HH:mm”
            '.BackColor = Color.Transparent
        End With

        With lvMails.Columns.Add("发件人", 100)
            .TextAlign = HorizontalAlignment.Left
            '.BackColor = Color.Transparent
        End With
        With lvMails.Columns.Add("主题", 300)
            .TextAlign = HorizontalAlignment.Left
            '.BackColor = Color.Transparent
        End With

        ' 设置文本省略模式
        'For Each column As ColumnHeader In lvMails.Columns
        '    column.Width = -2  ' 自动调整列宽以适应内容
        'Next

        ' 创建分页导航面板（压缩高度）
        Dim paginationPanel As New Panel With {
            .Height = 20,
            .Dock = DockStyle.Bottom,
            .Padding = New Padding(0, 0, 0, 0)
        }
        ' 不在这里设置颜色，等待ApplyTheme方法调用

        ' 创建分页导航控件
        Dim btnFirstPage As New Button With {
            .Text = "首页",
            .Size = New Size(45, 18),
            .Location = New Point(5, 1)
        }

        Dim btnPrevPage As New Button With {
            .Text = "上页",
            .Size = New Size(45, 18),
            .Location = New Point(55, 1)
        }

        Dim lblPageInfo As New Label With {
            .Text = "第1页/共1页",
            .Size = New Size(90, 18),
            .Location = New Point(105, 1),
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = Color.Transparent
        }

        Dim btnNextPage As New Button With {
            .Text = "下页",
            .Size = New Size(45, 18),
            .Location = New Point(200, 1)
        }

        Dim btnLastPage As New Button With {
            .Text = "末页",
            .Size = New Size(45, 18),
            .Location = New Point(250, 1)
        }

        Dim lblItemCount As New Label With {
            .Text = "共0项",
            .Size = New Size(70, 18),
            .Location = New Point(300, 1),
            .TextAlign = ContentAlignment.MiddleLeft,
            .BackColor = Color.Transparent
        }

        ' 添加分页开关控件
        Dim chkPagination As New CheckBox With {
            .Text = "分页",
            .Size = New Size(50, 18),
            .Location = New Point(375, 1),
            .Checked = _isPaginationEnabled,
            .BackColor = Color.Transparent,
            .Visible = False
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

        ' 添加绘制事件处理 - 启用自定义绘制以支持主题颜色
        AddHandler lvMails.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler lvMails.DrawSubItem, AddressOf ListView_DrawSubItem

        ' 添加虚拟模式事件处理
        AddHandler lvMails.RetrieveVirtualItem, AddressOf ListView_RetrieveVirtualItem
        AddHandler lvMails.CacheVirtualItems, AddressOf ListView_CacheVirtualItems

        ' 添加拖拽事件处理
        AddHandler lvMails.ItemDrag, AddressOf lvMails_ItemDrag
        AddHandler lvMails.DragEnter, AddressOf lvMails_DragEnter
        AddHandler lvMails.DragDrop, AddressOf lvMails_DragDrop
    End Sub

    Private Sub SetupContextMenu()
        ' 创建右键菜单
        mailContextMenu = New ContextMenuStrip()

        ' 添加菜单项：显示会话ID
        Dim showConversationIdItem As New ToolStripMenuItem("显示会话ID")
        AddHandler showConversationIdItem.Click, AddressOf ShowConversationId_Click
        mailContextMenu.Items.Add(showConversationIdItem)

        ' 添加菜单项：复制会话ID
        Dim copyConversationIdItem As New ToolStripMenuItem("复制会话ID")
        AddHandler copyConversationIdItem.Click, AddressOf CopyConversationId_Click
        mailContextMenu.Items.Add(copyConversationIdItem)

        ' 添加菜单项：显示任务关联状态
        Dim showTaskStatusItem As New ToolStripMenuItem("显示任务关联状态")
        AddHandler showTaskStatusItem.Click, AddressOf ShowTaskStatus_Click
        mailContextMenu.Items.Add(showTaskStatusItem)

        ' 添加菜单项：自定义会话ID
        Dim customConversationIdItem As New ToolStripMenuItem("自定义会话ID")
        AddHandler customConversationIdItem.Click, AddressOf CustomConversationId_Click
        mailContextMenu.Items.Add(customConversationIdItem)



        ' 添加分隔线
        mailContextMenu.Items.Add(New ToolStripSeparator())

        ' 添加菜单项：跳转到联系人信息（动态显示发件人姓名）
        Dim gotoMailHistoryItem As New ToolStripMenuItem("最近往来邮件")
        AddHandler gotoMailHistoryItem.Click, AddressOf GotoMailHistory_Click
        mailContextMenu.Items.Add(gotoMailHistoryItem)

        ' 添加菜单项：待办邮件（动态显示发件人姓名）
        Dim pendingMailsItem As New ToolStripMenuItem("待办邮件")
        AddHandler pendingMailsItem.Click, AddressOf PendingMails_Click
        mailContextMenu.Items.Add(pendingMailsItem)

        mailContextMenu.Items.Add(New ToolStripSeparator())

        Dim customHistoryItem As New ToolStripMenuItem("自定义来往邮件")
        AddHandler customHistoryItem.Click, AddressOf CustomMailHistory_Click
        mailContextMenu.Items.Add(customHistoryItem)

        Dim customPendingItem As New ToolStripMenuItem("自定义待办邮件")
        AddHandler customPendingItem.Click, AddressOf CustomPendingMails_Click
        mailContextMenu.Items.Add(customPendingItem)

        AddHandler mailContextMenu.Opening, AddressOf MailContextMenu_Opening
        mailContextMenu.AutoClose = True
        AddHandler Me.MouseDown, AddressOf HideContextMenuOnMouseDown
        AddHandler lvMails.MouseDown, AddressOf HideContextMenuOnMouseDown

        ' 将右键菜单绑定到ListView
        lvMails.ContextMenuStrip = mailContextMenu
    End Sub

    Private Sub SetupTreeContextMenu()
        ' 创建TreeView右键菜单
        treeContextMenu = New ContextMenuStrip()

        ' 添加菜单项：显示邮件ID
        Dim showMailIdItem As New ToolStripMenuItem("显示邮件ID")
        AddHandler showMailIdItem.Click, AddressOf ShowMailId_Click
        treeContextMenu.Items.Add(showMailIdItem)

        ' 添加菜单项：显示会话ID
        Dim showConversationIdItem As New ToolStripMenuItem("显示会话ID")
        AddHandler showConversationIdItem.Click, AddressOf ShowTreeConversationId_Click
        treeContextMenu.Items.Add(showConversationIdItem)

        ' 添加菜单项：显示智能会话ID
        Dim showSmartConversationItem As New ToolStripMenuItem("显示智能会话ID")
        AddHandler showSmartConversationItem.Click, AddressOf ShowSmartConversationId_Click
        treeContextMenu.Items.Add(showSmartConversationItem)

        ' 添加菜单项：显示任务关联状态
        Dim showTaskStatusItem As New ToolStripMenuItem("显示任务关联状态")
        AddHandler showTaskStatusItem.Click, AddressOf ShowTreeTaskStatus_Click
        treeContextMenu.Items.Add(showTaskStatusItem)

        ' 添加分隔线
        treeContextMenu.Items.Add(New ToolStripSeparator())

        ' 添加菜单项：复制邮件ID
        Dim copyMailIdItem As New ToolStripMenuItem("复制邮件ID")
        AddHandler copyMailIdItem.Click, AddressOf CopyMailId_Click
        treeContextMenu.Items.Add(copyMailIdItem)

        ' 添加菜单项：在Outlook中打开
        Dim openInOutlookItem As New ToolStripMenuItem("在Outlook中打开")
        AddHandler openInOutlookItem.Click, AddressOf OpenInOutlook_Click
        treeContextMenu.Items.Add(openInOutlookItem)

        ' 添加分隔线
        treeContextMenu.Items.Add(New ToolStripSeparator())

        ' 添加菜单项：标记为相关会话
        Dim markRelatedItem As New ToolStripMenuItem("标记为相关会话")
        AddHandler markRelatedItem.Click, AddressOf MarkRelatedConversation_Click
        treeContextMenu.Items.Add(markRelatedItem)
    End Sub



    Private Sub ListView_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        Try
            ' 使用主题颜色绘制列头
            Using headerBackBrush As New SolidBrush(currentBackColor),
                  headerTextBrush As New SolidBrush(currentForeColor),
                  sf As New StringFormat(),
                  borderPen As New Pen(Color.FromArgb(180, currentForeColor.R, currentForeColor.G, currentForeColor.B))

                ' 填充列头背景
                e.Graphics.FillRectangle(headerBackBrush, e.Bounds)

                ' 设置字符串格式
                sf.Alignment = StringAlignment.Near
                sf.LineAlignment = StringAlignment.Center
                sf.Trimming = StringTrimming.EllipsisCharacter

                ' 为文本添加左边距，避免紧贴边框
                Dim textRect As New Rectangle(e.Bounds.X + 4, e.Bounds.Y, e.Bounds.Width - 8, e.Bounds.Height)
                e.Graphics.DrawString(e.Header.Text, Me.Font, headerTextBrush, textRect, sf)

                ' 绘制右边框和底边框，形成分隔线效果
                e.Graphics.DrawLine(borderPen, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom - 1)
                e.Graphics.DrawLine(borderPen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right - 1, e.Bounds.Bottom - 1)
            End Using
        Catch ex As System.Exception
            ' 如果自定义绘制失败，回退到默认绘制
            e.DrawDefault = True
            Debug.WriteLine($"ListView_DrawColumnHeader异常: {ex.Message}")
        End Try
    End Sub

    Private Sub PendingMailList_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        Try
            ' 使用主题颜色绘制待办邮件ListView项目
            Dim backgroundColor As Color = If(e.Item.BackColor = Color.Empty, currentBackColor, e.Item.BackColor)

            ' 根据背景色动态调整文字颜色，确保高对比度
            Dim textColor As Color = currentForeColor
            If e.Item.BackColor = highlightColor Then
                ' 高亮行使用对比度更高的文字颜色
                textColor = If(IsDarkTheme(highlightColor), Color.White, Color.Black)
            End If

            Using backBrush As Brush = New SolidBrush(backgroundColor),
                  textBrush As Brush = New SolidBrush(textColor),
                  sf As New StringFormat()

                e.Graphics.FillRectangle(backBrush, e.Bounds)

                ' 绘制项目文本
                sf.Alignment = StringAlignment.Near
                sf.LineAlignment = StringAlignment.Center
                sf.Trimming = StringTrimming.EllipsisCharacter

                e.Graphics.DrawString(e.Item.Text, Me.Font, textBrush, e.Bounds, sf)

                ' 绘制子项目
                e.DrawDefault = False
                For i As Integer = 0 To e.Item.SubItems.Count - 1
                    If i < pendingMailList.Columns.Count Then
                        Dim subItemBounds As Rectangle = e.Item.GetBounds(ItemBoundsPortion.Entire)
                        Dim columnWidth As Integer = pendingMailList.Columns(i).Width
                        Dim x As Integer = 0
                        For j As Integer = 0 To i - 1
                            x += pendingMailList.Columns(j).Width
                        Next
                        subItemBounds = New Rectangle(x, subItemBounds.Y, columnWidth, subItemBounds.Height)

                        Using subItemBackBrush As Brush = New SolidBrush(backgroundColor)
                            e.Graphics.FillRectangle(subItemBackBrush, subItemBounds)
                            e.Graphics.DrawString(e.Item.SubItems(i).Text, Me.Font, textBrush, subItemBounds, sf)
                        End Using
                    End If
                Next
            End Using
        Catch ex As System.Exception
            ' 如果绘制失败，使用默认绘制
            e.DrawDefault = True
            Debug.WriteLine($"PendingMailList_DrawItem 异常: {ex.Message}")
        End Try
    End Sub

    Private Sub TaskList_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        Try
            ' 使用主题颜色绘制任务ListView项目
            Dim backgroundColor As Color = If(e.Item.BackColor = Color.Empty, currentBackColor, e.Item.BackColor)

            ' 根据背景色动态调整文字颜色，确保高对比度
            Dim textColor As Color = currentForeColor
            If e.Item.BackColor = highlightColor Then
                ' 高亮行使用对比度更高的文字颜色
                textColor = If(IsDarkTheme(highlightColor), Color.White, Color.Black)
            End If

            Using backBrush As Brush = New SolidBrush(backgroundColor),
                  textBrush As Brush = New SolidBrush(textColor),
                  sf As New StringFormat()

                e.Graphics.FillRectangle(backBrush, e.Bounds)

                ' 绘制项目文本
                sf.Alignment = StringAlignment.Near
                sf.LineAlignment = StringAlignment.Center
                sf.Trimming = StringTrimming.EllipsisCharacter

                e.Graphics.DrawString(e.Item.Text, Me.Font, textBrush, e.Bounds, sf)

                ' 绘制子项目
                e.DrawDefault = False
                For i As Integer = 0 To e.Item.SubItems.Count - 1
                    If i < taskList.Columns.Count Then
                        Dim subItemBounds As Rectangle = e.Item.GetBounds(ItemBoundsPortion.Entire)
                        Dim columnWidth As Integer = taskList.Columns(i).Width
                        Dim x As Integer = 0
                        For j As Integer = 0 To i - 1
                            x += taskList.Columns(j).Width
                        Next
                        subItemBounds = New Rectangle(x, subItemBounds.Y, columnWidth, subItemBounds.Height)

                        Using subItemBackBrush As Brush = New SolidBrush(backgroundColor)
                            e.Graphics.FillRectangle(subItemBackBrush, subItemBounds)
                            e.Graphics.DrawString(e.Item.SubItems(i).Text, Me.Font, textBrush, subItemBounds, sf)
                        End Using
                    End If
                Next
            End Using
        Catch ex As System.Exception
            ' 如果绘制失败，使用默认绘制
            e.DrawDefault = True
            Debug.WriteLine($"TaskList_DrawItem 异常: {ex.Message}")
        End Try
    End Sub

    Private Sub MailHistoryList_DrawItem(sender As Object, e As DrawListViewItemEventArgs)
        Try
            ' 使用主题颜色绘制邮件历史ListView项目
            Dim backgroundColor As Color = If(e.Item.BackColor = Color.Empty, currentBackColor, e.Item.BackColor)

            ' 根据背景色动态调整文字颜色，确保高对比度
            Dim textColor As Color = currentForeColor
            If e.Item.BackColor = highlightColor Then
                ' 高亮行使用对比度更高的文字颜色
                textColor = If(IsDarkTheme(highlightColor), Color.White, Color.Black)
            End If

            Using backBrush As Brush = New SolidBrush(backgroundColor),
                  textBrush As Brush = New SolidBrush(textColor),
                  sf As New StringFormat()

                e.Graphics.FillRectangle(backBrush, e.Bounds)

                ' 绘制项目文本
                sf.Alignment = StringAlignment.Near
                sf.LineAlignment = StringAlignment.Center
                sf.Trimming = StringTrimming.EllipsisCharacter

                e.Graphics.DrawString(e.Item.Text, Me.Font, textBrush, e.Bounds, sf)

                ' 绘制子项目
                e.DrawDefault = False
                For i As Integer = 0 To e.Item.SubItems.Count - 1
                    If i < mailHistoryList.Columns.Count Then
                        Dim subItemBounds As Rectangle = e.Item.GetBounds(ItemBoundsPortion.Entire)
                        Dim columnWidth As Integer = mailHistoryList.Columns(i).Width
                        Dim x As Integer = 0
                        For j As Integer = 0 To i - 1
                            x += mailHistoryList.Columns(j).Width
                        Next
                        subItemBounds = New Rectangle(x, subItemBounds.Y, columnWidth, subItemBounds.Height)

                        Using subItemBackBrush As Brush = New SolidBrush(backgroundColor)
                            e.Graphics.FillRectangle(subItemBackBrush, subItemBounds)
                            e.Graphics.DrawString(e.Item.SubItems(i).Text, Me.Font, textBrush, subItemBounds, sf)
                        End Using
                    End If
                Next
            End Using
        Catch ex As System.Exception
            ' 如果绘制失败，使用默认绘制
            e.DrawDefault = True
            Debug.WriteLine($"MailHistoryList_DrawItem 异常: {ex.Message}")
        End Try
    End Sub

    Private Sub ListView_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        Try
            ' 使用ListView的背景色或项目的背景色（如果项目有特殊背景色）
            Dim backgroundColor As Color = If(e.Item.BackColor = Color.Empty, currentBackColor, e.Item.BackColor)

            ' 根据背景色动态调整文字颜色，确保高对比度
            Dim textColor As Color = currentForeColor
            If e.Item.BackColor = highlightColor Then
                ' 高亮行使用对比度更高的文字颜色
                textColor = If(IsDarkTheme(highlightColor), Color.White, Color.Black)
            End If

            Using backBrush As New SolidBrush(backgroundColor),
                  sf As New StringFormat(),
                  textBrush As New SolidBrush(textColor)

                ' 填充背景
                e.Graphics.FillRectangle(backBrush, e.Bounds)

                ' 设置字符串格式
                sf.Trimming = StringTrimming.EllipsisCharacter
                sf.FormatFlags = StringFormatFlags.NoWrap

                If e.ColumnIndex = 0 Then
                    If e.SubItem.Text.Contains("🚩") Then
                        ' 使用特殊颜色和字体
                        Using specialFont As New Font(iconFont, FontStyle.Bold)
                            e.Graphics.DrawString(e.SubItem.Text, specialFont, Brushes.Red, e.Bounds, sf)
                        End Using
                    Else
                        e.Graphics.DrawString(e.SubItem.Text, iconFont, textBrush, e.Bounds, sf)
                    End If
                Else
                    ' 根据是否高亮使用不同字体
                    Dim font As Font = If(e.Item.BackColor = highlightColor, highlightFont, normalFont)
                    e.Graphics.DrawString(e.SubItem.Text, font, textBrush, e.Bounds, sf)
                End If
            End Using
        Catch ex As System.Exception
            ' 如果自定义绘制失败，回退到默认绘制
            e.DrawDefault = True
            Debug.WriteLine($"ListView_DrawSubItem异常: {ex.Message}")
        End Try
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

    Private Sub TabControl_DrawItem(sender As Object, e As DrawItemEventArgs)
        ' 使用主题颜色绘制Tab标签页
        Dim tabControl As TabControl = DirectCast(sender, TabControl)
        Dim tabPage As TabPage = tabControl.TabPages(e.Index)

        ' 确定是否为选中的标签页
        Dim isSelected As Boolean = (e.Index = tabControl.SelectedIndex)

        ' 设置背景色和文字色 - 优化黑色主题下的对比度
        Dim backColor As Color
        Dim textColor As Color = currentForeColor

        If isSelected Then
            backColor = currentBackColor
        Else
            If IsDarkTheme(currentBackColor) Then
                ' 黑色主题：未选中标签页使用稍亮的背景色
                backColor = Color.FromArgb(
                    Math.Min(255, currentBackColor.R + 25),
                    Math.Min(255, currentBackColor.G + 25),
                    Math.Min(255, currentBackColor.B + 25)
                )
            Else
                ' 浅色主题：未选中标签页使用稍暗的背景色
                backColor = Color.FromArgb(
                    Math.Max(0, currentBackColor.R - 20),
                    Math.Max(0, currentBackColor.G - 20),
                    Math.Max(0, currentBackColor.B - 20)
                )
            End If
        End If

        ' 填充背景
        Using backBrush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(backBrush, e.Bounds)
        End Using

        ' 绘制文字
        Using textBrush As New SolidBrush(textColor)
            Dim sf As New StringFormat()
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            e.Graphics.DrawString(tabPage.Text, Me.Font, textBrush, e.Bounds, sf)
            sf.Dispose()
        End Using

        ' 绘制边框（可选）
        If isSelected Then
            Using borderPen As New Pen(currentForeColor)
                e.Graphics.DrawRectangle(borderPen, e.Bounds)
            End Using
        End If
    End Sub


    Private Sub SetupTabPages()
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill,
            .DrawMode = TabDrawMode.OwnerDrawFixed
        }
        ' 不在这里设置颜色，等待ApplyTheme方法调用
        splitter2.Panel2.Controls.Add(tabControl)

        ' 设置四个tab
        SetupMailInteractionTab()  ' 来往邮件tab
        SetupPendingMailTab()      ' 待办邮件tab
        SetupNotesTab()            ' 笔记tab
        SetupTaskManagementTab()   ' 任务tab

        ' 添加自定义绘制事件处理器
        AddHandler tabControl.DrawItem, AddressOf TabControl_DrawItem

        ' 设置默认选中第一个tab
        tabControl.SelectedIndex = 0

        ' 延迟应用主题，确保所有控件都已初始化
        Me.BeginInvoke(Sub()
                           If needsThemeApplication Then
                               Debug.WriteLine("延迟应用主题到TabPages...")
                               ApplyTheme(currentBackColor, currentForeColor)
                           End If
                       End Sub)
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



    Private Sub GetAllMailFolders(folder As Microsoft.Office.Interop.Outlook.Folder, folderList As List(Of Microsoft.Office.Interop.Outlook.Folder))
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
                          isMailItem = (folder.DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
                      End Sub)

            If isMailItem AndAlso coreFolders.Contains(folder.Name) Then
                folderList.Add(folder)
                Debug.WriteLine($"添加邮件文件夹: {folder.Name}")
            End If

            ' 只在核心文件夹中递归搜索
            Dim subFolders As Microsoft.Office.Interop.Outlook.Folders = Nothing
            Me.Invoke(Sub()
                          subFolders = folder.Folders
                      End Sub)

            If subFolders IsNot Nothing Then
                For Each subFolder As Microsoft.Office.Interop.Outlook.Folder In subFolders
                    If coreFolders.Contains(subFolder.Name) Then
                        GetAllMailFolders(subFolder, folderList)
                    End If
                Next
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"处理文件夹 {folder.Name} 时出错: {ex.Message}")
        End Try
    End Sub

    ' SMTP地址转换缓存
    Private Shared smtpAddressCache As New Dictionary(Of String, String)

    ''' <summary>
    ''' 将Exchange内部地址转换为SMTP地址
    ''' </summary>
    ''' <param name="exchangeAddress">Exchange内部地址</param>
    ''' <returns>SMTP地址，如果转换失败则返回原地址</returns>
    Private Function GetSMTPAddress(exchangeAddress As String) As String
        Try
            ' 基本验证
            If String.IsNullOrEmpty(exchangeAddress) Then
                Return String.Empty
            End If

            ' 如果已经是SMTP格式，直接返回
            If exchangeAddress.Contains("@") AndAlso Not exchangeAddress.StartsWith("/O=") Then
                Return exchangeAddress
            End If

            ' 检查缓存（遵循全局缓存开关）
            If IsCacheEnabled() AndAlso smtpAddressCache.ContainsKey(exchangeAddress) Then
                Return smtpAddressCache(exchangeAddress)
            End If

            ' 如果不是Exchange内部地址格式，返回原地址
            If Not exchangeAddress.StartsWith("/O=") Then
                If IsCacheEnabled() Then smtpAddressCache(exchangeAddress) = exchangeAddress
                Return exchangeAddress
            End If

            Dim smtpAddress As String = String.Empty
            Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing
            Dim recipient As Microsoft.Office.Interop.Outlook.Recipient = Nothing
            Dim addressEntry As Microsoft.Office.Interop.Outlook.AddressEntry = Nothing
            Dim exchangeUser As Microsoft.Office.Interop.Outlook.ExchangeUser = Nothing

            Try
                ' 策略1: 通过CreateRecipient和AddressEntry获取SMTP地址
                session = Globals.ThisAddIn.Application.GetNamespace("MAPI")
                recipient = session.CreateRecipient(exchangeAddress)

                If recipient IsNot Nothing Then
                    recipient.Resolve()
                    addressEntry = recipient.AddressEntry

                    If addressEntry IsNot Nothing Then
                        ' 尝试获取Exchange用户的SMTP地址
                        If addressEntry.AddressEntryUserType = Microsoft.Office.Interop.Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Then
                            exchangeUser = addressEntry.GetExchangeUser()
                            If exchangeUser IsNot Nothing AndAlso Not String.IsNullOrEmpty(exchangeUser.PrimarySmtpAddress) Then
                                smtpAddress = exchangeUser.PrimarySmtpAddress
                            End If
                        End If

                        ' 策略2: 如果Exchange用户方法失败，尝试使用Address属性
                        If String.IsNullOrEmpty(smtpAddress) AndAlso Not String.IsNullOrEmpty(addressEntry.Address) AndAlso addressEntry.Address.Contains("@") Then
                            smtpAddress = addressEntry.Address
                        End If

                        ' 策略3: 尝试使用PropertyAccessor获取SMTP地址
                        If String.IsNullOrEmpty(smtpAddress) Then
                            Try
                                Dim propertyAccessor As Microsoft.Office.Interop.Outlook.PropertyAccessor = addressEntry.PropertyAccessor
                                Dim smtpProp As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                                smtpAddress = TryCast(propertyAccessor.GetProperty(smtpProp), String)
                            Catch propEx As System.Exception
                                Debug.WriteLine($"PropertyAccessor获取SMTP地址失败: {propEx.Message}")
                            End Try
                        End If
                    End If
                End If

                ' 策略4: 如果所有方法都失败，尝试从Exchange地址中提取用户名
                If String.IsNullOrEmpty(smtpAddress) Then
                    Try
                        ' 从Exchange地址格式中提取用户名: /O=ORGANIZATION/OU=SITE/CN=RECIPIENTS/CN=USERNAME
                        Dim cnIndex As Integer = exchangeAddress.LastIndexOf("/CN=")
                        If cnIndex > 0 Then
                            Dim username As String = exchangeAddress.Substring(cnIndex + 4)
                            ' 移除可能的额外CN部分
                            Dim nextCnIndex As Integer = username.IndexOf("/CN=")
                            If nextCnIndex > 0 Then
                                username = username.Substring(0, nextCnIndex)
                            End If

                            ' 尝试构造SMTP地址（这是一个猜测，可能不准确）
                            If Not String.IsNullOrEmpty(username) AndAlso session IsNot Nothing Then
                                Try
                                    Dim defaultDomain As String = session.CurrentUser.AddressEntry.GetExchangeUser()?.PrimarySmtpAddress
                                    If Not String.IsNullOrEmpty(defaultDomain) AndAlso defaultDomain.Contains("@") Then
                                        Dim domain As String = defaultDomain.Substring(defaultDomain.IndexOf("@"))
                                        smtpAddress = username + domain
                                    End If
                                Catch domainEx As System.Exception
                                    Debug.WriteLine($"构造SMTP地址失败: {domainEx.Message}")
                                End Try
                            End If
                        End If
                    Catch extractEx As System.Exception
                        Debug.WriteLine($"从Exchange地址提取用户名失败: {extractEx.Message}")
                    End Try
                End If

            Finally
                ' 清理COM对象
                If exchangeUser IsNot Nothing Then Runtime.InteropServices.Marshal.ReleaseComObject(exchangeUser)
                If addressEntry IsNot Nothing Then Runtime.InteropServices.Marshal.ReleaseComObject(addressEntry)
                If recipient IsNot Nothing Then Runtime.InteropServices.Marshal.ReleaseComObject(recipient)
                If session IsNot Nothing Then Runtime.InteropServices.Marshal.ReleaseComObject(session)
            End Try

            ' 如果转换失败，返回原地址
            If String.IsNullOrEmpty(smtpAddress) Then
                smtpAddress = exchangeAddress
                Debug.WriteLine($"Exchange地址转换失败，使用原地址: {exchangeAddress}")
            Else
                Debug.WriteLine($"Exchange地址 '{exchangeAddress}' 转换为SMTP地址: '{smtpAddress}'")
            End If

            ' 缓存结果（遵循全局缓存开关）
            If IsCacheEnabled() Then smtpAddressCache(exchangeAddress) = smtpAddress
            Return smtpAddress

        Catch ex As System.Exception
            Debug.WriteLine($"转换Exchange地址为SMTP地址时出错: {ex.Message}")
            ' 出错时返回原地址
            Return exchangeAddress
        End Try
    End Function

    ' 添加一个新的辅助方法用于递归获取所有邮件文件夹
    Private Sub GetAllMailFoldersAll(folder As Microsoft.Office.Interop.Outlook.Folder, folderList As List(Of Microsoft.Office.Interop.Outlook.Folder))
        Try
            Me.Invoke(Sub()
                          ' 添加当前文件夹（如果是邮件文件夹）
                          If folder.DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                              folderList.Add(folder)
                          End If

                          ' 递归处理子文件夹
                          For Each subFolder As Microsoft.Office.Interop.Outlook.Folder In folder.Folders
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
            Dim currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem Is Nothing Then Return "未选择邮件项"

            Dim senderEmail As String = String.Empty
            Dim senderName As String = String.Empty

            ' 获取发件人信息
            If TypeOf currentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                Dim mail = DirectCast(currentItem, Microsoft.Office.Interop.Outlook.MailItem)
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
            ElseIf TypeOf currentItem Is Microsoft.Office.Interop.Outlook.MeetingItem Then
                Dim meeting = DirectCast(currentItem, Microsoft.Office.Interop.Outlook.MeetingItem)
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

            ' 转换Exchange地址为SMTP地址
            Dim originalEmail = senderEmail
            senderEmail = GetSMTPAddress(senderEmail)
            Debug.WriteLine($"GetContactInfoBackground - 原始邮箱: {originalEmail}, 转换后邮箱: {senderEmail}")

            ' 清理过期缓存
            CleanExpiredCache()

            ' 检查缓存（遵循全局缓存开关）
            If IsCacheEnabled() AndAlso contactMailCache.ContainsKey(senderEmail) Then
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
                    Dim contacts = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts)
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

            ' 检查会议统计缓存（遵循全局缓存开关）
            If IsCacheEnabled() AndAlso meetingStatsCache.ContainsKey(meetingCacheKey) AndAlso
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
                    Dim calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar)
                    Dim startDate = DateTime.Now.AddMonths(-2)
                    Dim endDate = DateTime.Now.AddMonths(1)

                    ' 构建两个过滤条件：一个获取与该联系人相关的会议（必要与会者），一个获取可选与会者
                    ' 移除结束时间限制，只限制开始时间，避免排除长期会议
                    Dim requiredFilter = $"[Start] >= '{startDate:MM/dd/yyyy}' AND [Start] <= '{endDate:MM/dd/yyyy}' AND ([RequiredAttendees] LIKE '%{senderEmail}%' OR [RequiredAttendees] LIKE '%{senderEmail.Split("@"c)(0)}%')"
                    Dim optionalFilter = $"[Start] >= '{startDate:MM/dd/yyyy}' AND [Start] <= '{endDate:MM/dd/yyyy}' AND ([OptionalAttendees] LIKE '%{senderEmail}%' OR [OptionalAttendees] LIKE '%{senderEmail.Split("@"c)(0)}%')"

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

                    ' 缓存会议统计结果（遵循全局缓存开关）
                    If IsCacheEnabled() Then meetingStatsCache(meetingCacheKey) = New MeetingStatsData With {
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
            Dim folders As New List(Of Microsoft.Office.Interop.Outlook.Folder)
            Dim store As Microsoft.Office.Interop.Outlook.Store = Globals.ThisAddIn.Application.Session.DefaultStore

            ' 获取收件箱及其指定子文件夹
            Dim inbox As Microsoft.Office.Interop.Outlook.Folder = TryCast(store.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox), Microsoft.Office.Interop.Outlook.Folder)
            If inbox IsNot Nothing Then
                folders.Add(inbox)
                For Each subFolder As Microsoft.Office.Interop.Outlook.Folder In inbox.Folders
                    If subFolder.Name.Equals("Doc", StringComparison.OrdinalIgnoreCase) OrElse
                       subFolder.Name.Equals("Processed Mail", StringComparison.OrdinalIgnoreCase) OrElse
                       subFolder.Name.Equals("Todo", StringComparison.OrdinalIgnoreCase) Then
                        folders.Add(subFolder)
                    End If
                Next
            End If

            ' 获取已发送邮件文件夹
            Dim sentItems As Microsoft.Office.Interop.Outlook.Folder = TryCast(store.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail), Microsoft.Office.Interop.Outlook.Folder)
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
            Debug.WriteLine($"搜索文件夹数量: {folders.Count}")

            ' 只获取最近3个月的最多100封邮件，不再统计总数
            Dim tempRecentMails As New List(Of (Received As DateTime, Subject As String))
            For Each folder In folders
                Try
                    Debug.WriteLine($"正在搜索文件夹: {folder.Name}")
                    ' 搜索双向邮件：对方发给我们的邮件
                    ' 同时搜索原始Exchange地址和转换后的SMTP地址
                    Dim mailFilter As String
                    If originalEmail <> senderEmail Then
                        ' 如果地址被转换了，同时搜索两种格式
                        mailFilter = $"([SenderEmailAddress] LIKE '%{originalEmail}%' OR [SenderEmailAddress] LIKE '%{senderEmail}%') AND [ReceivedTime] >= '{dateFilter}'"
                        Debug.WriteLine($"邮件搜索过滤条件(双地址): {mailFilter}")
                    Else
                        ' 如果地址没有转换，只搜索一种格式
                        mailFilter = $"[SenderEmailAddress] LIKE '%{senderEmail}%' AND [ReceivedTime] >= '{dateFilter}'"
                        Debug.WriteLine($"邮件搜索过滤条件(单地址): {mailFilter}")
                    End If
                    Dim table As Microsoft.Office.Interop.Outlook.Table = folder.GetTable(mailFilter)
                    table.Columns.Add("Subject")
                    table.Columns.Add("ReceivedTime")
                    ' 使用PR_ENTRYID获取长格式EntryID
                    table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")

                    Do Until table.EndOfTable OrElse tempRecentMails.Count >= 100
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

            ' 搜索我们发给对方的邮件（已发送邮件）
            Try
                Dim sentFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail)
                Dim sentFilter = $"[SentOn] >= '{dateFilter}'"
                Dim sentTable = sentFolder.GetTable(sentFilter)
                sentTable.Columns.Add("To")
                sentTable.Columns.Add("Subject")
                sentTable.Columns.Add("SentOn")

                Do Until sentTable.EndOfTable OrElse tempRecentMails.Count >= 100
                    Dim row = sentTable.GetNextRow()
                    Try
                        Dim toRecipients As String = If(row("To")?.ToString(), "")
                        ' 检查收件人是否包含目标邮件地址（同时检查两种格式）
                        Dim containsTarget As Boolean = False
                        If Not String.IsNullOrEmpty(toRecipients) Then
                            containsTarget = toRecipients.Contains(senderEmail)
                            If Not containsTarget AndAlso originalEmail <> senderEmail Then
                                containsTarget = toRecipients.Contains(originalEmail)
                            End If
                        End If
                        If containsTarget Then
                            Dim sentObj = row("SentOn")
                            Dim subjectObj = row("Subject")
                            Dim sentTime As DateTime = If(sentObj IsNot Nothing AndAlso Not String.IsNullOrEmpty(sentObj.ToString()), DateTime.Parse(sentObj.ToString()), DateTime.MinValue)
                            Dim subject As String = If(subjectObj IsNot Nothing, subjectObj.ToString(), "Unknown Subject")
                            tempRecentMails.Add((sentTime, $"[已发送] {subject}"))
                        End If
                    Catch
                        ' 忽略单个邮件获取错误
                    End Try
                Loop

                Runtime.InteropServices.Marshal.ReleaseComObject(sentTable)
                Runtime.InteropServices.Marshal.ReleaseComObject(sentFolder)
            Catch ex As System.Exception
                Debug.WriteLine($"搜索已发送邮件时出错: {ex.Message}")
            End Try

            ' 按时间排序并显示最近邮件，添加序号（不再生成可点击链接）
            recentMails = tempRecentMails.OrderByDescending(Function(m) m.Received).Take(100).ToList()

            swMail.Stop()
            elapsedMailStats = swMail.ElapsedMilliseconds

            info.AppendLine($"邮件往来统计:")
            info.AppendLine($"最近邮件 (最多100封):")

            For i As Integer = 0 To recentMails.Count - 1
                Dim m = recentMails(i)
                info.AppendLine($"- [{i + 1}] {m.Received:yyyy-MM-dd HH:mm} {m.Subject.Replace("[EXT]", "")}")
            Next

            ' 保存到缓存（遵循全局缓存开关）
            Dim result = info.ToString()
            If IsCacheEnabled() Then contactMailCache(senderEmail) = (result, DateTime.Now)
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
            Dim mailItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryID)
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
        Catch ex As System.Runtime.InteropServices.COMException
            ' 记录调试信息
            If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                Debug.WriteLine($"OpenOutlookMail COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
            End If

            ' 统一COM错误处理
            Globals.ThisAddIn.ShowErrorWithConfig("无法打开邮件，可能已被删除或移动", ex.Message)
        Catch ex As System.Exception
            ' 记录调试信息
            If ErrorNotificationSettings.Instance.LogErrorsToDebug Then
                Debug.WriteLine($"OpenOutlookMail 异常: {ex.Message}")
            End If

            ' 根据配置决定是否显示错误
            If ErrorNotificationSettings.Instance.ShowErrorDialogs Then
                ' 统一错误处理
                Globals.ThisAddIn.ShowErrorWithConfig("无法打开邮件，可能已被删除或移动", ex.Message)
            End If
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
            .ForeColor = currentForeColor,
            .View = System.Windows.Forms.View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .OwnerDraw = False  ' 使用默认绘制
        }
        OutlookMyList.Handlers.TaskHandler.SetupTaskList(taskList)
        taskList.Columns.Add("主题", 200)
        taskList.Columns.Add("到期日", 100)
        taskList.Columns.Add("状态", 100)
        taskList.Columns.Add("完成百分比", 100)
        taskList.Columns.Add("关联邮件", 200)


        ' Add the event handler here, after taskList is initialized
        AddHandler taskList.DoubleClick, AddressOf TaskList_DoubleClick
        ' AddHandler taskList.DrawColumnHeader, AddressOf ListView_DrawColumnHeader  ' 移除自定义绘制
        ' AddHandler taskList.DrawItem, AddressOf TaskList_DrawItem  ' 移除自定义绘制

        ' 检查是否需要应用主题（处理异步创建的时序问题）
        If needsThemeApplication Then
            taskList.BackColor = currentBackColor
            taskList.ForeColor = currentForeColor
            taskList.Refresh()
            needsThemeApplication = False
            Debug.WriteLine("taskList创建后应用了延迟的主题设置")
        End If

        Dim containerPanel As New Panel With {
            .Dock = DockStyle.Fill
        }
        containerPanel.Controls.Add(taskList)
        containerPanel.Controls.Add(taskButtonPanel)
        tabPage2.Controls.Add(containerPanel)
        tabControl.TabPages.Add(tabPage2)
    End Sub

    Private Sub SetupMailInteractionTab()
        Dim tabPage As New TabPage("来往邮件")
        btnPanel = New Panel With {
            .Dock = DockStyle.Fill
        }

        ' 创建按钮面板
        Dim buttonPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 20,
            .Visible = False
        }

        ' 创建TreeView替代ListView来展示联系人信息
        contactInfoTree = New TreeView With {
            .Dock = DockStyle.Fill,
            .FullRowSelect = True,
            .ShowLines = True,
            .ShowPlusMinus = True,
            .ShowRootLines = True,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor
        }



        ' 设置TreeView右键菜单
        SetupTreeContextMenu()
        contactInfoTree.ContextMenuStrip = treeContextMenu

        ' 添加双击事件处理邮件链接
        AddHandler contactInfoTree.NodeMouseDoubleClick, AddressOf ContactInfoTree_NodeMouseDoubleClick
        ' 添加单击事件处理邮件链接
        AddHandler contactInfoTree.AfterSelect, AddressOf ContactInfoTree_AfterSelect
        ' 添加拖拽事件（作为拖拽源）
        AddHandler contactInfoTree.ItemDrag, AddressOf contactInfoTree_ItemDrag

        ' 只创建按钮，不预先创建文本框
        Dim x As Integer = 10
        For i As Integer = 1 To 3
            Dim btn As New Button With {
                .Text = If(i = 1, "联系人信息", If(i = 2, $"会话信息({lvMails.Items.Count})", "邮件历史")),
                .Location = New Point(x, 2),
                .Size = New Size(100, 15),
                .Visible = True,
                .TextAlign = If(i = 2, ContentAlignment.MiddleLeft, ContentAlignment.MiddleCenter)
            }

            ' 保存会话信息按钮的引用
            If i = 2 Then
                conversationInfoButton = btn
            End If

            ' 为每个按钮添加双击隐藏功能
            Dim buttonIndex As Integer = i ' 捕获循环变量

            ' 单击事件 - 保留原有功能
            If i = 1 Then
                AddHandler btn.Click, Sub(s, e)
                                          ' 调用独立的邮件历史信息处理方法
                                          ShowMailHistoryInfo()
                                      End Sub
            Else
                AddHandler btn.Click, Sub(s, e)
                                          ' 显示会话信息
                                          contactInfoTree.Nodes.Clear()
                                          Dim rootNode As New TreeNode("会话信息")
                                          contactInfoTree.Nodes.Add(rootNode)

                                          Dim item1 As New TreeNode($"会话ID: {currentConversationId}")
                                          rootNode.Nodes.Add(item1)

                                          Dim item2 As New TreeNode($"邮件数量: {lvMails.Items.Count}")
                                          rootNode.Nodes.Add(item2)

                                          Dim item3 As New TreeNode($"当前邮件: {currentMailEntryID}")
                                          rootNode.Nodes.Add(item3)

                                          ' 默认不展开TreeView
                                      End Sub
            End If

            ' 双击事件 - 切换按钮显示隐藏
            AddHandler btn.DoubleClick, Sub(s, e)
                                            Select Case buttonIndex
                                                Case 1
                                                    button1Visible = Not button1Visible
                                                    btn.Visible = button1Visible
                                                Case 2
                                                    button2Visible = Not button2Visible
                                                    btn.Visible = button2Visible
                                                Case 3
                                                    button3Visible = Not button3Visible
                                                    btn.Visible = button3Visible
                                            End Select
                                        End Sub

            buttonPanel.Controls.Add(btn)
            x += 125
        Next

        ' 为按钮面板添加右键菜单以显示隐藏的按钮
        Dim contextMenu As New ContextMenuStrip()
        Dim showAllButtonsItem As New ToolStripMenuItem("显示所有按钮")
        AddHandler showAllButtonsItem.Click, Sub(s, e)
                                                 button1Visible = True
                                                 button2Visible = True
                                                 button3Visible = True
                                                 buttonPanel.Visible = True
                                                 For Each ctrl As Control In buttonPanel.Controls
                                                     If TypeOf ctrl Is Button Then
                                                         ctrl.Visible = True
                                                     End If
                                                 Next
                                             End Sub
        contextMenu.Items.Add(showAllButtonsItem)
        buttonPanel.ContextMenuStrip = contextMenu

        ' 为主面板也添加相同的右键菜单，以便在按钮面板隐藏时也能显示按钮
        btnPanel.ContextMenuStrip = contextMenu

        ' 先添加按钮面板到主面板（Dock Top）
        btnPanel.Controls.Add(buttonPanel)
        ' 再添加TreeView到主面板（Dock Fill）
        btnPanel.Controls.Add(contactInfoTree)

        tabPage.Controls.Add(btnPanel)
        tabControl.TabPages.Add(tabPage)
    End Sub

    Private Sub SetupPendingMailTab()
        Dim tabPage As New TabPage("待办邮件")

        ' 创建待办邮件的ListView（继承原有功能）
        pendingMailList = New ListView With {
            .View = System.Windows.Forms.View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .Dock = DockStyle.Fill,
            .Visible = True,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor,
            .OwnerDraw = False  ' 使用默认绘制
        }

        ' 设置ListView列
        pendingMailList.Columns.Add("时间", 120)
        pendingMailList.Columns.Add("主题", 250)
        pendingMailList.Columns.Add("Due Date", 120)

        ' 添加点击事件处理程序
        AddHandler pendingMailList.Click, AddressOf MailHistory_Click
        AddHandler pendingMailList.DoubleClick, AddressOf MailHistory_DoubleClick

        ' 添加自定义绘制事件处理程序以支持黑色主题
        pendingMailList.OwnerDraw = True
        AddHandler pendingMailList.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler pendingMailList.DrawItem, AddressOf PendingMailList_DrawItem
        AddHandler pendingMailList.DrawSubItem, AddressOf ListView_DrawSubItem

        tabPage.Controls.Add(pendingMailList)
        tabControl.TabPages.Add(tabPage)
    End Sub

    Private Sub SetupNotesTab()
        Dim tabPage As New TabPage("笔记")

        ' 创建笔记按钮面板
        Dim noteButtonPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 40
        }

        Dim btnAddNote As New Button With {
            .Text = "新建笔记",
            .Location = New Point(10, 5),
            .Size = New Size(80, 30)
        }
        ' 可以在这里添加新建笔记的事件处理
        noteButtonPanel.Controls.Add(btnAddNote)

        ' 创建笔记显示区域（可以是RichTextBox或其他控件）
        Dim noteTextBox As New RichTextBox With {
            .Dock = DockStyle.Fill,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor
        }

        Dim containerPanel As New Panel With {
            .Dock = DockStyle.Fill
        }
        containerPanel.Controls.Add(noteTextBox)
        containerPanel.Controls.Add(noteButtonPanel)

        tabPage.Controls.Add(containerPanel)
        tabControl.TabPages.Add(tabPage)
    End Sub

    ' UpdateNoteList方法，用于更新笔记列表
    Private Sub UpdateNoteList(noteList As List(Of (CreateTime As String, Title As String, Link As String)))
        ' 由于新的笔记tab结构简化，这里暂时不实现具体功能
        ' 可以在后续需要时添加具体的笔记列表更新逻辑
        Debug.WriteLine($"UpdateNoteList called with {noteList?.Count} notes")
    End Sub

    Private Sub SetupTaskManagementTab()
        Dim tabPage As New TabPage("任务")

        ' 创建任务按钮面板
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

        ' 创建任务列表
        taskList = New ListView With {
            .Dock = DockStyle.Fill,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor,
            .View = System.Windows.Forms.View.Details,
            .FullRowSelect = True,
            .GridLines = True
        }

        OutlookMyList.Handlers.TaskHandler.SetupTaskList(taskList)
        taskList.Columns.Add("主题", 200)
        taskList.Columns.Add("到期日", 100)
        taskList.Columns.Add("状态", 100)
        taskList.Columns.Add("完成百分比", 100)
        taskList.Columns.Add("关联邮件", 200)

        ' 添加自定义绘制事件处理程序以支持黑色主题
        taskList.OwnerDraw = True
        AddHandler taskList.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler taskList.DrawItem, AddressOf TaskList_DrawItem
        AddHandler taskList.DrawSubItem, AddressOf ListView_DrawSubItem

        ' Add the event handler here, after taskList is initialized
        AddHandler taskList.DoubleClick, AddressOf TaskList_DoubleClick

        Dim containerPanel As New Panel With {
            .Dock = DockStyle.Fill
        }
        containerPanel.Controls.Add(taskList)
        containerPanel.Controls.Add(taskButtonPanel)
        tabPage.Controls.Add(containerPanel)
        tabControl.TabPages.Add(tabPage)
    End Sub

    Private Sub SetupActionsTab()
        Dim tabPage3 As New TabPage("操作")
        btnPanel = New Panel With {
            .Dock = DockStyle.Fill
        }

        ' 创建按钮面板
        Dim buttonPanel As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 20,
            .Visible = False
        }

        ' 创建TreeView替代ListView来展示联系人信息
        contactInfoTree = New TreeView With {
            .Dock = DockStyle.Fill,
            .FullRowSelect = True,
            .ShowLines = True,
            .ShowPlusMinus = True,
            .ShowRootLines = True,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor
        }

        ' 创建ListView用于来往邮件功能
        mailHistoryList = New ListView With {
            .View = System.Windows.Forms.View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .Dock = DockStyle.None,
            .Visible = False,
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor,
            .OwnerDraw = True  ' 启用自定义绘制以支持黑色主题
        }

        ' 设置ListView列
        mailHistoryList.Columns.Add("时间", 120)
        mailHistoryList.Columns.Add("主题", 250)
        mailHistoryList.Columns.Add("Due Date", 120)

        ' 添加点击事件处理程序
        AddHandler mailHistoryList.Click, AddressOf MailHistory_Click
        AddHandler mailHistoryList.DoubleClick, AddressOf MailHistory_DoubleClick
        ' 添加自定义绘制事件处理程序以支持黑色主题
        AddHandler mailHistoryList.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler mailHistoryList.DrawItem, AddressOf MailHistoryList_DrawItem
        AddHandler mailHistoryList.DrawSubItem, AddressOf ListView_DrawSubItem
        ' AddHandler mailHistoryList.DrawItem, AddressOf MailHistoryList_DrawItem  ' 移除自定义绘制

        ' 设置TreeView右键菜单
        SetupTreeContextMenu()
        contactInfoTree.ContextMenuStrip = treeContextMenu

        ' 添加双击事件处理邮件链接
        AddHandler contactInfoTree.NodeMouseDoubleClick, AddressOf ContactInfoTree_NodeMouseDoubleClick
        ' 添加单击事件处理邮件链接
        AddHandler contactInfoTree.AfterSelect, AddressOf ContactInfoTree_AfterSelect
        ' 添加拖拽事件（作为拖拽源）
        AddHandler contactInfoTree.ItemDrag, AddressOf contactInfoTree_ItemDrag

        ' 只创建按钮，不预先创建文本框
        Dim x As Integer = 10
        For i As Integer = 1 To 3
            Dim btn As New Button With {
                .Text = If(i = 1, "联系人信息", $"按钮 {i}"),
                .Location = New Point(x, 2),
                .Size = New Size(100, 15),
                .Visible = True
            }

            ' 为每个按钮添加双击隐藏功能
            Dim buttonIndex As Integer = i ' 捕获循环变量

            ' 单击事件 - 保留原有功能
            If i = 1 Then
                AddHandler btn.Click, Sub(s, e)
                                          ' 调用独立的邮件历史信息处理方法
                                          ShowMailHistoryInfo()
                                      End Sub
            Else
                AddHandler btn.Click, Sub(s, e)
                                          ' 显示会话信息
                                          contactInfoTree.Nodes.Clear()
                                          Dim rootNode As New TreeNode("会话信息")
                                          contactInfoTree.Nodes.Add(rootNode)

                                          Dim item1 As New TreeNode($"会话ID: {currentConversationId}")
                                          rootNode.Nodes.Add(item1)

                                          Dim item2 As New TreeNode($"邮件数量: {lvMails.Items.Count}")
                                          rootNode.Nodes.Add(item2)

                                          Dim item3 As New TreeNode($"当前邮件: {currentMailEntryID}")
                                          rootNode.Nodes.Add(item3)

                                          ' 默认不展开TreeView
                                      End Sub
            End If

            ' 双击事件 - 切换按钮显示隐藏
            AddHandler btn.DoubleClick, Sub(s, e)
                                            Select Case buttonIndex
                                                Case 1
                                                    button1Visible = Not button1Visible
                                                    btn.Visible = button1Visible
                                                Case 2
                                                    button2Visible = Not button2Visible
                                                    btn.Visible = button2Visible
                                                Case 3
                                                    button3Visible = Not button3Visible
                                                    btn.Visible = button3Visible
                                            End Select
                                        End Sub

            buttonPanel.Controls.Add(btn)
            x += 125
        Next

        ' 为按钮面板添加右键菜单以显示隐藏的按钮
        Dim contextMenu As New ContextMenuStrip()
        Dim showAllButtonsItem As New ToolStripMenuItem("显示所有按钮")
        AddHandler showAllButtonsItem.Click, Sub(s, e)
                                                 button1Visible = True
                                                 button2Visible = True
                                                 button3Visible = True
                                                 buttonPanel.Visible = True
                                                 For Each ctrl As Control In buttonPanel.Controls
                                                     If TypeOf ctrl Is Button Then
                                                         ctrl.Visible = True
                                                     End If
                                                 Next
                                             End Sub
        contextMenu.Items.Add(showAllButtonsItem)
        buttonPanel.ContextMenuStrip = contextMenu

        ' 为主面板也添加相同的右键菜单，以便在按钮面板隐藏时也能显示按钮
        btnPanel.ContextMenuStrip = contextMenu

        ' 先添加按钮面板到主面板（Dock Top）
        btnPanel.Controls.Add(buttonPanel)
        ' 再添加TreeView到主面板（Dock Fill）
        btnPanel.Controls.Add(contactInfoTree)
        ' 添加ListView到主面板（用于待办邮件功能）
        btnPanel.Controls.Add(mailHistoryList)

        tabPage3.Controls.Add(btnPanel)
        tabControl.TabPages.Add(tabPage3)
    End Sub

    ' 新增：联系人信息树支持与双击打开邮件
    Private Async Sub GetContactInfoTreeHandler()
        Try
            If contactInfoTree Is Nothing Then Return

            ' 在开始收集联系人信息时立即抑制 WebView 更新
            suppressWebViewUpdate += 1

            ' 显示进度指示器
            'ShowProgress("正在收集联系人来往邮件信息...")

            contactInfoTree.Nodes.Clear()
            Dim loading As New TreeNode("正在收集联系人来往邮件信息...")
            contactInfoTree.Nodes.Add(loading)

            Dim result = Await GetContactInfoData(CancellationToken)

            ' 检查是否被取消
            If CancellationToken.IsCancellationRequested Then
                Return
            End If

            If Me.InvokeRequired Then
                Me.Invoke(Sub() PopulateContactInfoTree(result))
            Else
                PopulateContactInfoTree(result)
            End If
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("联系人来往邮件信息收集被取消")
        Catch ex As System.Exception
            Debug.WriteLine("GetContactInfoTreeHandler error: " & ex.Message)
        Finally
            ' 隐藏进度指示器并释放抑制计数器
            HideProgress()
            suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
        End Try
    End Sub

    ' 带发件人参数的联系人信息处理方法
    Private Async Sub GetContactInfoTreeHandlerWithSender(senderName As String, senderEmail As String)
        Debug.WriteLine($"=== GetContactInfoTreeHandlerWithSender 开始，发件人: {senderName}, 邮箱: {senderEmail} ===")
        Try
            If contactInfoTree Is Nothing Then
                Debug.WriteLine("contactInfoTree 为 Nothing，退出")
                Return
            End If

            Debug.WriteLine("contactInfoTree 存在，继续处理")

            ' 在开始收集联系人信息时立即抑制 WebView 更新
            suppressWebViewUpdate += 1
            Debug.WriteLine($"suppressWebViewUpdate 增加到: {suppressWebViewUpdate}")

            ' 显示进度指示器
            Debug.WriteLine("显示进度指示器...")
            ShowProgress("正在收集联系人来往邮件信息...")

            Debug.WriteLine("清空并添加加载节点...")
            contactInfoTree.Nodes.Clear()
            Dim loading As New TreeNode("正在收集联系人来往邮件信息...")
            contactInfoTree.Nodes.Add(loading)

            Debug.WriteLine("开始异步获取联系人数据...")
            Dim result = Await Task.Run(Function() GetContactInfoDataWithSender(senderName, senderEmail, CancellationToken))
            Debug.WriteLine("联系人数据获取完成")

            ' 检查是否被取消
            If CancellationToken.IsCancellationRequested Then
                Debug.WriteLine("操作被取消")
                Return
            End If

            Debug.WriteLine("开始填充联系人树...")
            If Me.InvokeRequired Then
                Debug.WriteLine("使用Invoke填充联系人树")
                Me.Invoke(Sub() PopulateContactInfoTree(result))
            Else
                Debug.WriteLine("直接填充联系人树")
                PopulateContactInfoTree(result)
            End If
            Debug.WriteLine("联系人树填充完成")
        Catch ex As System.OperationCanceledException
            Debug.WriteLine("联系人来往邮件信息收集被取消")
        Catch ex As System.Exception
            Debug.WriteLine("GetContactInfoTreeHandlerWithSender error: " & ex.Message)
            Debug.WriteLine($"GetContactInfoTreeHandlerWithSender 异常堆栈: {ex.StackTrace}")
        Finally
            ' 隐藏进度指示器并释放抑制计数器
            Debug.WriteLine("隐藏进度指示器并释放抑制计数器...")
            HideProgress()
            suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
            Debug.WriteLine($"suppressWebViewUpdate 减少到: {suppressWebViewUpdate}")
            Debug.WriteLine("=== GetContactInfoTreeHandlerWithSender 结束 ===")
        End Try
    End Sub

    ' 生成联系人信息的结构化数据
    Private Async Function GetContactInfoData(Optional cancellationToken As Threading.CancellationToken = Nothing) As Task(Of (SenderName As String, SenderEmail As String, MeetingStats As Dictionary(Of String, Integer), Upcoming As List(Of (MeetingDate As DateTime, Title As String, EntryID As String)), MailCount As Integer, RecentMailIds As List(Of (EntryID As String, Subject As String, Received As DateTime)), ConversationGroups As Dictionary(Of String, List(Of (EntryID As String, Subject As String, Received As DateTime)))))
        Dim senderName As String = ""
        Dim senderEmail As String = ""
        Dim meetingStats As New Dictionary(Of String, Integer)
        Dim upcoming As New List(Of (DateTime, String, String))
        Dim mailCount As Integer = 0
        Dim recentMails As New List(Of (String, String, DateTime))
        Dim conversationGroups As New Dictionary(Of String, List(Of (String, String, DateTime)))
        Try
            Dim currentItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem Is Nothing Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)

            If TypeOf currentItem Is Outlook.MailItem Then
                Dim mail = DirectCast(currentItem, Outlook.MailItem)
                Try
                    senderEmail = mail.SenderEmailAddress
                    senderName = mail.SenderName
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常获取邮件发件人信息 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
                Catch ex As System.Exception
                    Debug.WriteLine($"获取邮件发件人信息时发生异常: {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
                End Try
            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                Dim meeting = DirectCast(currentItem, Outlook.MeetingItem)
                Try
                    senderEmail = meeting.SenderEmailAddress
                    senderName = meeting.SenderName
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常获取会议发件人信息 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
                Catch ex As System.Exception
                    Debug.WriteLine($"获取会议发件人信息时发生异常: {ex.Message}")
                    Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
                End Try
            End If
            If String.IsNullOrEmpty(senderEmail) Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)

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
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            Catch ex As System.Exception
                Debug.WriteLine($"获取日历文件夹时发生异常: {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End Try

            If meetings Is Nothing Then
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End If
            Dim meetingsCount As Integer = meetings.Count
            For i = meetingsCount To 1 Step -1
                Dim ap As Microsoft.Office.Interop.Outlook.AppointmentItem = Nothing
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
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            Catch ex As System.Exception
                Debug.WriteLine($"获取邮件存储时发生异常: {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End Try

            Debug.WriteLine($"找到 {folders.Count} 个邮件文件夹")
            If folders.Count = 0 Then
                Debug.WriteLine("没有找到任何邮件文件夹")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End If
            Dim dateFilter = DateTime.Now.AddMonths(-3).ToString("MM/dd/yyyy")
            Dim tasks As New List(Of Task(Of (Count As Integer, Mails As List(Of (EntryID As String, Subject As String, Received As DateTime)))))
            For Each folder In folders
                tasks.Add(Task.Run(Function()
                                       Try
                                           ' 搜索对方发给我们的邮件
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
                                           Debug.WriteLine($"文件夹 {folder.Name} 找到 {count} 封邮件")
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
            Dim searchResults = Await Task.WhenAll(tasks)
            For Each r In searchResults
                mailCount += r.Count
                For Each mail In r.Mails
                    recentMails.Add((mail.EntryID, mail.Subject, mail.Received))

                    ' 使用智能会话识别进行分组
                    Dim groupKey As String = String.Empty
                    Try
                        Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mail.EntryID)
                        groupKey = GetSmartConversationKey(mailItem)

                        ' 如果智能识别失败，回退到简单主题分组
                        If String.IsNullOrEmpty(groupKey) Then
                            groupKey = $"主题:{GetNormalizedSubject(mail.Subject)}"
                        End If

                        ' 释放COM对象
                        If mailItem IsNot Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine($"获取邮件分组信息失败: {ex.Message}")
                        ' 异常时使用智能主题分组
                        groupKey = $"主题:{GetNormalizedSubject(mail.Subject)}"
                    End Try

                    If Not conversationGroups.ContainsKey(groupKey) Then
                        conversationGroups(groupKey) = New List(Of (String, String, DateTime))
                    End If
                    conversationGroups(groupKey).Add((mail.EntryID, mail.Subject, mail.Received))
                Next
            Next
            Debug.WriteLine($"总共找到 {mailCount} 封邮件，最近邮件 {recentMails.Count} 封")
            recentMails = recentMails.OrderByDescending(Function(m) m.Item3).Take(50).ToList()
            ' 对每个会话内的邮件按时间排序
            For Each kvp In conversationGroups.ToList()
                conversationGroups(kvp.Key) = kvp.Value.OrderByDescending(Function(m) m.Item3).ToList()
            Next
        Catch ex As System.Exception
            Debug.WriteLine("GetContactInfoData error: " & ex.Message)
        End Try
        Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
    End Function

    ' 带发件人参数的联系人信息数据获取方法
    Private Function GetContactInfoDataWithSender(senderName As String, senderEmail As String, Optional cancellationToken As Threading.CancellationToken = Nothing) As (SenderName As String, SenderEmail As String, MeetingStats As Dictionary(Of String, Integer), Upcoming As List(Of (MeetingDate As DateTime, Title As String, EntryID As String)), MailCount As Integer, RecentMailIds As List(Of (EntryID As String, Subject As String, Received As DateTime)), ConversationGroups As Dictionary(Of String, List(Of (EntryID As String, Subject As String, Received As DateTime))))
        Debug.WriteLine($"=== GetContactInfoDataWithSender 开始，发件人: {senderName}, 邮箱: {senderEmail} ===")
        Dim meetingStats As New Dictionary(Of String, Integer)
        Dim upcoming As New List(Of (DateTime, String, String))
        Dim mailCount As Integer = 0
        Dim recentMails As New List(Of (String, String, DateTime))
        Dim conversationGroups As New Dictionary(Of String, List(Of (String, String, DateTime)))

        Try
            If String.IsNullOrEmpty(senderEmail) Then
                Debug.WriteLine("发件人邮箱为空，返回空结果")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End If

            ' 转换Exchange内部地址为SMTP地址
            Dim originalEmail = senderEmail
            senderEmail = GetSMTPAddress(senderEmail)
            If originalEmail <> senderEmail Then
                Debug.WriteLine($"Exchange地址转换: {originalEmail} -> {senderEmail}")
            End If

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
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            Catch ex As System.Exception
                Debug.WriteLine($"获取日历文件夹时发生异常: {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End Try

            If meetings IsNot Nothing Then
                Dim meetingsCount As Integer = meetings.Count
                For i = meetingsCount To 1 Step -1
                    If cancellationToken.IsCancellationRequested Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)

                    Dim ap As Microsoft.Office.Interop.Outlook.AppointmentItem = Nothing
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
                        If match.Success Then
                            projectName = match.Groups(1).Value
                        End If

                        If meetingStats.ContainsKey(projectName) Then
                            meetingStats(projectName) += 1
                        Else
                            meetingStats(projectName) = 1
                        End If

                        If startTime > DateTime.Now Then
                            upcoming.Add((startTime, subject, entryId))
                        End If
                    End If
                Next
            End If

            ' 邮件统计和最近邮件 - 使用GetAllMailFolders获取所有核心文件夹
            Dim allMailFolders As New List(Of Outlook.Folder)
            Try
                ' 获取根文件夹并使用GetAllMailFolders函数获取所有核心文件夹
                Dim rootFolder As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
                GetAllMailFolders(rootFolder, allMailFolders)

                ' 如果没有找到文件夹，至少添加收件箱作为备选
                If allMailFolders.Count = 0 Then
                    allMailFolders.Add(DirectCast(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), Outlook.Folder))
                End If

                Debug.WriteLine($"使用GetAllMailFolders找到 {allMailFolders.Count} 个核心文件夹进行搜索")
                For Each folder In allMailFolders
                    Debug.WriteLine($"  - {folder.Name}")
                Next

            Catch ex As System.Runtime.InteropServices.COMException
                Debug.WriteLine($"COM异常获取邮件文件夹 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            Catch ex As System.Exception
                Debug.WriteLine($"获取邮件文件夹时发生异常: {ex.Message}")
                Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
            End Try

            ' 处理所有核心文件夹中的邮件 - 使用多个字段进行搜索
            Debug.WriteLine($"开始搜索所有核心文件夹中的邮件，原始地址: {originalEmail}, SMTP地址: {senderEmail}, 发件人姓名: {senderName}")

            For Each folder As Outlook.Folder In allMailFolders
                Try
                    Debug.WriteLine($"正在搜索文件夹: {folder.Name}")

                    ' 方法1: 使用SMTP地址搜索
                    Dim filter As String = $"[SenderEmailAddress] = '{senderEmail}'"
                    Dim items = folder.Items.Restrict(filter)

                    Debug.WriteLine($"方法1 - 在{folder.Name}中使用SMTP地址过滤器: {filter}")
                    Debug.WriteLine($"在{folder.Name}中找到 {items.Count} 封匹配的邮件")

                    For Each item As Object In items
                        If cancellationToken.IsCancellationRequested Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)

                        If TypeOf item Is Outlook.MailItem Then
                            Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                            Try
                                ' 检查是否已经添加过这封邮件（避免重复）
                                Dim alreadyAdded = recentMails.Any(Function(m) m.Item1 = mail.EntryID)
                                If Not alreadyAdded Then
                                    mailCount += 1
                                    recentMails.Add((mail.EntryID, mail.Subject, mail.ReceivedTime))

                                    If Not String.IsNullOrEmpty(mail.ConversationTopic) Then
                                        If Not conversationGroups.ContainsKey(mail.ConversationTopic) Then
                                            conversationGroups(mail.ConversationTopic) = New List(Of (String, String, DateTime))
                                        End If
                                        conversationGroups(mail.ConversationTopic).Add((mail.EntryID, mail.Subject, mail.ReceivedTime))
                                    End If

                                    Debug.WriteLine($"{folder.Name}邮件(SMTP): {mail.Subject}, 发件人: {mail.SenderEmailAddress}")
                                Else
                                    Debug.WriteLine($"邮件已存在，跳过: {mail.Subject}")
                                End If
                            Catch ex As System.Runtime.InteropServices.COMException
                                Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                            Catch ex As System.Exception
                                Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                            End Try
                        End If
                    Next

                    ' 方法2: 使用Exchange地址搜索
                    If originalEmail <> senderEmail Then
                        Debug.WriteLine($"方法2 - 在{folder.Name}中尝试原始Exchange地址: {originalEmail}")
                        filter = $"[SenderEmailAddress] = '{originalEmail}'"
                        items = folder.Items.Restrict(filter)

                        Debug.WriteLine($"在{folder.Name}中使用Exchange地址过滤器: {filter}")
                        Debug.WriteLine($"在{folder.Name}中找到 {items.Count} 封匹配的邮件")

                        For Each item As Object In items
                            If cancellationToken.IsCancellationRequested Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)

                            If TypeOf item Is Outlook.MailItem Then
                                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                                Try
                                    ' 检查是否已经添加过这封邮件（避免重复）
                                    Dim alreadyAdded = recentMails.Any(Function(m) m.Item1 = mail.EntryID)
                                    If Not alreadyAdded Then
                                        mailCount += 1
                                        recentMails.Add((mail.EntryID, mail.Subject, mail.ReceivedTime))

                                        If Not String.IsNullOrEmpty(mail.ConversationTopic) Then
                                            If Not conversationGroups.ContainsKey(mail.ConversationTopic) Then
                                                conversationGroups(mail.ConversationTopic) = New List(Of (String, String, DateTime))
                                            End If
                                            conversationGroups(mail.ConversationTopic).Add((mail.EntryID, mail.Subject, mail.ReceivedTime))
                                        End If

                                        Debug.WriteLine($"{folder.Name}邮件(Exchange): {mail.Subject}, 发件人: {mail.SenderEmailAddress}")
                                    Else
                                        Debug.WriteLine($"邮件已存在，跳过: {mail.Subject}")
                                    End If
                                Catch ex As System.Runtime.InteropServices.COMException
                                    Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                                Catch ex As System.Exception
                                    Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                                End Try
                            End If
                        Next
                    End If

                    ' 方法3: 使用发件人姓名搜索
                    If Not String.IsNullOrEmpty(senderName) Then
                        Debug.WriteLine($"方法3 - 在{folder.Name}中尝试发件人姓名: {senderName}")
                        filter = $"[SenderName] = '{senderName}'"
                        items = folder.Items.Restrict(filter)

                        Debug.WriteLine($"在{folder.Name}中使用发件人姓名过滤器: {filter}")
                        Debug.WriteLine($"在{folder.Name}中找到 {items.Count} 封匹配的邮件")

                        For Each item As Object In items
                            If cancellationToken.IsCancellationRequested Then Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)

                            If TypeOf item Is Outlook.MailItem Then
                                Dim mail As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                                Try
                                    ' 检查是否已经添加过这封邮件（避免重复）
                                    Dim alreadyAdded = recentMails.Any(Function(m) m.Item1 = mail.EntryID)
                                    If Not alreadyAdded Then
                                        mailCount += 1
                                        recentMails.Add((mail.EntryID, mail.Subject, mail.ReceivedTime))

                                        If Not String.IsNullOrEmpty(mail.ConversationTopic) Then
                                            If Not conversationGroups.ContainsKey(mail.ConversationTopic) Then
                                                conversationGroups(mail.ConversationTopic) = New List(Of (String, String, DateTime))
                                            End If
                                            conversationGroups(mail.ConversationTopic).Add((mail.EntryID, mail.Subject, mail.ReceivedTime))
                                        End If

                                        Debug.WriteLine($"{folder.Name}邮件(姓名): {mail.Subject}, 发件人: {mail.SenderName}, 邮箱: {mail.SenderEmailAddress}")
                                    Else
                                        Debug.WriteLine($"邮件已存在，跳过: {mail.Subject}")
                                    End If
                                Catch ex As System.Runtime.InteropServices.COMException
                                    Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                                Catch ex As System.Exception
                                    Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                                End Try
                            End If
                        Next
                    End If
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常处理{folder.Name}邮件 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                Catch ex As System.Exception
                    Debug.WriteLine($"处理{folder.Name}邮件时发生异常: {ex.Message}")
                End Try
            Next

            ' 完成所有文件夹搜索
            Debug.WriteLine($"完成所有核心文件夹搜索")

            ' 添加搜索结果总结
            Debug.WriteLine($"邮件搜索完成 - 总邮件数: {mailCount}, 最近邮件数: {recentMails.Count}, 会话组数: {conversationGroups.Count}")

            ' 排序
            recentMails = recentMails.OrderByDescending(Function(m) m.Item3).Take(50).ToList()
            upcoming = upcoming.OrderBy(Function(m) m.Item1).Take(10).ToList()

            For Each kvp In conversationGroups.ToList()
                conversationGroups(kvp.Key) = kvp.Value.OrderByDescending(Function(m) m.Item3).ToList()
            Next

        Catch ex As System.Runtime.InteropServices.COMException
            Debug.WriteLine($"GetContactInfoDataWithSender COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
            Debug.WriteLine($"GetContactInfoDataWithSender COM异常堆栈: {ex.StackTrace}")
        Catch ex As System.Exception
            Debug.WriteLine("GetContactInfoDataWithSender error: " & ex.Message)
            Debug.WriteLine($"GetContactInfoDataWithSender 异常堆栈: {ex.StackTrace}")
        End Try

        Debug.WriteLine($"=== GetContactInfoDataWithSender 结束，邮件数量: {mailCount}, 最近邮件: {recentMails.Count}, 会话组: {conversationGroups.Count} ===")
        Return (senderName, senderEmail, meetingStats, upcoming, mailCount, recentMails, conversationGroups)
    End Function

    Private Sub PopulateContactInfoTree(result As (SenderName As String, SenderEmail As String, MeetingStats As Dictionary(Of String, Integer), Upcoming As List(Of (MeetingDate As DateTime, Title As String, EntryID As String)), MailCount As Integer, RecentMailIds As List(Of (EntryID As String, Subject As String, Received As DateTime)), ConversationGroups As Dictionary(Of String, List(Of (EntryID As String, Subject As String, Received As DateTime)))))
        ' 保存会话分组数据到类级别变量
        currentConversationGroups = result.ConversationGroups

        ' 在填充联系人树期间抑制 WebView 更新
        suppressWebViewUpdate += 1
        contactInfoTree.BeginUpdate()
        Try
            contactInfoTree.Nodes.Clear()

            ' 添加基本信息节点
            Dim infoNode As New TreeNode($"联系人信息: {result.SenderName} ({result.SenderEmail})")
            contactInfoTree.Nodes.Add(infoNode)

            ' 添加会议信息节点
            Dim totalMeetings = result.MeetingStats.Values.Sum()
            Dim meetingNode As New TreeNode($"会议统计(近2月): 总计{totalMeetings}次")
            infoNode.Nodes.Add(meetingNode)
            For Each kv In result.MeetingStats.OrderByDescending(Function(x) x.Value)
                Dim projectNode As New TreeNode($"{kv.Key}: {kv.Value}次")
                meetingNode.Nodes.Add(projectNode)
            Next

            ' 添加即将会议节点
            If result.Upcoming.Count > 0 Then
                Dim upcomingNode As New TreeNode("即将会议")
                infoNode.Nodes.Add(upcomingNode)
                For Each up In result.Upcoming.OrderBy(Function(x) x.MeetingDate).Take(3)
                    Dim meetingItem As New TreeNode($"{up.MeetingDate.ToString("MM/dd HH:mm")} - {up.Title}")
                    meetingItem.Tag = up.EntryID
                    upcomingNode.Nodes.Add(meetingItem)
                Next
            End If

            ' 添加邮件会话节点
            Dim mailRootNode As New TreeNode($"邮件往来: 总计{result.MailCount}封")
            contactInfoTree.Nodes.Add(mailRootNode)

            ' 按会话分组显示邮件，按最新邮件时间排序
            Dim sortedConversations = result.ConversationGroups.OrderByDescending(Function(kvp) kvp.Value.Max(Function(m) m.Item3))

            For Each conversation In sortedConversations
                Dim convId = conversation.Key
                Dim mails = conversation.Value

                If mails.Count = 1 Then
                    ' 只有一封邮件时，也添加会话前缀保持对齐
                    Dim mail = mails.First()
                    Dim mailNode As New TreeNode($"会话 (1封) - {mail.Item3.ToString("yyyy-MM-dd HH:mm")} - {mail.Item2.Replace("[EXT]", "")}")
                    mailNode.Tag = mail.Item1 ' EntryID
                    mailRootNode.Nodes.Add(mailNode)
                Else
                    ' 多封邮件时，创建会话节点
                    Dim latestMail = mails.First() ' 已经按时间排序
                    Dim convNode As New TreeNode($"会话 ({mails.Count}封) - {latestMail.Item3.ToString("yyyy-MM-dd HH:mm")} - {latestMail.Item2.Replace("[EXT]", "")}")
                    convNode.Tag = $"CONVERSATION:{convId}"
                    mailRootNode.Nodes.Add(convNode)

                    ' 添加该会话的所有邮件
                    For Each mail In mails
                        Dim mailNode As New TreeNode($"{mail.Item3.ToString("yyyy-MM-dd HH:mm")} - {mail.Item2.Replace("[EXT]", "")}")
                        mailNode.Tag = mail.Item1 ' EntryID
                        convNode.Nodes.Add(mailNode)
                    Next
                End If
            Next

            ' 默认不展开TreeView，用户可手动展开

        Finally
            contactInfoTree.EndUpdate()
            suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
        End Try
    End Sub

    Private Sub ContactInfoTree_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs)
        Try
            ' 抑制模式下不响应双击
            If suppressWebViewUpdate > 0 Then Return

            If e.Node Is Nothing Then Return
            Dim entryId = TryCast(e.Node.Tag, String)

            ' 如果选择的是会话节点，自动选择该会话中最新的邮件
            If Not String.IsNullOrEmpty(entryId) AndAlso entryId.StartsWith("CONVERSATION:") Then
                ' 从entryId中提取会话标识符
                Dim conversationKey As String = entryId.Substring("CONVERSATION:".Length)

                ' 从currentConversationGroups中获取对应会话的邮件列表
                If currentConversationGroups IsNot Nothing AndAlso currentConversationGroups.ContainsKey(conversationKey) Then
                    Dim conversationMails = currentConversationGroups(conversationKey)
                    If conversationMails.Count > 0 Then
                        ' 获取最新邮件的EntryID（列表已按时间降序排列）
                        Dim latestMailEntryId As String = conversationMails(0).EntryID
                        If Not String.IsNullOrEmpty(latestMailEntryId) Then
                            ' 创建一个虚拟节点来处理最新邮件
                            Dim virtualNode As New TreeNode()
                            virtualNode.Tag = latestMailEntryId
                            Dim latestMailEventArgs As New TreeViewEventArgs(virtualNode, TreeViewAction.ByMouse)
                            ContactInfoTree_AfterSelect(sender, latestMailEventArgs)
                        End If
                    End If
                End If
                Return
            End If

            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                ' 增加隔离标志，避免与 lvMails 联动或触发 WebView 刷新冲突
                suppressWebViewUpdate += 1
                Try
                    SafeOpenOutlookMail(entryId)
                Finally
                    suppressWebViewUpdate = Math.Max(0, suppressWebViewUpdate - 1)
                End Try
            End If
        Catch ex As System.Exception
            Debug.WriteLine("ContactInfoTree_NodeMouseDoubleClick error: " & ex.Message)
        End Try
    End Sub

    ' TreeView右键菜单事件处理方法
    Private Sub ShowMailId_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                MessageBox.Show($"邮件ID: {entryId}", "邮件信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ShowMailId_Click error: {ex.Message}")
        End Try
    End Sub

    Private Sub ShowTreeConversationId_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                Try
                    Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
                    Dim conversationId As String = GetSafeConversationID(mailItem)

                    If Not String.IsNullOrEmpty(conversationId) Then
                        MessageBox.Show($"会话ID: {conversationId}", "会话信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("无法获取会话ID", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                    ' 释放COM对象
                    If mailItem IsNot Nothing Then
                        Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                    End If
                Catch ex As System.Exception
                    Globals.ThisAddIn.ShowErrorWithConfig("获取会话ID失败", ex.Message)
                End Try
            Else
                MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ShowTreeConversationId_Click error: {ex.Message}")
        End Try
    End Sub

    Private Sub ShowTreeTaskStatus_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                ' 检查任务关联状态
                Dim taskInfo = OutlookMyList.Handlers.TaskHandler.GetTaskByMailEntryID(entryId)
                If taskInfo IsNot Nothing Then
                    MessageBox.Show($"任务关联信息:\n主题: {taskInfo.Subject}\n状态: {taskInfo.Status}\n完成百分比: {taskInfo.PercentComplete}%\n到期日: {taskInfo.DueDate}", "任务关联状态", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("该邮件未关联任务", "任务关联状态", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ShowTreeTaskStatus_Click error: {ex.Message}")
            Globals.ThisAddIn.ShowErrorWithConfig("获取任务关联状态失败", ex.Message)
        End Try
    End Sub

    Private Sub CopyMailId_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                Clipboard.SetText(entryId)
                MessageBox.Show("邮件ID已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"CopyMailId_Click error: {ex.Message}")
        End Try
    End Sub

    Private Sub OpenInOutlook_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                SafeOpenOutlookMail(entryId)
            Else
                ' 记录调试信息
                If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                    Debug.WriteLine("OpenInOutlook_Click: 所选节点不是邮件节点")
                End If

                ' 根据配置决定是否显示错误提示
                If ThisAddIn.ErrorSettings.ShowErrorDialogs Then
                    MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch ex As System.Exception
            ' 记录调试信息
            If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                Debug.WriteLine($"OpenInOutlook_Click error: {ex.Message}")
            End If
        End Try
    End Sub

    Private Sub ShowSmartConversationId_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                Try
                    Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
                    Dim smartKey As String = GetSmartConversationKey(mailItem)

                    If Not String.IsNullOrEmpty(smartKey) Then
                        MessageBox.Show($"智能会话ID: {smartKey}", "智能会话信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("无法生成智能会话ID", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If

                    ' 释放COM对象
                    If mailItem IsNot Nothing Then
                        Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                    End If
                Catch ex As System.Exception
                    Globals.ThisAddIn.ShowErrorWithConfig("获取智能会话ID失败", ex.Message)
                End Try
            Else
                MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ShowSmartConversationId_Click error: {ex.Message}")
        End Try
    End Sub

    Private Sub MarkRelatedConversation_Click(sender As Object, e As EventArgs)
        Try
            Dim selectedNode = contactInfoTree.SelectedNode
            If selectedNode Is Nothing Then Return

            Dim entryId = TryCast(selectedNode.Tag, String)
            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                Try
                    Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
                    Dim smartKey As String = GetSmartConversationKey(mailItem)

                    ' 在当前会话中查找具有相同智能会话ID的邮件
                    Dim relatedMails As New List(Of String)
                    For Each item As ListViewItem In lvMails.Items
                        Try
                            Dim itemEntryId As String = item.Tag?.ToString()
                            If Not String.IsNullOrEmpty(itemEntryId) Then
                                Dim itemMail As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(itemEntryId)
                                Dim itemSmartKey As String = GetSmartConversationKey(itemMail)

                                If itemSmartKey = smartKey Then
                                    relatedMails.Add(item.SubItems(1).Text) ' 主题
                                End If

                                If itemMail IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(itemMail)
                                End If
                            End If
                        Catch
                            ' 忽略单个邮件处理错误
                        End Try
                    Next

                    If relatedMails.Count > 1 Then
                        Dim message As String = $"找到 {relatedMails.Count} 封相关邮件：\n\n" & String.Join("\n", relatedMails.Take(10))
                        If relatedMails.Count > 10 Then
                            message &= "\n...（还有更多）"
                        End If
                        MessageBox.Show(message, "相关会话邮件", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("未找到其他相关邮件", "相关会话邮件", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                    ' 释放COM对象
                    If mailItem IsNot Nothing Then
                        Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                    End If
                Catch ex As System.Exception
                    Globals.ThisAddIn.ShowErrorWithConfig("查找相关会话失败", ex.Message)
                End Try
            Else
                MessageBox.Show("所选节点不是邮件节点", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"MarkRelatedConversation_Click error: {ex.Message}")
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

    ' 智能主题标准化函数，去除Re:、FW:等前缀
    Private Function GetNormalizedSubject(subject As String) As String
        If String.IsNullOrEmpty(subject) Then
            Return "无主题"
        End If

        ' 去除常见的邮件前缀
        Dim normalizedSubject As String = subject.Trim()

        ' 循环去除前缀，直到没有更多前缀
        Dim hasPrefix As Boolean = True
        While hasPrefix
            hasPrefix = False

            ' 英文前缀
            If normalizedSubject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(3).Trim()
                hasPrefix = True
            ElseIf normalizedSubject.StartsWith("RE:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(3).Trim()
                hasPrefix = True
            ElseIf normalizedSubject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(3).Trim()
                hasPrefix = True
            ElseIf normalizedSubject.StartsWith("FWD:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(4).Trim()
                hasPrefix = True
            ElseIf normalizedSubject.StartsWith("Forward:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(8).Trim()
                hasPrefix = True
                ' 中文前缀
            ElseIf normalizedSubject.StartsWith("回复:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(3).Trim()
                hasPrefix = True
            ElseIf normalizedSubject.StartsWith("转发:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(3).Trim()
                hasPrefix = True
            ElseIf normalizedSubject.StartsWith("答复:", StringComparison.OrdinalIgnoreCase) Then
                normalizedSubject = normalizedSubject.Substring(3).Trim()
                hasPrefix = True
            End If
        End While

        ' 进一步清理主题：去除多余空格、特殊字符等
        normalizedSubject = System.Text.RegularExpressions.Regex.Replace(normalizedSubject, "\s+", " ").Trim()

        ' 如果处理后为空，返回默认值
        If String.IsNullOrEmpty(normalizedSubject) Then
            Return "无主题"
        End If

        Return normalizedSubject
    End Function

    ''' <summary>
    ''' 智能会话识别函数，基于多个维度识别同一溯源的会话
    ''' </summary>
    ''' <param name="mailItem">邮件项</param>
    ''' <returns>智能会话标识符</returns>
    Private Const CustomConversationPropertyName As String = "CustomConversationId"

    Public Function ReadCustomConversationIdFromItem(mailItem As Object) As String
        Try
            If mailItem Is Nothing Then Return String.Empty

            Dim userProps As Outlook.UserProperties = Nothing
            If TypeOf mailItem Is Outlook.MailItem Then
                userProps = DirectCast(mailItem, Outlook.MailItem).UserProperties
            ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                userProps = DirectCast(mailItem, Outlook.AppointmentItem).UserProperties
            ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                userProps = DirectCast(mailItem, Outlook.MeetingItem).UserProperties
            End If

            If userProps IsNot Nothing Then
                Dim prop = userProps.Find(CustomConversationPropertyName)
                If prop IsNot Nothing AndAlso prop.Value IsNot Nothing Then
                    Dim val As String = prop.Value.ToString()
                    If Not String.IsNullOrWhiteSpace(val) Then Return val
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ReadCustomConversationIdFromItem error: {ex.Message}")
        End Try
        Return String.Empty
    End Function

    Private Function GetCustomConversationIdByEntryID(entryID As String) As String
        Try
            If String.IsNullOrEmpty(entryID) Then Return String.Empty
            Dim item = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryID)
            If item Is Nothing Then Return String.Empty
            Dim customId = ReadCustomConversationIdFromItem(item)
            If Not String.IsNullOrEmpty(customId) Then Return customId
            ' 回落：使用原始ConversationID
            Return GetSafeConversationID(item)
        Catch ex As System.Exception
            Debug.WriteLine($"GetCustomConversationIdByEntryID error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Public Function SetCustomConversationIdByEntryID(entryID As String, convId As String, Optional storeId As String = Nothing) As Boolean
        Try
            Debug.WriteLine($"SetCustomConversationIdByEntryID: entryID={entryID}, convId={convId}, storeId={(If(String.IsNullOrEmpty(storeId), "(null)", storeId))}")
            If String.IsNullOrEmpty(entryID) Then
                Debug.WriteLine("SetCustomConversationIdByEntryID: entryID为空，返回False")
                Return False
            End If

            ' 注意：允许convId为空字符串，这表示要清除自定义会话ID
            ' 只有当convId为Nothing时才返回False
            If convId Is Nothing Then
                Debug.WriteLine("SetCustomConversationIdByEntryID: convId为Nothing，返回False")
                Return False
            End If

            Dim item = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryID, storeId)
            If item Is Nothing Then
                Debug.WriteLine("SetCustomConversationIdByEntryID: 无法获取邮件项目，返回False")
                Return False
            End If
            Debug.WriteLine($"SetCustomConversationIdByEntryID: 成功获取邮件项目，类型: {item.GetType().Name}")

            Dim userProps As Outlook.UserProperties = Nothing
            Try
                If TypeOf item Is Outlook.MailItem Then
                    userProps = DirectCast(item, Outlook.MailItem).UserProperties
                    Debug.WriteLine("SetCustomConversationIdByEntryID: 邮件类型为MailItem")
                ElseIf TypeOf item Is Outlook.AppointmentItem Then
                    userProps = DirectCast(item, Outlook.AppointmentItem).UserProperties
                    Debug.WriteLine("SetCustomConversationIdByEntryID: 邮件类型为AppointmentItem")
                ElseIf TypeOf item Is Outlook.MeetingItem Then
                    userProps = DirectCast(item, Outlook.MeetingItem).UserProperties
                    Debug.WriteLine("SetCustomConversationIdByEntryID: 邮件类型为MeetingItem")
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"SetCustomConversationIdByEntryID: 获取UserProperties时出错: {ex.Message}")
                Return False
            End Try

            If userProps Is Nothing Then
                Debug.WriteLine("SetCustomConversationIdByEntryID: 无法获取UserProperties，返回False")
                Return False
            End If

            Dim prop As Outlook.UserProperty = Nothing
            Try
                prop = userProps.Find(CustomConversationPropertyName)

                ' 如果convId为空字符串，表示要清除自定义会话ID
                If String.IsNullOrEmpty(convId) Then
                    If prop IsNot Nothing Then
                        Debug.WriteLine($"SetCustomConversationIdByEntryID: 删除自定义属性 {CustomConversationPropertyName}")
                        prop.Delete()
                        Debug.WriteLine("SetCustomConversationIdByEntryID: 自定义属性已删除")
                    Else
                        Debug.WriteLine("SetCustomConversationIdByEntryID: 自定义属性不存在，无需删除")
                    End If
                Else
                    ' 设置或更新自定义会话ID
                    If prop Is Nothing Then
                        Debug.WriteLine($"SetCustomConversationIdByEntryID: 创建新的自定义属性 {CustomConversationPropertyName}")
                        prop = userProps.Add(CustomConversationPropertyName, Outlook.OlUserPropertyType.olText)
                    Else
                        Debug.WriteLine($"SetCustomConversationIdByEntryID: 找到现有的自定义属性 {CustomConversationPropertyName}")
                    End If
                    prop.Value = convId
                    Debug.WriteLine($"SetCustomConversationIdByEntryID: 设置属性值为: {convId}")
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"SetCustomConversationIdByEntryID: 设置属性值时出错: {ex.Message}")
                Return False
            End Try

            ' 保存更改
            Try
                ' 确保属性已添加到UserProperties集合
                If prop Is Nothing Then
                    Debug.WriteLine("SetCustomConversationIdByEntryID: 属性对象为空，无法保存")
                    Return False
                End If

                ' 强制保存属性
                prop.Value = convId

                ' 保存邮件项
                If TypeOf item Is Outlook.MailItem Then
                    Dim mailItem = DirectCast(item, Outlook.MailItem)
                    mailItem.Save()

                    ' 验证保存是否成功
                    Dim savedProp = mailItem.UserProperties.Find(CustomConversationPropertyName)
                    If savedProp IsNot Nothing AndAlso savedProp.Value.ToString() = convId Then
                        Debug.WriteLine("SetCustomConversationIdByEntryID: MailItem已成功保存并验证")
                    Else
                        Debug.WriteLine("SetCustomConversationIdByEntryID: MailItem保存后验证失败")
                        Return False
                    End If
                ElseIf TypeOf item Is Outlook.AppointmentItem Then
                    DirectCast(item, Outlook.AppointmentItem).Save()
                    Debug.WriteLine("SetCustomConversationIdByEntryID: AppointmentItem已保存")
                ElseIf TypeOf item Is Outlook.MeetingItem Then
                    DirectCast(item, Outlook.MeetingItem).Save()
                    Debug.WriteLine("SetCustomConversationIdByEntryID: MeetingItem已保存")
                End If

                Debug.WriteLine("SetCustomConversationIdByEntryID: 操作成功完成")
                Return True
            Catch ex As System.Exception
                Debug.WriteLine($"SetCustomConversationIdByEntryID: 保存邮件项时出错: {ex.Message}")
                Return False
            End Try
        Catch ex As System.Exception
            Debug.WriteLine($"SetCustomConversationIdByEntryID error: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function GetSmartConversationKey(mailItem As Object) As String
        Try
            If mailItem Is Nothing Then
                Return String.Empty
            End If

            ' 智能会话ID功能已禁用 - 优先使用自定义会话ID或原生会话ID
            Dim conversationId As String = ""

            ' 获取邮件基本信息
            If TypeOf mailItem Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(mailItem, Outlook.MailItem)
                conversationId = If(mail.ConversationID, "")
            ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(mailItem, Outlook.AppointmentItem)
                conversationId = If(appointment.ConversationID, "")
            ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                conversationId = If(meeting.ConversationID, "")
            End If

            ' 优先使用自定义会话ID（如果存在且有效）
            Dim customId As String = ReadCustomConversationIdFromItem(mailItem)
            If Not String.IsNullOrEmpty(customId) Then
                Return $"conv:{customId}"
            End If

            ' 次优先使用原生ConversationID
            If Not String.IsNullOrEmpty(conversationId) Then
                Return $"conv:{conversationId}"
            End If

            ' 智能会话ID功能已禁用，不再生成智能会话ID
            Return String.Empty
        Catch ex As System.Exception
            Debug.WriteLine($"GetSmartConversationKey error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    ' 安全获取ConversationID的方法
    Private Function GetSafeConversationID(mailItem As Object) As String
        Try
            If mailItem Is Nothing Then
                Return String.Empty
            End If

            ' 尝试直接获取ConversationID
            If TypeOf mailItem Is Outlook.MailItem Then
                Dim mail As Outlook.MailItem = DirectCast(mailItem, Outlook.MailItem)
                Return If(mail.ConversationID, String.Empty)
            ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                Dim appointment As Outlook.AppointmentItem = DirectCast(mailItem, Outlook.AppointmentItem)
                Return If(appointment.ConversationID, String.Empty)
            ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                Return If(meeting.ConversationID, String.Empty)
            End If

        Catch ex As System.Runtime.InteropServices.COMException
            ' COM异常，可能是某些文件夹不支持ConversationID
            Debug.WriteLine($"获取ConversationID时发生COM异常: {ex.Message}")
            Return String.Empty
        Catch ex As System.Exception
            Debug.WriteLine($"获取ConversationID时发生异常: {ex.Message}")
            Return String.Empty
        End Try

        Return String.Empty
    End Function

    ' 获取当前激活邮件的会话ID
    Private Function GetCurrentActiveMailConversationId() As String
        Try
            ' 首先尝试使用已存储的currentConversationId
            If Not String.IsNullOrEmpty(currentConversationId) Then
                Debug.WriteLine($"使用已存储的会话ID: {currentConversationId}")
                Return currentConversationId
            End If

            ' 如果没有存储的会话ID，尝试从当前邮件获取
            If Not String.IsNullOrEmpty(currentMailEntryID) Then
                Debug.WriteLine($"从当前邮件EntryID获取会话ID: {currentMailEntryID}")
                Dim currentItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem IsNot Nothing Then
                    Dim conversationId As String = GetSafeConversationID(currentItem)
                    If Not String.IsNullOrEmpty(conversationId) Then
                        Debug.WriteLine($"从当前邮件获取到会话ID: {conversationId}")
                        Return conversationId
                    End If
                End If
            End If

            ' 最后尝试从Outlook当前选中的邮件获取
            Try
                Dim outlookApp As Outlook.Application = Globals.ThisAddIn.Application
                If outlookApp IsNot Nothing AndAlso outlookApp.ActiveExplorer IsNot Nothing Then
                    Dim selection As Outlook.Selection = outlookApp.ActiveExplorer.Selection
                    If selection IsNot Nothing AndAlso selection.Count > 0 Then
                        Dim selectedItem As Object = selection.Item(1)
                        If selectedItem IsNot Nothing Then
                            Dim conversationId As String = GetSafeConversationID(selectedItem)
                            If Not String.IsNullOrEmpty(conversationId) Then
                                Debug.WriteLine($"从Outlook选中邮件获取到会话ID: {conversationId}")
                                Return conversationId
                            End If
                        End If
                    End If
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"从Outlook获取当前选中邮件失败: {ex.Message}")
            End Try

            Debug.WriteLine("无法获取当前激活邮件的会话ID")
            Return String.Empty

        Catch ex As System.Exception
            Debug.WriteLine($"获取当前激活邮件会话ID时出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Sub ContactInfoTree_AfterSelect(sender As Object, e As TreeViewEventArgs)
        Try
            ' 允许在本窗格中点击联系人邮件时总是更新右侧 mailBrowser
            ' 抑制标志仅用于避免与外部触发的刷新串扰，不用于本地点击后的内容展示

            If e.Node Is Nothing Then Return
            Dim entryId = TryCast(e.Node.Tag, String)

            ' 如果选择的是会话节点，自动选择该会话中最新的邮件
            If Not String.IsNullOrEmpty(entryId) AndAlso entryId.StartsWith("CONVERSATION:") Then
                ' 从entryId中提取会话标识符
                Dim conversationKey As String = entryId.Substring("CONVERSATION:".Length)

                ' 从currentConversationGroups中获取对应会话的邮件列表
                If currentConversationGroups IsNot Nothing AndAlso currentConversationGroups.ContainsKey(conversationKey) Then
                    Dim conversationMails = currentConversationGroups(conversationKey)
                    If conversationMails.Count > 0 Then
                        ' 获取最新邮件的EntryID（列表已按时间降序排列）
                        Dim latestMailEntryId As String = conversationMails(0).EntryID
                        If Not String.IsNullOrEmpty(latestMailEntryId) Then
                            ' 创建一个虚拟节点来处理最新邮件
                            Dim virtualNode As New TreeNode()
                            virtualNode.Tag = latestMailEntryId
                            Dim latestMailEventArgs As New TreeViewEventArgs(virtualNode, TreeViewAction.ByMouse)
                            ContactInfoTree_AfterSelect(sender, latestMailEventArgs)
                        End If
                    End If
                End If
                Return
            End If

            If Not String.IsNullOrEmpty(entryId) AndAlso Not entryId.StartsWith("CONVERSATION:") Then
                ' 本地点击不抬高抑制计数（保持为局部更新）
                Try
                    Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
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
                    ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                        Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                        Try
                            Dim subject As String = If(String.IsNullOrEmpty(meeting.Subject), "无主题", meeting.Subject)
                            Dim senderName As String = If(String.IsNullOrEmpty(meeting.SenderName), "未知", meeting.SenderName)
                            Dim receivedTime As String = If(meeting.ReceivedTime = DateTime.MinValue, "未知", meeting.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
                            Dim body As String = If(String.IsNullOrEmpty(meeting.Body), "", ReplaceTableTag(meeting.Body))

                            displayContent = $"<html><body style='font-family: Arial; padding: 10px; Font-size:12px;'>" &
                                $"<h4 style='color: var(--theme-color, #0078d7);'>{subject}</h4>" &
                                $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                                $"<strong style='color: var(--theme-color, #0078d7);'>发件人:</strong> {senderName}<br/>" &
                                $"<strong style='color: var(--theme-color, #0078d7);'>时间:</strong> {receivedTime}" &
                                $"</div>" &
                                $"<div style='border-top: 1px solid var(--theme-color, #0078d7); padding-top: 10px;'>" &
                                $"<style>.hidden-table {{display: none;}} img {{display: none;}}</style>" &
                                $"{body}" &
                                $"</div>" &
                                "</body></html>"
                        Catch ex As System.Runtime.InteropServices.COMException
                            Debug.WriteLine($"COM异常访问会议邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                            displayContent = "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议邮件属性</body></html>"
                        Catch ex As System.Exception
                            Debug.WriteLine($"访问会议邮件属性时发生异常: {ex.Message}")
                            displayContent = "<html><body style='font-family: Arial; padding: 10px;'>无法访问会议邮件属性</body></html>"
                        End Try
                    End If
                    ' 本地点击：始终更新当前窗格的 WebView
                    DisplayMailInWebView(entryId)
                    Return
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
            Debug.WriteLine("ContactInfoTree_AfterSelect error: " & ex.Message)
        End Try
    End Sub

    ' 启用拖拽：从联系人树或列表收集EntryID并拖入lvMails进行会话合并
    Private Sub contactInfoTree_ItemDrag(sender As Object, e As ItemDragEventArgs)
        Try
            Dim node As TreeNode = TryCast(e.Item, TreeNode)
            If node Is Nothing Then Return

            Dim ids As New List(Of String)
            Dim tagStr As String = TryCast(node.Tag, String)
            If Not String.IsNullOrEmpty(tagStr) Then
                If tagStr.StartsWith("CONVERSATION:") Then
                    Dim conversationKey As String = tagStr.Substring("CONVERSATION:".Length)
                    If currentConversationGroups IsNot Nothing AndAlso currentConversationGroups.ContainsKey(conversationKey) Then
                        For Each m In currentConversationGroups(conversationKey)
                            If Not String.IsNullOrEmpty(m.EntryID) Then ids.Add(m.EntryID)
                        Next
                    End If
                Else
                    ids.Add(tagStr)
                End If
            End If

            If ids.Count > 0 Then
                Dim dataObj As New DataObject()
                dataObj.SetData("EntryIDList", ids)
                dataObj.SetData(DataFormats.Text, String.Join(Environment.NewLine, ids))
                contactInfoTree.DoDragDrop(dataObj, DragDropEffects.Copy)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"contactInfoTree_ItemDrag error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 验证字符串是否是有效的EntryID格式
    ''' </summary>
    Private Function IsValidEntryID(id As String) As Boolean
        If String.IsNullOrEmpty(id) Then Return False
        ' EntryID通常是长的十六进制字符串，长度通常在40-200字符之间
        If id.Length < 40 OrElse id.Length > 200 Then Return False
        ' 检查是否只包含十六进制字符
        For Each c As Char In id
            If Not Char.IsDigit(c) AndAlso Not "ABCDEF".Contains(c.ToString().ToUpper()) Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 解析Outlook的RenPrivateMessages格式数据，提取EntryID
    ''' </summary>
    ''' <param name="data">RenPrivateMessages格式的字节数组</param>
    ''' <returns>解析出的EntryID列表</returns>

    ''' <summary>
    ''' 解析RenPrivateItem格式数据，提取EntryID
    ''' </summary>
    ''' <param name="data">RenPrivateItem格式的二进制数据</param>
    ''' <returns>提取到的EntryID列表</returns>
    Private Function ParseRenPrivateItem(data As Byte()) As List(Of String)
        Dim entryIds As New List(Of String)

        Try
            Debug.WriteLine($"ParseRenPrivateItem: 开始解析 {data.Length} 字节数据")

            If data Is Nothing OrElse data.Length < 4 Then Return entryIds

            ' RenPrivateItem格式可能包含单个邮件项目的信息
            ' 尝试不同的解析策略

            ' 策略1: 查找EntryID模式（十六进制字符串）
            Dim dataStr As String = System.Text.Encoding.ASCII.GetString(data)
            Dim pattern As String = "[0-9A-Fa-f]{40,200}"
            Dim matches As System.Text.RegularExpressions.MatchCollection =
                System.Text.RegularExpressions.Regex.Matches(dataStr, pattern)

            Debug.WriteLine($"ParseRenPrivateItem: 在ASCII字符串中找到 {matches.Count} 个可能的EntryID匹配")

            For Each match As System.Text.RegularExpressions.Match In matches
                Dim possibleEntryId As String = match.Value.ToUpper()
                Debug.WriteLine($"ParseRenPrivateItem: 检查可能的EntryID: {possibleEntryId.Substring(0, Math.Min(50, possibleEntryId.Length))}...")
                If IsValidEntryID(possibleEntryId) Then
                    entryIds.Add(possibleEntryId)
                    Debug.WriteLine($"ParseRenPrivateItem: 添加有效EntryID: {possibleEntryId.Substring(0, Math.Min(50, possibleEntryId.Length))}...")
                End If
            Next

            ' 策略2: 如果ASCII解析失败，尝试直接从二进制数据中查找EntryID模式
            If entryIds.Count = 0 Then
                Debug.WriteLine("ParseRenPrivateItem: ASCII解析失败，尝试二进制模式匹配")

                ' 查找可能的EntryID长度标记（通常EntryID前面有长度信息）
                For i As Integer = 0 To data.Length - 8 Step 4
                    If i + 4 < data.Length Then
                        Dim possibleLength As Integer = BitConverter.ToInt32(data, i)
                        If possibleLength > 40 AndAlso possibleLength < 200 AndAlso i + 4 + possibleLength <= data.Length Then
                            ' 提取可能的EntryID数据
                            Dim entryIdBytes(possibleLength - 1) As Byte
                            Array.Copy(data, i + 4, entryIdBytes, 0, possibleLength)
                            Dim entryId As String = BitConverter.ToString(entryIdBytes).Replace("-", "")

                            Debug.WriteLine($"ParseRenPrivateItem: 二进制模式找到可能的EntryID，长度={possibleLength}: {entryId.Substring(0, Math.Min(50, entryId.Length))}...")

                            If IsValidEntryID(entryId) Then
                                entryIds.Add(entryId)
                                Debug.WriteLine($"ParseRenPrivateItem: 添加有效EntryID: {entryId.Substring(0, Math.Min(50, entryId.Length))}...")
                                Exit For ' 找到一个就够了
                            End If
                        End If
                    End If
                Next
            End If

            Debug.WriteLine($"ParseRenPrivateItem: 最终解析出 {entryIds.Count} 个有效EntryID")

        Catch ex As System.Exception
            Debug.WriteLine($"ParseRenPrivateItem异常: {ex.Message}")
            Debug.WriteLine($"ParseRenPrivateItem异常堆栈: {ex.StackTrace}")
        End Try

        Return entryIds
    End Function

    Private Function ParseRenPrivateMessages(data As Byte()) As List(Of String)
        Dim entryIds As New List(Of String)

        Try
            Debug.WriteLine($"ParseRenPrivateMessages: 开始解析 {data.Length} 字节数据")

            If data Is Nothing OrElse data.Length < 8 Then Return entryIds

            ' 显示更多的数据预览
            Dim hexPreview As String = String.Join(" ", data.Take(Math.Min(64, data.Length)).Select(Function(b) b.ToString("X2")))
            Debug.WriteLine($"ParseRenPrivateMessages: 数据十六进制预览(前64字节): {hexPreview}")

            ' RenPrivateMessages格式通常包含多个邮件项目
            ' 格式可能是: [邮件数量][邮件1数据][邮件2数据]...

            ' 策略1: 尝试读取邮件数量（前4字节）
            Dim mailCount As Integer = BitConverter.ToInt32(data, 0)
            Debug.WriteLine($"ParseRenPrivateMessages: 读取到邮件数量: {mailCount}")

            ' 也尝试其他位置的邮件数量
            If data.Length >= 8 Then
                Dim mailCount2 As Integer = BitConverter.ToInt32(data, 4)
                Debug.WriteLine($"ParseRenPrivateMessages: 偏移4字节处的数量: {mailCount2}")
            End If

            If mailCount > 0 AndAlso mailCount < 1000 Then ' 合理的邮件数量范围
                Dim offset As Integer = 4
                For i As Integer = 0 To mailCount - 1
                    If offset + 4 >= data.Length Then
                        Debug.WriteLine($"ParseRenPrivateMessages: 偏移{offset}超出数据范围，退出")
                        Exit For
                    End If

                    ' 读取当前邮件数据长度
                    Dim mailDataLength As Integer = BitConverter.ToInt32(data, offset)
                    offset += 4

                    Debug.WriteLine($"ParseRenPrivateMessages: 邮件{i + 1}数据长度: {mailDataLength}, 偏移: {offset}")

                    If mailDataLength > 0 AndAlso mailDataLength < data.Length AndAlso offset + mailDataLength <= data.Length Then
                        ' 提取邮件数据
                        Dim mailData(mailDataLength - 1) As Byte
                        Array.Copy(data, offset, mailData, 0, mailDataLength)

                        ' 显示邮件数据预览
                        Dim mailHexPreview As String = String.Join(" ", mailData.Take(Math.Min(32, mailData.Length)).Select(Function(b) b.ToString("X2")))
                        Debug.WriteLine($"ParseRenPrivateMessages: 邮件{i + 1}数据预览: {mailHexPreview}")

                        ' 使用现有的ParseRenPrivateItem函数解析单个邮件
                        Dim mailEntryIds As List(Of String) = ParseRenPrivateItem(mailData)
                        entryIds.AddRange(mailEntryIds)
                        Debug.WriteLine($"ParseRenPrivateMessages: 邮件{i + 1}解析出{mailEntryIds.Count}个EntryID")

                        offset += mailDataLength
                    Else
                        Debug.WriteLine($"ParseRenPrivateMessages: 邮件{i + 1}数据长度无效({mailDataLength})，偏移{offset}，总长度{data.Length}，跳过")
                        Exit For
                    End If
                Next
            Else
                Debug.WriteLine($"ParseRenPrivateMessages: 邮件数量无效({mailCount})，尝试其他解析策略")
            End If

            ' 策略2: 如果按邮件数量解析失败，尝试全局搜索EntryID模式
            If entryIds.Count = 0 Then
                Debug.WriteLine("ParseRenPrivateMessages: 按邮件数量解析失败，尝试全局搜索")

                ' 在整个数据中搜索可能的EntryID
                Try
                    Dim dataStr As String = System.Text.Encoding.ASCII.GetString(data)
                    Debug.WriteLine($"ParseRenPrivateMessages: ASCII字符串长度: {dataStr.Length}")
                    Debug.WriteLine($"ParseRenPrivateMessages: ASCII字符串预览: {dataStr.Substring(0, Math.Min(100, dataStr.Length)).Replace(vbNullChar, ".")}")

                    Dim pattern As String = "[0-9A-Fa-f]{40,200}"
                    Dim matches As System.Text.RegularExpressions.MatchCollection =
                        System.Text.RegularExpressions.Regex.Matches(dataStr, pattern)

                    Debug.WriteLine($"ParseRenPrivateMessages: 全局搜索找到 {matches.Count} 个可能的EntryID")

                    For Each match As System.Text.RegularExpressions.Match In matches
                        Dim possibleEntryId As String = match.Value.ToUpper()
                        Debug.WriteLine($"ParseRenPrivateMessages: 检查可能的EntryID: {possibleEntryId.Substring(0, Math.Min(50, possibleEntryId.Length))}...")
                        If IsValidEntryID(possibleEntryId) Then
                            entryIds.Add(possibleEntryId)
                            Debug.WriteLine($"ParseRenPrivateMessages: 添加有效EntryID: {possibleEntryId.Substring(0, Math.Min(50, possibleEntryId.Length))}...")
                        End If
                    Next
                Catch ex As System.Exception
                    Debug.WriteLine($"ParseRenPrivateMessages: ASCII搜索异常: {ex.Message}")
                End Try
            End If

            ' 策略3: 尝试二进制搜索
            If entryIds.Count = 0 Then
                Debug.WriteLine("ParseRenPrivateMessages: 全局搜索失败，尝试二进制搜索")

                Dim foundCount As Integer = 0
                For i As Integer = 0 To data.Length - 8 Step 1
                    If i + 4 < data.Length Then
                        Dim possibleLength As Integer = BitConverter.ToInt32(data, i)
                        If possibleLength >= 40 AndAlso possibleLength <= 200 AndAlso i + 4 + possibleLength <= data.Length Then
                            Debug.WriteLine($"ParseRenPrivateMessages: 在偏移{i}找到可能的长度标记: {possibleLength}")

                            ' 提取可能的EntryID数据
                            Dim entryIdBytes(possibleLength - 1) As Byte
                            Array.Copy(data, i + 4, entryIdBytes, 0, possibleLength)

                            ' 显示数据预览
                            Dim bytesPreview As String = String.Join(" ", entryIdBytes.Take(Math.Min(16, entryIdBytes.Length)).Select(Function(b) b.ToString("X2")))
                            Debug.WriteLine($"ParseRenPrivateMessages: 可能的EntryID数据预览: {bytesPreview}")

                            ' 检查是否看起来像EntryID（大部分是可打印字符或十六进制）
                            Dim isValidBytes As Boolean = True
                            Dim nullCount As Integer = 0
                            For j As Integer = 0 To entryIdBytes.Length - 1
                                If entryIdBytes(j) = 0 Then
                                    nullCount += 1
                                    If nullCount > entryIdBytes.Length \ 4 Then ' 如果超过1/4是null字节，认为无效
                                        isValidBytes = False
                                        Exit For
                                    End If
                                End If
                            Next

                            Debug.WriteLine($"ParseRenPrivateMessages: 数据有效性检查: {isValidBytes}, null字节数: {nullCount}")

                            If isValidBytes Then
                                Dim entryId As String = BitConverter.ToString(entryIdBytes).Replace("-", "")
                                Debug.WriteLine($"ParseRenPrivateMessages: 生成的EntryID: {entryId.Substring(0, Math.Min(50, entryId.Length))}...")
                                If IsValidEntryID(entryId) Then
                                    entryIds.Add(entryId)
                                    Debug.WriteLine($"ParseRenPrivateMessages: 二进制搜索找到EntryID: {entryId.Substring(0, Math.Min(50, entryId.Length))}...")
                                    i += possibleLength + 4 ' 跳过已处理的数据
                                    foundCount += 1
                                    If foundCount >= 10 Then Exit For ' 限制搜索数量
                                End If
                            End If
                        End If
                    End If
                Next
                Debug.WriteLine($"ParseRenPrivateMessages: 二进制搜索完成，检查了{foundCount}个可能的位置")
            End If

            Debug.WriteLine($"ParseRenPrivateMessages: 最终解析出 {entryIds.Count} 个有效EntryID")

        Catch ex As System.Exception
            Debug.WriteLine($"ParseRenPrivateMessages异常: {ex.Message}")
            Debug.WriteLine($"ParseRenPrivateMessages异常堆栈: {ex.StackTrace}")
        End Try

        Return entryIds
    End Function

    Private Sub lvMails_ItemDrag(sender As Object, e As ItemDragEventArgs)
        Try
            If lvMails Is Nothing OrElse lvMails.SelectedItems.Count = 0 Then Return
            Debug.WriteLine($"lvMails_ItemDrag: 开始处理拖拽，选中项数量: {lvMails.SelectedItems.Count}")

            Dim ids As New List(Of String)
            For Each it As ListViewItem In lvMails.SelectedItems
                Debug.WriteLine($"  项目文本: {it.Text}")
                Debug.WriteLine($"  Tag类型: {If(it.Tag Is Nothing, "Nothing", it.Tag.GetType().Name)}")
                Debug.WriteLine($"  Tag内容: {If(it.Tag Is Nothing, "Nothing", it.Tag.ToString())}")

                Dim entryId As String = TryCast(it.Tag, String)
                Debug.WriteLine($"  转换后EntryID: {If(String.IsNullOrEmpty(entryId), "空或无效", "有效")}")
                If Not String.IsNullOrEmpty(entryId) Then
                    ids.Add(entryId)
                    Debug.WriteLine($"  添加到拖拽列表: {entryId.Substring(0, Math.Min(50, entryId.Length))}...")
                End If
            Next

            Debug.WriteLine($"lvMails_ItemDrag: 有效EntryID数量: {ids.Count}")
            If ids.Count > 0 Then
                Dim dataObj As New DataObject()
                ' 使用标准格式传递EntryID列表
                dataObj.SetData("EntryIDList", ids)
                dataObj.SetData(DataFormats.StringFormat, String.Join(Environment.NewLine, ids))
                dataObj.SetData(DataFormats.Text, String.Join(Environment.NewLine, ids))
                Debug.WriteLine($"lvMails_ItemDrag: 设置拖拽数据完成")
                Debug.WriteLine($"  EntryIDList: {ids.Count}个ID")
                Debug.WriteLine($"  StringFormat: {String.Join(Environment.NewLine, ids).Length}字符")
                Debug.WriteLine($"  Text: {String.Join(Environment.NewLine, ids).Length}字符")
                lvMails.DoDragDrop(dataObj, DragDropEffects.Copy)
            Else
                Debug.WriteLine("lvMails_ItemDrag: 没有有效的EntryID，取消拖拽")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_ItemDrag error: {ex.Message}")
        End Try
    End Sub

    Private Sub lvMails_DragEnter(sender As Object, e As DragEventArgs)
        Try
            If e.Data IsNot Nothing Then
                ' 检查是否包含任何我们支持的拖拽格式
                Dim supportedFormats() As String = {
                    "EntryIDList", "RenPrivateItem", "RenPrivateMessages",
                    "Outlook.MailItem", "Outlook.Item", "CF_OUTLOOK",
                    DataFormats.StringFormat, DataFormats.Text
                }

                Dim hasValidFormat As Boolean = False
                For Each format As String In supportedFormats
                    If e.Data.GetDataPresent(format) Then
                        hasValidFormat = True
                        Debug.WriteLine($"DragEnter: 检测到支持的格式 '{format}'")
                        Exit For
                    End If
                Next

                If hasValidFormat Then
                    e.Effect = DragDropEffects.Copy
                    Debug.WriteLine("DragEnter: 允许拖拽操作")
                Else
                    e.Effect = DragDropEffects.None
                    Debug.WriteLine("DragEnter: 不支持的拖拽格式")
                End If
            Else
                e.Effect = DragDropEffects.None
                Debug.WriteLine("DragEnter: 拖拽数据为空")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_DragEnter error: {ex.Message}")
        End Try
    End Sub

    Private Async Sub lvMails_DragDrop(sender As Object, e As DragEventArgs)
        Try
            Debug.WriteLine("=== 简化拖拽处理开始 ===")

            ' 简化方案：拖拽完成后，直接从当前激活的邮件获取会话ID
            ' 然后提示用户选择要合并的邮件

            ' 1. 获取当前激活邮件的会话ID
            Dim targetConversationId As String = GetCurrentActiveMailConversationId()
            Debug.WriteLine($"当前激活邮件的会话ID: '{targetConversationId}'")

            If String.IsNullOrEmpty(targetConversationId) Then
                MessageBox.Show("无法获取当前邮件的会话ID，请确保已选中一封邮件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 2. 从拖拽数据中获取被拖拽邮件的EntryID
            Dim draggedMailIds As New List(Of String)

            ' 尝试从EntryIDList格式获取拖拽的邮件ID
            If e.Data.GetDataPresent("EntryIDList") Then
                Try
                    Dim entryIdData As Byte() = CType(e.Data.GetData("EntryIDList"), Byte())
                    If entryIdData IsNot Nothing AndAlso entryIdData.Length > 0 Then
                        Debug.WriteLine($"获取到EntryIDList数据，长度: {entryIdData.Length}")

                        ' 解析EntryIDList格式的数据
                        ' EntryIDList格式：前4字节是条目数，然后每个条目包含长度和EntryID
                        Dim offset As Integer = 0
                        If entryIdData.Length >= 4 Then
                            Dim entryCount As Integer = BitConverter.ToInt32(entryIdData, offset)
                            offset += 4
                            Debug.WriteLine($"EntryIDList包含 {entryCount} 个条目")

                            For i As Integer = 0 To entryCount - 1
                                If offset + 4 <= entryIdData.Length Then
                                    Dim entryIdLength As Integer = BitConverter.ToInt32(entryIdData, offset)
                                    offset += 4

                                    If offset + entryIdLength <= entryIdData.Length Then
                                        Dim entryIdBytes As Byte() = New Byte(entryIdLength - 1) {}
                                        Array.Copy(entryIdData, offset, entryIdBytes, 0, entryIdLength)
                                        offset += entryIdLength

                                        ' 将字节数组转换为十六进制字符串作为EntryID
                                        Dim entryId As String = BitConverter.ToString(entryIdBytes).Replace("-", "")
                                        draggedMailIds.Add(entryId)
                                        Debug.WriteLine($"解析到EntryID: {entryId}")
                                    End If
                                End If
                            Next
                        End If
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"解析EntryIDList数据失败: {ex.Message}")
                End Try
            End If

            ' 如果无法从拖拽数据获取，则提示用户确认并使用当前选中的邮件
            If draggedMailIds.Count = 0 Then
                Debug.WriteLine($"从拖拽数据中未获取到邮件ID，当前选中邮件数量: {lvMails.SelectedItems.Count}")

                Dim result As DialogResult = MessageBox.Show(
                    $"检测到拖拽操作。{Environment.NewLine}{Environment.NewLine}" &
                    $"是否要将当前选中的邮件合并到目标会话？{Environment.NewLine}" &
                    $"目标会话ID: {targetConversationId.Substring(0, Math.Min(20, targetConversationId.Length))}...",
                    "邮件合并确认",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question)

                If result = DialogResult.No Then
                    Debug.WriteLine("用户取消了邮件合并操作")
                    Return
                End If

                ' 使用当前ListView中选中的邮件
                Debug.WriteLine("开始从ListView获取选中邮件的EntryID...")
                For Each item As ListViewItem In lvMails.SelectedItems
                    If item.Tag IsNot Nothing Then
                        Dim entryId As String = ConvertEntryIDToString(item.Tag)
                        draggedMailIds.Add(entryId)
                        Debug.WriteLine($"添加邮件ID: {entryId}")
                    Else
                        Debug.WriteLine("发现一个没有Tag的ListView项目")
                    End If
                Next
                Debug.WriteLine($"从ListView获取到 {draggedMailIds.Count} 个邮件ID")
            End If

            If draggedMailIds.Count = 0 Then
                MessageBox.Show("没有找到要合并的邮件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Debug.WriteLine($"准备合并 {draggedMailIds.Count} 封邮件")

            ' 4. 执行邮件合并操作
            ShowProgress($"正在合并邮件到会话...")

            Dim success As Integer = 0
            Dim totalMails As Integer = 0

            Debug.WriteLine($"开始合并会话，目标ID: {targetConversationId}，被拖拽邮件数: {draggedMailIds.Count}")

            ' 获取所有需要处理的会话ID和对应的邮件
            Dim sourceConversationIds As New HashSet(Of String)
            Dim conversationMailsMap As New Dictionary(Of String, List(Of String))

            Await Task.Run(Sub()
                               ' 第一步：获取被拖拽邮件的原始会话ID，并收集每个会话的所有邮件
                               For Each mailId As String In draggedMailIds
                                   Try
                                       Dim mailItem = Globals.ThisAddIn.Application.Session.GetItemFromID(mailId)
                                       If mailItem IsNot Nothing Then
                                           Dim originalConversationId As String = GetSafeConversationID(mailItem)
                                           Debug.WriteLine($"邮件 {mailId} 的原始会话ID: {originalConversationId}, 目标会话ID: {targetConversationId}")

                                           If Not String.IsNullOrEmpty(originalConversationId) AndAlso originalConversationId <> targetConversationId Then
                                               sourceConversationIds.Add(originalConversationId)
                                               Debug.WriteLine($"发现需要处理的源会话ID: {originalConversationId}")

                                               ' 如果还没有收集过这个会话的邮件，则收集
                                               If Not conversationMailsMap.ContainsKey(originalConversationId) Then
                                                   Dim conversationMails As New List(Of String)

                                                   ' 使用第一个邮件来获取整个会话的邮件列表
                                                   Try
                                                       Dim allMailsInConversation = GetAllMailsInConversation(mailId)
                                                       conversationMails.AddRange(allMailsInConversation)
                                                       conversationMailsMap(originalConversationId) = conversationMails
                                                       Debug.WriteLine($"会话 {originalConversationId} 包含 {conversationMails.Count} 封邮件")
                                                   Catch ex As System.Exception
                                                       Debug.WriteLine($"获取会话 {originalConversationId} 的邮件列表时出错: {ex.Message}")
                                                       conversationMailsMap(originalConversationId) = New List(Of String)
                                                   End Try
                                               End If
                                           Else
                                               If String.IsNullOrEmpty(originalConversationId) Then
                                                   Debug.WriteLine($"邮件 {mailId} 没有会话ID")
                                               Else
                                                   Debug.WriteLine($"邮件 {mailId} 已经在目标会话中，无需处理")
                                               End If
                                           End If
                                           Marshal.ReleaseComObject(mailItem)
                                       End If
                                   Catch ex As System.Exception
                                       Debug.WriteLine($"获取邮件 {mailId} 的会话ID时出错: {ex.Message}")
                                   End Try
                               Next

                               ' 计算总邮件数
                               For Each kvp In conversationMailsMap
                                   totalMails += kvp.Value.Count
                               Next

                               ' 第二步：对每个源会话的所有邮件更新为目标会话ID
                               For Each kvp In conversationMailsMap
                                   Dim sourceConversationId As String = kvp.Key
                                   Dim conversationMailIds As List(Of String) = kvp.Value

                                   Try
                                       Debug.WriteLine($"开始处理源会话: {sourceConversationId}，包含 {conversationMailIds.Count} 封邮件")

                                       ' 更新该会话的所有邮件
                                       For Each mailEntryId As String In conversationMailIds
                                           Try
                                               Debug.WriteLine($"正在为邮件 {mailEntryId} 设置自定义会话ID为: {targetConversationId}")
                                               If SetCustomConversationIdByEntryID(mailEntryId, targetConversationId) Then
                                                   success += 1
                                                   Debug.WriteLine($"邮件 {mailEntryId} 自定义会话ID设置成功")
                                               Else
                                                   Debug.WriteLine($"邮件 {mailEntryId} 自定义会话ID设置失败")
                                               End If
                                           Catch ex As System.Exception
                                               Debug.WriteLine($"设置邮件 {mailEntryId} 的自定义会话ID时出错: {ex.Message}")
                                           End Try

                                           ' 进度更新
                                           Me.BeginInvoke(Sub() UpdateProgress($"已处理 {success}/{totalMails}"))
                                       Next

                                   Catch ex As System.Exception
                                       Debug.WriteLine($"处理源会话 {sourceConversationId} 时出错: {ex.Message}")
                                   End Try
                               Next
                           End Sub)

            HideProgress()

            ' 清除相关会话的缓存，确保刷新时重新加载数据
            Try
                SyncLock conversationMailsCache
                    ' 清除目标会话的缓存
                    If conversationMailsCache.ContainsKey(targetConversationId) Then
                        conversationMailsCache.Remove(targetConversationId)
                        Debug.WriteLine($"已清除目标会话缓存: {targetConversationId}")
                    End If

                    ' 清除所有源会话的缓存
                    For Each sourceConversationId As String In sourceConversationIds
                        If conversationMailsCache.ContainsKey(sourceConversationId) Then
                            conversationMailsCache.Remove(sourceConversationId)
                            Debug.WriteLine($"已清除源会话缓存: {sourceConversationId}")
                        End If
                    Next
                End SyncLock
            Catch ex As System.Exception
                Debug.WriteLine($"清除会话缓存时出错: {ex.Message}")
            End Try

            ' 强制刷新当前列表与右侧内容
            ' 由于邮件已合并到目标会话，需要强制重新加载列表来显示新的邮件
            Debug.WriteLine("拖拽完成，强制刷新会话列表以显示合并后的邮件")

            ' 临时清空当前会话ID，强制UpdateMailList重新加载
            Dim tempConversationId As String = currentConversationId
            currentConversationId = String.Empty

            ' 强制刷新列表
            UpdateMailList(tempConversationId, currentMailEntryID)

            ' 恢复会话ID
            currentConversationId = tempConversationId

            Dim message As String = $"邮件合并完成：成功 {success}/{totalMails}" & Environment.NewLine &
                                   $"已将选中的 {totalMails} 封邮件合并到当前会话"

            MessageBox.Show(message, "合并结果", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_DragDrop error: {ex.Message}")
            HideProgress()
            Globals.ThisAddIn.ShowErrorWithConfig("拖拽处理时出错", ex.Message)
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
                Dim success = OutlookMyList.Utils.OutlookUtils.FastOpenMailItem(entryID)

                If success Then
                    Debug.WriteLine("快速邮件打开成功")
                Else
                    Debug.WriteLine("快速邮件打开失败，尝试兜底方法")

                    ' 兜底：使用原有方法
                    Dim mailItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryID)
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
                            OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
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

        Catch ex As System.Runtime.InteropServices.COMException
            ' 记录调试信息
            If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                Debug.WriteLine($"SafeOpenOutlookMail COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
            End If
            ' SafeOpenOutlookMail 方法设计为绝不抛出异常，所以不显示错误对话框
        Catch ex As System.Exception
            ' 记录调试信息
            If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                Debug.WriteLine($"SafeOpenOutlookMail 异常: {ex.Message}")
            End If
            ' SafeOpenOutlookMail 方法设计为绝不抛出异常，所以不显示错误对话框
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
                                  DirectCast(outputTextBox, TextBox).Text = "正在收集联系人来往邮件信息..."
                              ElseIf TypeOf outputTextBox Is RichTextBox Then
                                  DirectCast(outputTextBox, RichTextBox).Text = "正在收集联系人来往邮件信息..."
                              End If
                          End Sub)
            Else
                If TypeOf outputTextBox Is TextBox Then
                    DirectCast(outputTextBox, TextBox).Text = "正在收集联系人来往邮件信息..."
                ElseIf TypeOf outputTextBox Is RichTextBox Then
                    DirectCast(outputTextBox, RichTextBox).Text = "正在收集联系人来往邮件信息..."
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

                Dim currentItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                Dim conversation As Outlook.Conversation = Nothing

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                    conversation = DirectCast(currentItem, Outlook.MeetingItem).GetConversation()
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
                            item = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryIdStr)

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
                            Dim item As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
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
            If ErrorNotificationSettings.Instance.ShowErrorDialogs Then
                MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End Try
    End Sub

    Private Async Sub btnNewNote_Click(sender As Object, e As EventArgs)
        Try
            ' 在后台线程中获取邮件主题，避免阻塞UI
            Dim subject As String = Await Task.Run(Function()
                                                       Try
                                                           Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
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
        AddHandler lvMails.MouseDown, AddressOf lvMails_MouseDown
        AddHandler lvMails.ColumnClick, AddressOf lvMails_ColumnClick
        AddHandler lvMails.DoubleClick, AddressOf lvMails_DoubleClick
        If mailBrowser IsNot Nothing Then
            AddHandler mailBrowser.Navigating, AddressOf MailBrowser_Navigating
        End If
        eventsBound = True
        Debug.WriteLine("BindEvents: 事件绑定完成")
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
                    If Not OutlookMyList.Utils.OutlookUtils.FastOpenMailItem(entityId, storeId) Then
                        ' 兜底：仍然交给系统处理
                        OutlookMyList.Handlers.MailHandler.OpenLink(urlStr)
                    End If
                Else
                    ' 未能解析 entityID，回退到系统打开
                    OutlookMyList.Handlers.MailHandler.OpenLink(urlStr)
                End If
            Else
                ' 普通 http/https 等链接，走系统默认浏览器
                OutlookMyList.Handlers.MailHandler.OpenLink(urlStr)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"MailBrowser_Navigating error: {ex.Message}")
        End Try
    End Sub

    Private Sub MailBrowser_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs)
        Try
            Dim backHexLog As String = $"#{currentBackColor.R:X2}{currentBackColor.G:X2}{currentBackColor.B:X2}"
            Dim foreHexLog As String = $"#{currentForeColor.R:X2}{currentForeColor.G:X2}{currentForeColor.B:X2}"
            Debug.WriteLine($"[ThemeDebug] Navigated: url={If(e.Url, Nothing)}, readyState={mailBrowser.ReadyState}, isDisplayingMailContent={isDisplayingMailContent}, back={backHexLog}, fore={foreHexLog}")
        Catch ex As System.Exception
            Debug.WriteLine($"MailBrowser_Navigated error: {ex.Message}")
        End Try
    End Sub

    Private Sub MailBrowser_ProgressChanged(sender As Object, e As WebBrowserProgressChangedEventArgs)
        Try
            Debug.WriteLine($"[ThemeDebug] ProgressChanged: current={e.CurrentProgress}, maximum={e.MaximumProgress}, readyState={mailBrowser.ReadyState}")
        Catch ex As System.Exception
            Debug.WriteLine($"MailBrowser_ProgressChanged error: {ex.Message}")
        End Try
    End Sub

    ' WebBrowser文档加载完成事件处理
    Private Sub WebBrowser_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        Try
            Debug.WriteLine($"WebBrowser_DocumentCompleted 触发，isDisplayingMailContent = {isDisplayingMailContent}")

            If mailBrowser.ReadyState <> WebBrowserReadyState.Complete Then
                Debug.WriteLine($"[ThemeDebug] DocumentCompleted: readyState={mailBrowser.ReadyState}")
                Return
            End If

            ' 只有在显示邮件内容时才应用主题样式
            If isDisplayingMailContent AndAlso mailBrowser.Document IsNot Nothing Then
                Debug.WriteLine("开始在DocumentCompleted中应用主题样式")
                Debug.WriteLine($"[ThemeDebug] DocumentCompleted: url={e.Url}")
                Dim doc = mailBrowser.Document

                ' 使用当前保存的主题颜色
                Dim bgColor As Color = currentBackColor
                Dim fgColor As Color = currentForeColor
                Dim accentColor As Color = Color.FromArgb(0, 120, 215) ' 默认蓝色强调色
                Dim bgColorHex As String = $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}"
                Dim fgColorHex As String = $"#{fgColor.R:X2}{fgColor.G:X2}{fgColor.B:X2}"
                Dim accentColorHex As String = $"#{accentColor.R:X2}{accentColor.G:X2}{accentColor.B:X2}"
                Debug.WriteLine($"[ThemeDebug] DocumentCompleted: back={bgColorHex}, fore={fgColorHex}")

                ' 通过JavaScript强制应用样式，彻底覆盖所有可能的颜色设置
                Dim script As String = $"
                    (function() {{
                        var elements = document.getElementsByTagName('*');
                        for (var i = 0; i < elements.length; i++) {{
                            var elem = elements[i];
                            var tagName = elem.tagName.toUpperCase();
                            
                            if (tagName !== 'STYLE' && tagName !== 'SCRIPT') {{
                                // 使用setProperty方法强制覆盖，包括内联样式
                                elem.style.setProperty('background-color', '{bgColorHex}', 'important');
                                elem.style.setProperty('color', '{fgColorHex}', 'important');
                                elem.style.removeProperty('background-image');
                                elem.style.removeProperty('background');
                                
                                // 特殊处理标题和强调元素
                                if (tagName === 'H1' || tagName === 'H2' || tagName === 'H3' || 
                                    tagName === 'H4' || tagName === 'H5' || tagName === 'H6' ||
                                    tagName === 'STRONG' || tagName === 'B' || tagName === 'A') {{
                                    elem.style.setProperty('color', '{accentColorHex}', 'important');
                                }}
                                
                                // 对于非body元素，设置透明背景
                                if (tagName !== 'BODY') {{
                                    elem.style.setProperty('background-color', 'transparent', 'important');
                                }}
                                
                                // 移除可能存在的内联颜色属性
                                elem.removeAttribute('color');
                                elem.removeAttribute('bgcolor');
                            }}
                        }}
                        
                        // 确保body元素的样式
                        document.body.style.setProperty('background-color', '{bgColorHex}', 'important');
                        document.body.style.setProperty('color', '{fgColorHex}', 'important');
                        try {{
                            var iframes = document.getElementsByTagName('iframe');
                            for (var j = 0; j < iframes.length; j++) {{
                                var idoc = null; try {{ idoc = iframes[j].contentDocument || (iframes[j].contentWindow && iframes[j].contentWindow.document); }} catch(_) {{}}
                                if (idoc && idoc.body) {{
                                    var es = idoc.getElementsByTagName('*');
                                    for (var k = 0; k < es.length; k++) {{
                                        var ee = es[k]; var tt = ee.tagName.toUpperCase();
                                        if (tt !== 'STYLE' && tt !== 'SCRIPT') {{
                                            ee.style.setProperty('background-color', '{bgColorHex}', 'important');
                                            ee.style.setProperty('color', '{fgColorHex}', 'important');
                                            if (tt !== 'BODY') {{ ee.style.setProperty('background-color', 'transparent', 'important'); }}
                                            ee.style.removeProperty('background-image');
                                            ee.style.removeProperty('background');
                                            ee.removeAttribute('color'); ee.removeAttribute('bgcolor');
                                        }}
                                    }}
                                    idoc.body.style.setProperty('background-color', '{bgColorHex}', 'important');
                                    idoc.body.style.setProperty('color', '{fgColorHex}', 'important');
                                }}
                            }}
                        }} catch(_) {{}}
                        
                        // 处理所有文本节点的父元素
                        var walker = document.createTreeWalker(
                            document.body,
                            NodeFilter.SHOW_TEXT,
                            null,
                            false
                        );
                        var textNode;
                        while (textNode = walker.nextNode()) {{
                            if (textNode.parentElement) {{
                                textNode.parentElement.style.setProperty('color', '{fgColorHex}', 'important');
                            }}
                        }}
                    }})();
                "

                Dim applyResult As Object = doc.InvokeScript("eval", New Object() {script})
                Debug.WriteLine("DocumentCompleted中主题样式应用完成")
                If Not IsNothing(applyResult) Then
                    Debug.WriteLine($"[ThemeDebug] DocumentCompleted: applySummary={applyResult.ToString()}")
                End If
                Dim dcBg As Object = doc.InvokeScript("eval", New Object() {"(function(){try{var el=document.body;if(!el) return 'no-body';var c=(window.getComputedStyle?window.getComputedStyle(el):el.currentStyle);var val=(c? (c.backgroundColor||'') : ''); if(!val||val==='transparent'){val=(el.style?el.style.backgroundColor:'')||val;} return val||'';}catch(e){return 'err:'+e.message;}})();"})
                Dim dcHtmlBg As Object = doc.InvokeScript("eval", New Object() {"(function(){try{var el=document.documentElement;if(!el) return 'no-html';var c=(window.getComputedStyle?window.getComputedStyle(el):el.currentStyle);var val=(c? (c.backgroundColor||'') : ''); if(!val||val==='transparent'){val=(el.style?el.style.backgroundColor:'')||val;} return val||'';}catch(e){return 'err:'+e.message;}})();"})
                Dim dcTextLen As Object = doc.InvokeScript("eval", New Object() {"(function(){try{var b=document.body;return b&&b.innerText?b.innerText.length:0;}catch(e){return -1;}})();"})
                Debug.WriteLine($"[ThemeDebug] DocumentCompleted: computedBodyBg={If(dcBg, Nothing)}, htmlBg={If(dcHtmlBg, Nothing)}, innerTextLen={If(dcTextLen, Nothing)}")
                'Dim computedJs As String = "(function(){try{var s={url:document.location.href,bodyBg:getComputedStyle(document.body).backgroundColor,innerTextLen:(document.body&&document.body.innerText?document.body.innerText.length:0),elements:document.getElementsByTagName('*').length};return JSON.stringify(s);}catch(e){return 'err:'+e.message;}})();"
                'Dim computed As Object = doc.InvokeScript("eval", New Object() {computedJs})
                'If computed IsNot Nothing Then
                '    Debug.WriteLine($"[ThemeDebug] DocumentCompleted: computed={computed.ToString()}")
                'End If
                RemoveHandler mailBrowser.Navigated, AddressOf MailBrowser_Navigated
                RemoveHandler mailBrowser.ProgressChanged, AddressOf MailBrowser_ProgressChanged
                RemoveHandler mailBrowser.DocumentCompleted, AddressOf WebBrowser_DocumentCompleted
            Else
                Debug.WriteLine("跳过DocumentCompleted中的主题应用 - 不是邮件内容或Document为空")
            End If


        Catch ex As System.Exception
            Debug.WriteLine($"WebBrowser_DocumentCompleted error: {ex.Message}")
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
                        ' 使用延迟机制确保主题正确设置
                        Dim delayTimer As New System.Windows.Forms.Timer()
                        delayTimer.Interval = 100
                        AddHandler delayTimer.Tick, Sub(sender, e)
                                                        delayTimer.Stop()
                                                        delayTimer.Dispose()
                                                        Try
                                                            mailBrowser.DocumentText = GetThemedDefaultContent()
                                                            
                                                            ' 立即应用当前主题到WebBrowser控件
                                                            Dim currentBgColor As Color = ColorTranslator.FromHtml(globalThemeBackgroundColor)
                                                            Dim currentFgColor As Color = ColorTranslator.FromHtml(globalThemeForegroundColor)
                                                            UpdateWebBrowserTheme(currentBgColor, currentFgColor)
                                                        Catch ex As System.Exception
                                                            Debug.WriteLine($"延迟设置默认内容失败: {ex.Message}")
                                                        End Try
                                                    End Sub
                        delayTimer.Start()
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

            ' 使用传入的conversationId作为有效会话ID（已禁用智能会话ID）
            Dim effectiveConversationId As String = conversationId

            ' 检查是否需要重新加载列表
            Dim needReload As Boolean = True
            If lvMails.Items.Count > 0 AndAlso Not String.IsNullOrEmpty(effectiveConversationId) AndAlso
               String.Equals(effectiveConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                needReload = False
            End If

            ' 单独处理无会话的邮件（无智能会话ID）
            If Not String.IsNullOrEmpty(mailEntryID) AndAlso String.IsNullOrEmpty(effectiveConversationId) Then
                Debug.WriteLine($"处理无会话邮件，强制重新加载({mailEntryID})")

                ' 异步加载列表（将当前单封邮件加入列表）
                Await LoadConversationMailsAsync(mailEntryID)

                ' 加载完成后再设置currentMailEntryID
                currentMailEntryID = mailEntryID

                ' 自动加载 WebView 内容
                If Me.IsHandleCreated Then
                    Me.BeginInvoke(Sub() LoadMailContentDeferred(mailEntryID))
                    Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
                    UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)
                End If

                Debug.WriteLine($"处理无会话邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Return
            End If

            If needReload Then
                ' 异步加载会话邮件，完全不阻塞主窗口
                Await LoadConversationMailsAsync(mailEntryID)
                currentMailEntryID = mailEntryID

                ' 更新当前会话ID并检查笔记
                If Not String.Equals(effectiveConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                    currentConversationId = effectiveConversationId
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

            ' 检查是否需要重新加载列表（使用智能会话ID）
            Dim needReload As Boolean = True
            Dim effectiveConversationId As String = conversationId
            Try
                If Not String.IsNullOrEmpty(mailEntryID) Then
                    Dim item = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mailEntryID)
                    If item IsNot Nothing Then
                        Dim smartKey = GetSmartConversationKey(item)
                        If Not String.IsNullOrEmpty(smartKey) Then
                            effectiveConversationId = smartKey
                        End If
                    End If
                End If
            Catch
            End Try
            If lvMails.Items.Count > 0 AndAlso Not String.IsNullOrEmpty(effectiveConversationId) AndAlso
               String.Equals(effectiveConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
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
                If Not String.Equals(effectiveConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                    currentConversationId = effectiveConversationId
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

                    AddItemWithTheme(lvMails, itemCopy)
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

            ' 更新会话信息按钮标题
            UpdateConversationInfoButtonTitle()
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

                    ' 检查缓存（遵循全局缓存开关）
                    If IsCacheEnabled() AndAlso Not String.IsNullOrEmpty(entryID) Then
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

                    ' 将结果存入缓存（遵循全局缓存开关）
                    If IsCacheEnabled() AndAlso props.IsValid AndAlso Not String.IsNullOrEmpty(props.EntryID) Then
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
                Dim quickItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If quickItem IsNot Nothing Then
                    If TypeOf quickItem Is Outlook.MailItem Then
                        quickConversationId = DirectCast(quickItem, Outlook.MailItem).ConversationID
                    ElseIf TypeOf quickItem Is Outlook.AppointmentItem Then
                        quickConversationId = DirectCast(quickItem, Outlook.AppointmentItem).ConversationID
                    ElseIf TypeOf quickItem Is Outlook.MeetingItem Then
                        quickConversationId = DirectCast(quickItem, Outlook.MeetingItem).ConversationID
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

            ' 检查当前邮件是否已完全加载，如果未加载则延迟重试
            Dim currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem IsNot Nothing AndAlso Not OutlookMyList.Utils.OutlookUtils.IsMailItemReady(currentItem) Then
                Debug.WriteLine("邮件未完全加载，延迟800ms后重试...")
                Await Task.Delay(800)
                
                ' 重试检查
                currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem IsNot Nothing AndAlso Not OutlookMyList.Utils.OutlookUtils.IsMailItemReady(currentItem) Then
                    Debug.WriteLine("邮件仍然未完全加载，尝试重试机制")
                    
                    ' 尝试等待邮件加载完成
                    Dim isReady As Boolean = Await OutlookMyList.Utils.OutlookUtils.WaitForMailItemReady(currentItem, 1000, 200)
                    
                    If Not isReady Then
                        Debug.WriteLine("邮件加载超时，显示加载中提示")
                        HideProgress()
                        
                        ' 清空并显示友好的加载提示
                        lvMails.BeginUpdate()
                        lvMails.Items.Clear()
                        
                        Dim loadingItem As New ListViewItem("邮件内容正在加载中，请稍候...")
                        loadingItem.Tag = "LOADING_PLACEHOLDER"
                        loadingItem.ForeColor = SystemColors.GrayText
                        
                        lvMails.Items.Add(loadingItem)
                        lvMails.EndUpdate()
                        
                        Return
                    Else
                        Debug.WriteLine("邮件加载完成，继续处理")
                    End If
                End If
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
                                   lvMails.EndUpdate()
                               End Sub)
            Else
                lvMails.BeginUpdate()
                lvMails.Items.Clear()
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

        ' 首先检查缓存（使用智能会话ID）
        Dim smartId As String = String.Empty
        Try
            currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
            If currentItem IsNot Nothing Then
                ' 检查邮件是否已完全加载，带重试机制
                Dim maxRetries As Integer = 3
                Dim retryCount As Integer = 0
                Dim isReady As Boolean = False
                
                While retryCount < maxRetries AndAlso Not isReady
                    If OutlookMyList.Utils.OutlookUtils.IsMailItemReady(currentItem) Then
                        isReady = True
                        Exit While
                    End If
                    
                    retryCount += 1
                    System.Threading.Thread.Sleep(100) ' 后台线程使用同步等待
                End While
                
                If Not isReady Then
                    Debug.WriteLine($"当前邮件未完全加载，重试{maxRetries}次后跳过: {currentMailEntryID}")
                    Return
                End If
                
                ' 统一使用智能会话键（优先自定义ID）
                smartId = GetSmartConversationKey(currentItem)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"获取会话ID失败: {ex.Message}")
        End Try

        ' 如果智能会话ID相同，只需要更新高亮
        If Not String.IsNullOrEmpty(smartId) AndAlso
           String.Equals(smartId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
            Debug.WriteLine($"智能会话ID未变化({smartId})，只更新高亮")
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
        Debug.WriteLine($"处理邮件: 智能会话ID={If(String.IsNullOrEmpty(smartId), "无", smartId)}, EntryID={currentMailEntryID}")

        ' 检查会话缓存（遵循全局缓存开关）
        If IsCacheEnabled() AndAlso Not String.IsNullOrEmpty(smartId) AndAlso conversationMailsCache.ContainsKey(smartId) Then
            Dim cachedData = conversationMailsCache(smartId)
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
                conversationMailsCache.Remove(smartId)
            End If
        End If

        Try
            Try
                currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem Is Nothing Then
                    Throw New System.Exception("无法获取邮件项")
                End If

                ' 检查当前邮件项是否已完全加载
                If Not OutlookMyList.Utils.OutlookUtils.IsMailItemReady(currentItem) Then
                    Debug.WriteLine("当前邮件项未完全加载，等待重试...")
                    ' 如果当前邮件未完全加载，跳过此次更新
                    Return
                End If

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                    conversation = DirectCast(currentItem, Outlook.MeetingItem).GetConversation()
                End If

                ' 如果当前邮件存在自定义会话ID，优先按自定义会话ID构建列表
                Dim customId As String = ReadCustomConversationIdFromItem(currentItem)
                If Not String.IsNullOrEmpty(customId) Then
                    Debug.WriteLine($"检测到自定义会话ID: {customId}，优先按自定义ID分组加载")

                    ' 初始化集合
                    allItems = New List(Of ListViewItem)(50)
                    tempMailItems = New List(Of (Index As Integer, EntryID As String))(50)

                    ' 获取所有核心邮件文件夹
                    Dim allMailFolders As New List(Of Outlook.Folder)
                    Try
                        ' 获取所有核心邮件文件夹（已包含归档文件夹）
                        Dim rootFolder As Outlook.Folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
                        GetAllMailFolders(rootFolder, allMailFolders)
                        
                        If allMailFolders.Count = 0 Then
                            allMailFolders.Add(DirectCast(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), Outlook.Folder))
                        End If
                        Debug.WriteLine($"自定义会话ID搜索将覆盖 {allMailFolders.Count} 个邮件文件夹")
                    Catch ex As System.Exception
                        Debug.WriteLine($"获取核心邮件文件夹失败: {ex.Message}")
                    End Try

                    Dim currentIndex As Integer = 0
                    For Each folder As Outlook.Folder In allMailFolders
                        Try
                            ' 优先尝试使用Restrict按用户定义属性过滤
                            Dim items As Outlook.Items = Nothing
                            Try
                                ' 使用更高效的搜索方式：先尝试使用Restrict，如果失败则使用Find方法
                                Dim filter As String = $"[CustomConversationId] = '{customId}'"
                                items = folder.Items.Restrict(filter)
                                Debug.WriteLine($"在文件夹 {folder.Name} 使用过滤器查找自定义会话邮件: {filter}, 结果数={items.Count}")
                            Catch ex As System.Runtime.InteropServices.COMException
                                Debug.WriteLine($"使用Restrict过滤自定义属性失败: {ex.Message}")
                                ' 如果自定义属性不存在，则使用Find方法搜索
                                If ex.Message.Contains("unknown") OrElse ex.Message.Contains("Unknown") Then
                                    ' 使用Find方法搜索，这比遍历所有项目更高效
                                    items = New Outlook.Items()
                                    Try
                                        ' 使用Find方法搜索Subject字段，这只是为了获取一个Items集合
                                        Dim tempItems = folder.Items
                                        tempItems.Sort("[ReceivedTime]", False)
                                        
                                        ' 限制搜索范围，只处理最近的100封邮件
                                        Dim maxItemsToCheck As Integer = Math.Min(100, tempItems.Count)
                                        Dim itemsChecked As Integer = 0
                                        
                                        For i As Integer = 1 To tempItems.Count
                                            If itemsChecked >= maxItemsToCheck Then Exit For
                                            
                                            Try
                                                Dim item As Object = tempItems.Item(i)
                                                itemsChecked += 1
                                                
                                                ' 检查自定义会话ID
                                                Dim itemCustomId As String = ReadCustomConversationIdFromItem(item)
                                                If Not String.IsNullOrEmpty(itemCustomId) AndAlso itemCustomId = customId Then
                                                    items.Add(item)
                                                End If
                                            Catch ex2 As System.Exception
                                                ' 忽略单个项目的错误，继续处理下一个
                                            End Try
                                        Next
                                        
                                        Debug.WriteLine($"在文件夹 {folder.Name} 中使用优化搜索，检查了 {itemsChecked} 项，找到 {items.Count} 个匹配项")
                                    Catch ex2 As System.Exception
                                        Debug.WriteLine($"优化搜索出现异常: {ex2.Message}")
                                    End Try
                                End If
                            Catch ex As System.Exception
                                Debug.WriteLine($"Restrict出现异常: {ex.Message}")
                            End Try

                            If items IsNot Nothing AndAlso items.Count > 0 Then
                                For Each it As Object In items
                                    Try
                                        
                                        ' 检查邮件项是否已完全加载
                                        If Not OutlookMyList.Utils.OutlookUtils.IsMailItemReady(it) Then
                                            Debug.WriteLine("邮件项未完全加载，跳过处理")
                                            Continue For
                                        End If

                                        Dim entryId As String = GetPermanentEntryID(it)
                                        Dim subject As String = "无主题"
                                        Dim senderName As String = "未知发件人"
                                        Dim receivedTime As DateTime = DateTime.MinValue
                                        Dim messageClass As String = ""

                                        Try
                                            subject = If(it.Subject, "无主题")
                                        Catch
                                            subject = "无法访问"
                                        End Try

                                        Try
                                            If TypeOf it Is Outlook.MailItem Then
                                                senderName = If(DirectCast(it, Outlook.MailItem).SenderName, "未知发件人")
                                                receivedTime = DirectCast(it, Outlook.MailItem).ReceivedTime
                                            ElseIf TypeOf it Is Outlook.AppointmentItem Then
                                                senderName = If(DirectCast(it, Outlook.AppointmentItem).Organizer, "未知组织者")
                                                receivedTime = DirectCast(it, Outlook.AppointmentItem).Start
                                            ElseIf TypeOf it Is Outlook.MeetingItem Then
                                                senderName = If(DirectCast(it, Outlook.MeetingItem).SenderName, "未知发件人")
                                                receivedTime = DirectCast(it, Outlook.MeetingItem).ReceivedTime
                                            End If
                                        Catch
                                        End Try

                                        Try
                                            messageClass = If(it.MessageClass, "")
                                        Catch
                                            messageClass = ""
                                        End Try

                                        ' 快速图标文本
                                        Dim hasAttach As Boolean = False
                                        Dim flagStatus As Integer = 0
                                        Try
                                            If TypeOf it Is Outlook.MailItem Then
                                                Dim mailForAttach = DirectCast(it, Outlook.MailItem)
                                                hasAttach = (mailForAttach.Attachments IsNot Nothing AndAlso mailForAttach.Attachments.Count > 0)
                                                flagStatus = mailForAttach.FlagStatus
                                            End If
                                        Catch
                                        End Try

                                        Dim iconText As String = GetIconTextFast(messageClass, hasAttach, flagStatus)

                                        Dim lvi As New ListViewItem(iconText) With {
                                            .Tag = entryId,
                                            .Name = currentIndex.ToString()
                                        }
                                        With lvi.SubItems
                                            .Add(If(receivedTime <> DateTime.MinValue, receivedTime.ToString("yyyy-MM-dd HH:mm"), "无时间"))
                                            .Add(senderName)
                                            .Add(subject)
                                        End With

                                        allItems.Add(lvi)
                                        tempMailItems.Add((currentIndex, entryId))
                                        currentIndex += 1
                                    Catch
                                    End Try
                                Next
                            End If
                        Catch ex As System.Exception
                            Debug.WriteLine($"扫描文件夹 {folder.Name} 时出错: {ex.Message}")
                        End Try
                    Next

                    ' 构建完成后跳转到UI更新
                    GoTo UpdateUI
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

                        ' 应用主题到新创建的项目
                        ApplyThemeToListViewItem(lvi)

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

                        ' 应用主题到错误项目
                        ApplyThemeToListViewItem(errorItem)

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
                ' 根据错误设置决定是否显示错误弹窗
                If ErrorNotificationSettings.Instance.ShowErrorDialogs Then
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
                Else
                    ' 静默失败，只在列表中显示友好的错误提示
                    Me.BeginInvoke(Sub()
                                       lvMails.BeginUpdate()
                                       lvMails.Items.Clear()
                                       Dim errorItem As New ListViewItem("邮件加载遇到问题，请稍后重试")
                                       errorItem.SubItems.Add("")
                                       errorItem.SubItems.Add("")
                                       errorItem.SubItems.Add("")
                                       errorItem.ForeColor = System.Drawing.Color.Gray
                                       lvMails.Items.Add(errorItem)
                                       lvMails.EndUpdate()
                                   End Sub)
                End If
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
        ' 优化缓存策略：只缓存合理大小的会话，减少内存占用（使用智能会话ID，遵循开关）
        If IsCacheEnabled() AndAlso Not String.IsNullOrEmpty(smartId) AndAlso allItems.Count > 0 AndAlso allItems.Count <= 50 Then
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

                conversationMailsCache(smartId) = (New List(Of (Index As Integer, EntryID As String))(tempMailItems), cacheItems, DateTime.Now)
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
                                           AddItemsWithTheme(lvMails, allItems.ToArray())
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
                                           AddItemsWithTheme(lvMails, clones)
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

                               ' 更新会话信息按钮标题
                               UpdateConversationInfoButtonTitle()

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
                                       Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
                                       UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)
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
                currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
                If currentItem Is Nothing Then
                    Throw New System.Exception("无法获取邮件项")
                End If

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                    conversation = DirectCast(currentItem, Outlook.MeetingItem).GetConversation()
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
                                .Add(GetFriendlySenderName(mail))
                                .Add(mail.Subject)
                            ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                                Dim appt As Outlook.AppointmentItem = DirectCast(currentItem, Outlook.AppointmentItem)
                                .Add(appt.Start.ToString("yyyy-MM-dd HH:mm"))
                                .Add(appt.Organizer)
                                .Add(appt.Subject)
                            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                                Dim meeting As Outlook.MeetingItem = DirectCast(currentItem, Outlook.MeetingItem)
                                .Add(meeting.ReceivedTime.ToString("yyyy-MM-dd HH:mm"))
                                .Add(GetFriendlyOrganizerName(meeting))
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

                    AddItemWithTheme(lvMails, lvi)
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
                            AddItemsWithTheme(lvMails, clones2.ToArray())
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

                        ' 更新会话信息按钮标题
                        UpdateConversationInfoButtonTitle()

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
                currentItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(currentMailEntryID)
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
                        AddItemsWithTheme(lvMails, clones3.ToArray())
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
                ' 记录调试信息
                If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                    Debug.WriteLine($"Failed to process mail item: {ex.Message}")
                End If
                
                ' 根据配置决定是否显示错误（使用统一的错误处理策略）
                If ThisAddIn.ErrorSettings.ShowErrorDialogs Then
                    ThisAddIn.ShowErrorWithConfig("处理邮件时出错", ex.Message)
                End If
            End Try

        Catch ex As System.Exception
            ' 记录调试信息
            If ThisAddIn.ErrorSettings.LogErrorsToDebug Then
                Debug.WriteLine($"LoadConversationMails error: {ex.Message}")
            End If
            
            ' 根据配置决定是否显示错误（使用统一的错误处理策略）
            If ThisAddIn.ErrorSettings.ShowErrorDialogs Then
                ThisAddIn.ShowErrorWithConfig("加载邮件时出错", ex.Message)
            End If
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

        ' 确保suppressWebViewUpdate初始状态为0
        suppressWebViewUpdate = 0
        Debug.WriteLine($"MailThreadPane构造函数: suppressWebViewUpdate重置为 {suppressWebViewUpdate}")

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
                            If IsCacheEnabled() Then
                                If Not entryIdCompareCache.TryGetValue(cacheKey, itemEntryID) Then
                                    itemEntryID = ConvertEntryIDToString(rawTag)
                                    entryIdCompareCache(cacheKey) = itemEntryID
                                End If
                            Else
                                itemEntryID = ConvertEntryIDToString(rawTag)
                            End If
                            ' 尝试使用CompareEntryIDs进行MAPI级别的比较，如果失败回退到字符串比较
                            Dim isMatchedOld As Boolean = False
                            Dim normalizedOldEntryID As String = ConvertStringToHexFormat(oldEntryID.Trim())
                            Try
                                isMatchedOld = Globals.ThisAddIn.Application.Session.CompareEntryIDs(itemEntryID, normalizedOldEntryID)
                            Catch ex As System.Exception
                                Debug.WriteLine($"UpdateHighlightByEntryID: CompareEntryIDs(Old)失败: {ex.Message}, 回退到字符串比较")
                                Dim shortOldEntryID As String = OutlookMyList.Utils.OutlookUtils.GetShortEntryID(normalizedOldEntryID)
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
                    Dim shortNewEntryID As String = OutlookMyList.Utils.OutlookUtils.GetShortEntryID(normalizedNewEntryID)
                    Debug.WriteLine($"UpdateHighlightByEntryID: 规范化后(长)={normalizedNewEntryID}, 转换短格式={shortNewEntryID}")
                    
                    ' 添加Inspector窗口场景的特殊调试
                    Debug.WriteLine($"UpdateHighlightByEntryID: lvMails项目总数={lvMails.Items.Count}")

                    For Each item As ListViewItem In lvMails.Items
                        If item.Tag IsNot Nothing Then
                            ' 取缓存的规范化ItemEntryID（避免重复Convert）
                            Dim rawTag = item.Tag
                            Dim cacheKey As String = If(TypeOf rawTag Is String, DirectCast(rawTag, String), ConvertEntryIDToString(rawTag))
                            Dim itemEntryID As String = String.Empty
                            If IsCacheEnabled() Then
                                If Not entryIdCompareCache.TryGetValue(cacheKey, itemEntryID) Then
                                    itemEntryID = ConvertEntryIDToString(rawTag)
                                    entryIdCompareCache(cacheKey) = itemEntryID
                                End If
                            Else
                                itemEntryID = ConvertEntryIDToString(rawTag)
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
                                shortNewEntryID = OutlookMyList.Utils.OutlookUtils.GetShortEntryID(normalizedNewEntryID)
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
        Try
            Debug.WriteLine($"SetItemHighlight: 开始设置项目高亮状态, isHighlighted={isHighlighted}")
            
            ' 检查item是否有效
            If item Is Nothing Then
                Debug.WriteLine("SetItemHighlight: 错误 - 传入的ListViewItem为null")
                Return
            End If
            
            ' 检查控件是否已创建句柄
            If Not lvMails.IsHandleCreated Then
                Debug.WriteLine("SetItemHighlight: 警告 - lvMails控件句柄尚未创建")
                Try
                    Dim handle = lvMails.Handle
                    Debug.WriteLine($"SetItemHighlight: 强制创建句柄结果: {handle}")
                Catch ex As Exception
                    Debug.WriteLine($"SetItemHighlight: 创建句柄失败: {ex.Message}")
                End Try
            End If
            
            ' 确保在UI线程上执行
            If Me.InvokeRequired Then
                Debug.WriteLine("SetItemHighlight: 需要跨线程调用，使用Invoke")
                Me.Invoke(Sub() SetItemHighlight(item, isHighlighted))
                Return
            End If
            
            ' 设置高亮状态
            If isHighlighted Then
                Debug.WriteLine("SetItemHighlight: 设置高亮状态")
                item.BackColor = highlightColor
                item.ForeColor = If(IsDarkTheme(currentBackColor), Color.White, Color.Black) ' 确保文本颜色与背景对比鲜明
                item.Font = New Font(item.Font, FontStyle.Bold) ' 使用粗体增强可见性
                item.Selected = True
                
                ' 确保项目可见
                item.EnsureVisible()
                
                ' 添加边框效果增强可见性
                item.UseItemStyleForSubItems = True
                
                ' 强制刷新控件
                lvMails.Update() ' 更新控件
                lvMails.Refresh() ' 刷新显示
                Application.DoEvents() ' 处理所有待处理的Windows消息
                
                ' 再次确保项目可见
                item.EnsureVisible()
                Debug.WriteLine($"SetItemHighlight: 设置高亮后 -> BackColor={item.BackColor.Name}, ForeColor={item.ForeColor.Name}, FontBold={item.Font.Bold}, Selected={item.Selected}")
            Else
                Debug.WriteLine("SetItemHighlight: 清除高亮状态")
                item.BackColor = currentBackColor  ' 使用当前主题背景色
                item.ForeColor = currentForeColor  ' 使用当前主题文字色
                item.Font = normalFont
                item.Selected = False  ' 确保取消选中状态
                Debug.WriteLine($"SetItemHighlight: 清除高亮后 -> BackColor={item.BackColor.Name}, ForeColor={item.ForeColor.Name}, FontBold={item.Font.Bold}, Selected={item.Selected}")
            End If
            
            ' 强制刷新控件
            lvMails.Refresh()
            Debug.WriteLine("SetItemHighlight: 高亮设置完成")
        Catch ex As System.Exception
            Debug.WriteLine($"SetItemHighlight: 设置高亮时发生异常: {ex.Message}")
            Debug.WriteLine($"SetItemHighlight: 异常堆栈: {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' 应用当前主题到ListView项目
    ''' </summary>
    ''' <param name="item">要应用主题的ListView项目</param>
    Private Sub ApplyThemeToListViewItem(item As ListViewItem)
        If item IsNot Nothing Then
            ' 如果项目当前是高亮的，不要覆盖高亮颜色
            If item.BackColor = highlightColor Then
                ' 高亮项目保持高亮颜色，只更新文字颜色
                item.ForeColor = currentForeColor
                For Each subItem As ListViewItem.ListViewSubItem In item.SubItems
                    subItem.BackColor = highlightColor
                    subItem.ForeColor = currentForeColor
                Next
            Else
                ' 非高亮项目应用主题颜色
                item.BackColor = currentBackColor
                item.ForeColor = currentForeColor
                
                ' 确保子项也应用相同的主题颜色
                For Each subItem As ListViewItem.ListViewSubItem In item.SubItems
                    subItem.BackColor = currentBackColor
                    subItem.ForeColor = currentForeColor
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' 应用主题到所有ListView控件的所有项目
    ''' </summary>
    Private Sub ApplyThemeToAllListViewItems()
        Try
            ' 应用到主邮件ListView
            If lvMails IsNot Nothing Then
                For Each item As ListViewItem In lvMails.Items
                    ApplyThemeToListViewItem(item)
                Next
                Debug.WriteLine($"已应用主题到 {lvMails.Items.Count} 个主邮件ListView项目")
            End If

            ' 应用到任务列表
            If taskList IsNot Nothing Then
                For Each item As ListViewItem In taskList.Items
                    ApplyThemeToListViewItem(item)
                Next
                Debug.WriteLine($"已应用主题到 {taskList.Items.Count} 个任务列表项目")
            End If

            ' 应用到邮件历史列表
            If mailHistoryList IsNot Nothing Then
                For Each item As ListViewItem In mailHistoryList.Items
                    ApplyThemeToListViewItem(item)
                Next
                Debug.WriteLine($"已应用主题到 {mailHistoryList.Items.Count} 个邮件历史项目")
            End If

            ' 应用到待办邮件列表
            If pendingMailList IsNot Nothing Then
                For Each item As ListViewItem In pendingMailList.Items
                    ApplyThemeToListViewItem(item)
                Next
                Debug.WriteLine($"已应用主题到 {pendingMailList.Items.Count} 个待办邮件项目")
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"ApplyThemeToAllListViewItems错误: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 添加项目到ListView并应用主题
    ''' </summary>
    ''' <param name="listView">目标ListView</param>
    ''' <param name="item">要添加的项目</param>
    Private Sub AddItemWithTheme(listView As ListView, item As ListViewItem)
        If listView IsNot Nothing AndAlso item IsNot Nothing Then
            ApplyThemeToListViewItem(item)
            listView.Items.Add(item)
        End If
    End Sub

    ''' <summary>
    ''' 批量添加ListView项目并应用主题
    ''' </summary>
    ''' <param name="listView">目标ListView</param>
    ''' <param name="items">要添加的项目数组</param>
    Private Sub AddItemsWithTheme(listView As ListView, items As ListViewItem())
        If listView IsNot Nothing AndAlso items IsNot Nothing Then
            For Each item In items
                If item IsNot Nothing Then
                    ApplyThemeToListViewItem(item)
                End If
            Next
            listView.Items.AddRange(items)
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

    Private Async Sub lvMails_MouseDown(sender As Object, e As MouseEventArgs) Handles lvMails.MouseDown
        Debug.WriteLine($"lvMails_MouseDown: 方法开始执行。Sender类型: {sender.GetType().Name}, EventArgs类型: {e.GetType().Name}, SelectedItems.Count = {lvMails.SelectedItems.Count}")
        Debug.WriteLine($"lvMails_MouseDown: Enabled: {lvMails.Enabled}, Visible: {lvMails.Visible}, Focused: {lvMails.Focused}, ContainsFocus: {lvMails.ContainsFocus}")
    End Sub

    Private Sub lvMails_SelectedIndexChanged(sender As Object, e As EventArgs)
        Debug.WriteLine($"lvMails_SelectedIndexChanged: 方法开始执行。Sender类型: {sender.GetType().Name}, EventArgs类型: {e.GetType().Name}")
        Debug.WriteLine($"lvMails_SelectedIndexChanged: 方法开始时选中项数量 = {lvMails.SelectedItems.Count}")
        Try
            Debug.WriteLine($"lvMails_SelectedIndexChanged: 选中项数量 = {lvMails.SelectedItems.Count}")
            If lvMails.SelectedItems.Count = 0 Then
                Debug.WriteLine("lvMails_SelectedIndexChanged: 没有选中项，返回。")
                Return
            End If

            Dim mailId As String = ConvertEntryIDToString(lvMails.SelectedItems(0).Tag)
            Debug.WriteLine($"lvMails_SelectedIndexChanged: 获取到的邮件ID = {mailId}")
            Debug.WriteLine($"lvMails_SelectedIndexChanged: 当前高亮邮件ID (currentMailEntryID) = {currentMailEntryID}")
            Debug.WriteLine($"[邮件选择] 当前全局主题变量: 背景={globalThemeBackgroundColor}, 前景={globalThemeForegroundColor}, 更新时间={globalThemeLastUpdate}")
            If String.IsNullOrEmpty(mailId) Then
                Debug.WriteLine("lvMails_SelectedIndexChanged: 邮件ID为空，返回。")
                Return
            End If

            ' 始终更新高亮，不受suppressWebViewUpdate影响
            If Not mailId.Equals(currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                Debug.WriteLine($"lvMails_SelectedIndexChanged: 邮件ID不同，准备更新高亮。旧邮件ID = {currentMailEntryID}, 新邮件ID = {mailId}")
                Dim oldMailId As String = currentMailEntryID
                currentMailEntryID = mailId
                UpdateHighlightByEntryID(oldMailId, mailId)
                lvMails.Refresh() ' 强制刷新ListView以确保高亮立即显示
                Debug.WriteLine("lvMails_SelectedIndexChanged: UpdateHighlightByEntryID 已调用。")

                ' 只有在非抑制模式下才加载WebView内容
                Debug.WriteLine($"lvMails_SelectedIndexChanged: suppressWebViewUpdate = {suppressWebViewUpdate}")
                If suppressWebViewUpdate = 0 Then
                    Debug.WriteLine($"lvMails_SelectedIndexChanged: 开始加载WebView内容，邮件ID = {mailId}")
                    Me.BeginInvoke(New Action(Of String)(AddressOf LoadMailContentDeferred), mailId)
                    Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
                    UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)
                Else
                    Debug.WriteLine($"lvMails_SelectedIndexChanged: WebView更新被抑制，suppressWebViewUpdate = {suppressWebViewUpdate}")
                End If
            Else
                Debug.WriteLine($"lvMails_SelectedIndexChanged: 邮件ID相同，跳过更新")
            End If


        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_SelectedIndexChanged error: {ex.Message}")
        End Try
        Debug.WriteLine("lvMails_SelectedIndexChanged: 方法执行结束")
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
                                                    Return OutlookMyList.Handlers.MailHandler.DisplayMailContent(mailId)
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
            Debug.WriteLine($"LoadMailContentDeferred 开始执行，邮件ID: {mailId}")

            ' 抑制期间不进行 WebView 更新，避免联系人信息列表构造时触发刷新
            If suppressWebViewUpdate > 0 Then
                Debug.WriteLine($"WebView更新被抑制，延迟重试 LoadMailContentDeferred: {mailId}, suppressWebViewUpdate = {suppressWebViewUpdate}")
                Await Task.Delay(100)
                If suppressWebViewUpdate = 0 AndAlso Me.IsHandleCreated Then
                    Debug.WriteLine($"重试 LoadMailContentDeferred: {mailId}")
                    Me.BeginInvoke(Sub() LoadMailContentDeferred(mailId))
                End If
                Return
            End If

            Debug.WriteLine($"开始获取邮件HTML内容，邮件ID: {mailId}")
            Dim html As String = Await Task.Run(Function() OutlookMyList.Handlers.MailHandler.DisplayMailContent(mailId))
            Debug.WriteLine($"获取到HTML内容，长度: {If(html IsNot Nothing, html.Length, 0)}")

            If mailBrowser IsNot Nothing AndAlso mailBrowser.IsHandleCreated AndAlso suppressWebViewUpdate = 0 Then
                Debug.WriteLine($"开始更新WebView内容，邮件ID: {mailId}")
                
                Debug.WriteLine("准备绑定WebBrowser事件")
                AddHandler mailBrowser.DocumentCompleted, AddressOf WebBrowser_DocumentCompleted
                AddHandler mailBrowser.Navigated, AddressOf MailBrowser_Navigated
                AddHandler mailBrowser.ProgressChanged, AddressOf MailBrowser_ProgressChanged
                Debug.WriteLine("WebBrowser事件绑定完成")

                Debug.WriteLine($"设置DocumentText，内容长度={If(html IsNot Nothing, html.Length, 0)}")
                mailBrowser.DocumentText = html
                isDisplayingMailContent = True
                Debug.WriteLine($"DocumentText已设置后状态: readyState={mailBrowser.ReadyState}, isDisplayingMailContent={isDisplayingMailContent}")

                ' 验证WebView内容是否真的被设置
                Dim actualContent = mailBrowser.DocumentText
                Debug.WriteLine($"WebView实际内容长度: {If(actualContent IsNot Nothing, actualContent.Length, 0)}")
                If actualContent IsNot Nothing AndAlso actualContent.Length > 0 Then
                    Dim preview = If(actualContent.Length > 200, actualContent.Substring(0, 200), actualContent)
                    Debug.WriteLine($"WebView内容预览: {preview}")
                End If

                Debug.WriteLine("已添加DocumentCompleted事件处理程序，等待应用主题")
                
                ' 立即应用主题到WebBrowser控件本身
                Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
                UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)

                Dim attempts As Integer = 0
                Dim themeTimer As New System.Windows.Forms.Timer()
                themeTimer.Interval = 150
                AddHandler themeTimer.Tick, Sub(sender2, e2)
                    Try
                        attempts += 1
                        Debug.WriteLine($"[ThemeDebug] TimerTick: attempt={attempts}, readyState={mailBrowser.ReadyState}, docIsNull={mailBrowser.Document Is Nothing}, isDisplayingMailContent={isDisplayingMailContent}")
                        If mailBrowser.ReadyState = WebBrowserReadyState.Complete AndAlso isDisplayingMailContent AndAlso mailBrowser.Document IsNot Nothing Then
                            Dim doc = mailBrowser.Document
                            Dim bgColorHex As String = $"#{currentBackColor.R:X2}{currentBackColor.G:X2}{currentBackColor.B:X2}"
                            Dim fgColorHex As String = $"#{currentForeColor.R:X2}{currentForeColor.G:X2}{currentForeColor.B:X2}"
                            Dim accentColorHex As String = "#0078d7"
                            Dim script As String = $"
                                (function() {{
                                    var elements = document.getElementsByTagName('*');
                                    for (var i = 0; i < elements.length; i++) {{
                                        var elem = elements[i];
                                        var tagName = elem.tagName.toUpperCase();
                                        if (tagName !== 'STYLE' && tagName !== 'SCRIPT') {{
                                            elem.style.setProperty('background-color', '{bgColorHex}', 'important');
                                            elem.style.setProperty('color', '{fgColorHex}', 'important');
                                            if (tagName !== 'BODY') {{
                                                elem.style.setProperty('background-color', 'transparent', 'important');
                                            }}
                                            elem.removeAttribute('color');
                                            elem.removeAttribute('bgcolor');
                                        }}
                                    }}
                                    document.body.style.setProperty('background-color', '{bgColorHex}', 'important');
                                    document.body.style.setProperty('color', '{fgColorHex}', 'important');
                                }})();
                            "
                            Debug.WriteLine($"[ThemeDebug] TimerTick: applying theme back={bgColorHex}, fore={fgColorHex}")
                            Dim applyResult As Object = doc.InvokeScript("eval", New Object() {script})
                            Debug.WriteLine("[ThemeDebug] TimerTick: theme applied via eval")
                            themeTimer.Stop()
                        ElseIf attempts >= 20 Then
                            Debug.WriteLine("[ThemeDebug] TimerTick: max attempts reached, stop timer")
                            themeTimer.Stop()
                        End If
                    Catch ex As System.Exception
                        Debug.WriteLine($"[ThemeDebug] TimerTick error: {ex.Message}")
                        themeTimer.Stop()
                    End Try
                End Sub
                themeTimer.Start()
            Else
                Debug.WriteLine($"跳过WebView更新 - mailBrowser IsNot Nothing: {mailBrowser IsNot Nothing}, IsHandleCreated: {If(mailBrowser IsNot Nothing, mailBrowser.IsHandleCreated, False)}, suppressWebViewUpdate: {suppressWebViewUpdate}")
            End If


        Catch ex As System.Exception
            Debug.WriteLine($"LoadMailContentDeferred error: {ex.Message}")
            Debug.WriteLine($"LoadMailContentDeferred StackTrace: {ex.StackTrace}")
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
            ' 添加空值检查
            If lvMails Is Nothing Then
                Debug.WriteLine("lvMails_DoubleClick: lvMails为Nothing")
                Return
            End If

            If lvMails.SelectedItems Is Nothing OrElse lvMails.SelectedItems.Count = 0 Then
                Debug.WriteLine("lvMails_DoubleClick: 没有选中项")
                Return
            End If

            Dim selectedItem As ListViewItem = lvMails.SelectedItems(0)
            If selectedItem Is Nothing Then
                Debug.WriteLine("lvMails_DoubleClick: 选中项为Nothing")
                Return
            End If

            Dim mailId As String = ""
            Try
                mailId = ConvertEntryIDToString(selectedItem.Tag)
            Catch ex As System.Exception
                Debug.WriteLine($"lvMails_DoubleClick: ConvertEntryIDToString异常: {ex.Message}")
                Return
            End Try

            If String.IsNullOrEmpty(mailId) Then
                Debug.WriteLine("lvMails_DoubleClick: EntryID为空")
                Return
            End If

            ' 更新邮件列表高亮显示
            Dim oldEntryId As String = currentHighlightEntryID
            UpdateHighlightByEntryID(oldEntryId, mailId)

            ' 使用更安全的方式打开邮件
            Try
                ' 优先使用快速打开（可进一步传StoreID优化）
                Dim openResult As Boolean = False
                Try
                    openResult = OutlookMyList.Utils.OutlookUtils.FastOpenMailItem(mailId)
                Catch ex As System.Exception
                    Debug.WriteLine($"lvMails_DoubleClick: FastOpenMailItem异常: {ex.Message}")
                    openResult = False
                End Try

                If Not openResult Then
                    ' 兜底：GetItemFromID + Display
                    Dim mailItem As Object = Nothing
                    Try
                        mailItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mailId)
                        If mailItem IsNot Nothing Then
                            Try
                                mailItem.Display()
                                Debug.WriteLine($"lvMails_DoubleClick: 邮件打开成功 EntryID={mailId}")
                            Catch displayEx As System.Exception
                                Debug.WriteLine($"lvMails_DoubleClick: Display异常: {displayEx.Message}")
                            End Try
                        Else
                            Debug.WriteLine($"lvMails_DoubleClick: 无法获取邮件项 EntryID={mailId}")
                        End If
                    Finally
                        If mailItem IsNot Nothing Then
                            Try
                                OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                            Catch releaseEx As System.Exception
                                Debug.WriteLine($"lvMails_DoubleClick: 释放COM对象异常: {releaseEx.Message}")
                            End Try
                        End If
                    End Try
                Else
                    Debug.WriteLine($"lvMails_DoubleClick: 快速打开成功 EntryID={mailId}")
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"lvMails_DoubleClick: 打开邮件过程异常: {ex.Message}")
            End Try

        Catch ex As System.Runtime.InteropServices.COMException
            ' COM异常静默处理，只记录调试信息，绝不抛出
            Debug.WriteLine($"lvMails_DoubleClick COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
        Catch ex As System.Exception
            ' 其他异常也静默处理，绝不抛出，避免弹出错误对话框
            Debug.WriteLine($"lvMails_DoubleClick 异常: {ex.Message}")
        End Try
    End Sub

    Private Sub TaskList_DoubleClick(sender As Object, e As EventArgs)
        Try
            ' 检查taskList是否为Nothing
            If taskList Is Nothing Then
                Debug.WriteLine("TaskList_DoubleClick: taskList is Nothing")
                Return
            End If

            If taskList.SelectedItems.Count > 0 Then
                Dim selectedItem As ListViewItem = taskList.SelectedItems(0)
                If selectedItem.Tag IsNot Nothing Then
                    Dim entryId As String = ""

                    ' 检查 Tag 类型，获取相应的 EntryID
                    If TypeOf selectedItem.Tag Is OutlookMyList.Models.TaskInfo Then
                        Dim taskInfo As OutlookMyList.Models.TaskInfo = DirectCast(selectedItem.Tag, OutlookMyList.Models.TaskInfo)
                        ' 优先使用 TaskEntryID，如果为空则使用 MailEntryID
                        entryId = If(Not String.IsNullOrEmpty(taskInfo.TaskEntryID), taskInfo.TaskEntryID, taskInfo.MailEntryID)
                    Else
                        ' 兜底：将 Tag 作为 EntryID 字符串处理
                        entryId = ConvertEntryIDToString(selectedItem.Tag)
                    End If

                    If Not String.IsNullOrEmpty(entryId) Then
                        ' 更新邮件列表高亮显示
                        Dim oldEntryId As String = currentHighlightEntryID
                        UpdateHighlightByEntryID(oldEntryId, entryId)
                        
                        ' 优先使用快速打开（传入 StoreID 可进一步优化）
                        Dim storeId As String = Nothing
                        If TypeOf selectedItem.Tag Is OutlookMyList.Models.TaskInfo Then
                            storeId = DirectCast(selectedItem.Tag, OutlookMyList.Models.TaskInfo).StoreID
                        End If
                        If Not OutlookMyList.Utils.OutlookUtils.FastOpenMailItem(entryId, storeId) Then
                            ' 兜底：GetItemFromID + Display
                            Dim taskItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId, storeId)
                            If taskItem IsNot Nothing Then
                                Try
                                    taskItem.Display()
                                Finally
                                    OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(taskItem)
                                End Try
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            ' COM异常静默处理，只记录调试信息，绝不抛出
            Debug.WriteLine($"TaskList_DoubleClick COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
        Catch ex As System.Exception
            ' 其他异常也静默处理，绝不抛出，避免弹出错误对话框
            Debug.WriteLine($"TaskList_DoubleClick 异常: {ex.Message}")
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
                               OutlookMyList.Handlers.TaskHandler.CreateNewTask(currentConversationId, currentMailEntryID)
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

    ' 右键菜单事件处理方法
    Private Sub ShowConversationId_Click(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then
                MessageBox.Show("请先选择一封邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedItem As ListViewItem = lvMails.SelectedItems(0)
            Dim entryId As String = selectedItem.Tag?.ToString()

            If String.IsNullOrEmpty(entryId) Then
                Globals.ThisAddIn.ShowErrorWithConfig("无法获取邮件EntryID", "操作失败")
                Return
            End If

            ' 获取邮件项并提取会话ID
            Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
            If mailItem IsNot Nothing Then
                Try
                    Dim conversationId As String = ""
                    If TypeOf mailItem Is Outlook.MailItem Then
                        conversationId = DirectCast(mailItem, Outlook.MailItem).ConversationID
                    ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                        conversationId = DirectCast(mailItem, Outlook.AppointmentItem).ConversationID
                    ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                        conversationId = DirectCast(mailItem, Outlook.MeetingItem).ConversationID
                    End If

                    If Not String.IsNullOrEmpty(conversationId) Then
                        MessageBox.Show($"会话ID: {conversationId}", "会话ID信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("该邮件没有会话ID", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Finally
                    OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                End Try
            Else
                Globals.ThisAddIn.ShowErrorWithConfig("无法获取邮件项", "操作失败")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ShowConversationId_Click error: {ex.Message}")
            Globals.ThisAddIn.ShowErrorWithConfig("获取会话ID时出错", ex.Message)
        End Try
    End Sub

    Private Sub CopyConversationId_Click(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then
                MessageBox.Show("请先选择一封邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedItem As ListViewItem = lvMails.SelectedItems(0)
            Dim entryId As String = selectedItem.Tag?.ToString()

            If String.IsNullOrEmpty(entryId) Then
                Globals.ThisAddIn.ShowErrorWithConfig("无法获取邮件EntryID", "操作失败")
                Return
            End If

            ' 获取邮件项并提取会话ID
            Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
            If mailItem IsNot Nothing Then
                Try
                    Dim conversationId As String = GetSafeConversationID(mailItem)
                    If Not String.IsNullOrEmpty(conversationId) Then
                        Clipboard.SetText(conversationId)
                        MessageBox.Show("会话ID已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("该邮件没有会话ID", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Finally
                    OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                End Try
            Else
                Globals.ThisAddIn.ShowErrorWithConfig("无法获取邮件项", "操作失败")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"CopyConversationId_Click error: {ex.Message}")
            Globals.ThisAddIn.ShowErrorWithConfig("复制会话ID时出错", ex.Message)
        End Try
    End Sub

    Private Sub ShowTaskStatus_Click(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then
                MessageBox.Show("请先选择一封邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedItem As ListViewItem = lvMails.SelectedItems(0)
            Dim entryId As String = selectedItem.Tag?.ToString()

            If String.IsNullOrEmpty(entryId) Then
                Globals.ThisAddIn.ShowErrorWithConfig("无法获取邮件EntryID", "操作失败")
                Return
            End If

            ' 获取邮件项并检查任务关联状态
            Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
            If mailItem IsNot Nothing Then
                Try
                    Dim taskStatus As TaskStatus = CheckItemHasTask(mailItem)
                    Dim statusText As String = ""

                    Select Case taskStatus
                        Case TaskStatus.InProgress
                            statusText = "进行中 🚩"
                        Case TaskStatus.Completed
                            statusText = "已完成 ⚑"
                        Case TaskStatus.None
                            statusText = "未关联任务"
                        Case Else
                            statusText = "未知状态"
                    End Select

                    MessageBox.Show($"任务关联状态: {statusText}", "任务状态信息", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Finally
                    OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                End Try
            Else
                Globals.ThisAddIn.ShowErrorWithConfig("无法获取邮件项", "操作失败")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"ShowTaskStatus_Click error: {ex.Message}")
            Globals.ThisAddIn.ShowErrorWithConfig("获取任务状态时出错", ex.Message)
        End Try
    End Sub

    Private Sub CustomConversationId_Click(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then
                MessageBox.Show("请先选择一封邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim selectedItem As ListViewItem = lvMails.SelectedItems(0)
            Dim entryId As String = selectedItem.Tag?.ToString()

            If String.IsNullOrEmpty(entryId) Then
                MessageBox.Show("无法获取邮件EntryID", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' 获取邮件项
            Dim mailItem As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
            If mailItem IsNot Nothing Then
                Try
                    ' 读取当前的自定义会话ID
                    Dim currentCustomId As String = ReadCustomConversationIdFromItem(mailItem)

                    ' 获取原始会话ID作为参考
                    Dim originalConversationId As String = ""
                    If TypeOf mailItem Is Outlook.MailItem Then
                        originalConversationId = DirectCast(mailItem, Outlook.MailItem).ConversationID
                    ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                        originalConversationId = DirectCast(mailItem, Outlook.AppointmentItem).ConversationID
                    ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                        originalConversationId = DirectCast(mailItem, Outlook.MeetingItem).ConversationID
                    End If

                    ' 构建提示信息
                    Dim promptMessage As String = "请输入自定义会话ID：" & Environment.NewLine & Environment.NewLine
                    promptMessage += $"原始会话ID: {originalConversationId}" & Environment.NewLine
                    If Not String.IsNullOrEmpty(currentCustomId) Then
                        promptMessage += $"当前自定义会话ID: {currentCustomId}" & Environment.NewLine
                    Else
                        promptMessage += "当前自定义会话ID: (未设置)" & Environment.NewLine
                    End If
                    promptMessage += Environment.NewLine & "留空则清除自定义会话ID"

                    ' 显示输入对话框（旧逻辑被封装为不可执行）
                    If False Then
                        Dim newCustomId As String = InputBox(promptMessage, "设置自定义会话ID", currentCustomId)

                    ' 如果用户点击了取消，InputBox 返回空字符串且用户没有输入任何内容
                    ' 我们需要区分用户点击取消和用户输入空字符串
                    If newCustomId <> currentCustomId Then
                        If String.IsNullOrEmpty(newCustomId) Then
                            ' 用户想要清除自定义会话ID
                            Dim storeId As String = Nothing
                            Try
                                Dim parentFolder = TryCast(CallByName(mailItem, "Parent", CallType.Get), Outlook.MAPIFolder)
                                If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                                    storeId = parentFolder.Store.StoreID
                                End If
                            Catch
                            End Try
                            If SetCustomConversationIdByEntryID(entryId, "", storeId) Then
                                MessageBox.Show("自定义会话ID已清除", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                ' 刷新邮件列表以反映更改
                                UpdateMailList(currentConversationId, entryId)
                            Else
                                Globals.ThisAddIn.ShowErrorWithConfig("清除自定义会话ID失败", "操作失败")
                            End If
                        Else
                            ' 用户输入了新的自定义会话ID
                            Dim storeId As String = Nothing
                            Try
                                Dim parentFolder = TryCast(CallByName(mailItem, "Parent", CallType.Get), Outlook.MAPIFolder)
                                If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                                    storeId = parentFolder.Store.StoreID
                                End If
                            Catch
                            End Try
                            If SetCustomConversationIdByEntryID(entryId, newCustomId.Trim(), storeId) Then
                                MessageBox.Show($"自定义会话ID已设置为: {newCustomId.Trim()}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                ' 刷新邮件列表以反映更改
                                UpdateMailList(currentConversationId, entryId)
                            Else
                                Globals.ThisAddIn.ShowErrorWithConfig("设置自定义会话ID失败", "操作失败")
                            End If
                        End If
                    End If
                    End If
                    ' 新逻辑：使用自定义窗口以区分“确定”和“取消/关闭”
                    Using dlg As New CustomConversationIdForm(originalConversationId, currentCustomId)
                        Dim owner As IWin32Window = Me.FindForm()
                        Dim result As DialogResult = dlg.ShowDialog(owner)

                        ' 取消或关闭：不做改动
                        If result <> DialogResult.OK Then
                            Return
                        End If

                        Dim newCustomId2 As String = If(dlg.EnteredId, String.Empty)

                        ' 未更改：不做改动
                        If String.Equals(newCustomId2, currentCustomId) Then
                            Return
                        End If

                        If String.IsNullOrWhiteSpace(newCustomId2) Then
                            ' 用户确认清除自定义会话ID
                            Dim storeId As String = Nothing
                            Try
                                Dim parentFolder = TryCast(CallByName(mailItem, "Parent", CallType.Get), Outlook.MAPIFolder)
                                If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                                    storeId = parentFolder.Store.StoreID
                                End If
                            Catch
                            End Try
                            If SetCustomConversationIdByEntryID(entryId, "", storeId) Then
                                MessageBox.Show("自定义会话ID已清除", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                UpdateMailList(currentConversationId, entryId)
                            Else
                                Globals.ThisAddIn.ShowErrorWithConfig("清除自定义会话ID失败", "操作失败")
                            End If
                        Else
                            ' 用户确认设置新的自定义会话ID
                            Dim storeId As String = Nothing
                            Try
                                Dim parentFolder = TryCast(CallByName(mailItem, "Parent", CallType.Get), Outlook.MAPIFolder)
                                If parentFolder IsNot Nothing AndAlso parentFolder.Store IsNot Nothing Then
                                    storeId = parentFolder.Store.StoreID
                                End If
                            Catch
                            End Try
                            Dim trimmedId As String = newCustomId2.Trim()
                            If SetCustomConversationIdByEntryID(entryId, trimmedId, storeId) Then
                                MessageBox.Show($"自定义会话ID已设置为: {trimmedId}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                UpdateMailList(currentConversationId, entryId)
                            Else
                                Globals.ThisAddIn.ShowErrorWithConfig("设置自定义会话ID失败", "操作失败")
                            End If
                        End If
                    End Using
                Finally
                    OutlookMyList.Utils.OutlookUtils.SafeReleaseComObject(mailItem)
                End Try
            Else
                MessageBox.Show("无法获取邮件项", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"CustomConversationId_Click error: {ex.Message}")
            MessageBox.Show($"设置自定义会话ID时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    ''' <summary>
    ''' 获取指定邮件所属会话中的所有邮件EntryID列表
    ''' </summary>
    ''' <param name="entryID">邮件的EntryID</param>
    ''' <returns>会话中所有邮件的EntryID列表</returns>
    Private Function GetAllMailsInConversation(entryID As String) As List(Of String)
        Dim result As New List(Of String)
        Dim mailItem As Object = Nothing
        Dim conversation As Outlook.Conversation = Nothing
        Dim table As Outlook.Table = Nothing

        Try
            ' 获取邮件项
            mailItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryID)
            If mailItem Is Nothing Then
                Debug.WriteLine($"GetAllMailsInConversation: 无法获取邮件项 {entryID}")
                Return result
            End If

            ' 获取会话对象
            If TypeOf mailItem Is Outlook.MailItem Then
                conversation = DirectCast(mailItem, Outlook.MailItem).GetConversation()
            ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                conversation = DirectCast(mailItem, Outlook.AppointmentItem).GetConversation()
            ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                conversation = DirectCast(mailItem, Outlook.MeetingItem).GetConversation()
            End If

            If conversation Is Nothing Then
                Debug.WriteLine($"GetAllMailsInConversation: 无法获取会话对象")
                ' 如果没有会话，只返回当前邮件的EntryID
                result.Add(entryID)
                Return result
            End If

            ' 获取会话中的所有邮件
            table = conversation.GetTable()
            table.Columns.RemoveAll()
            ' 使用PR_ENTRYID获取长格式EntryID
            table.Columns.Add("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102")

            Do Until table.EndOfTable
                Dim row As Outlook.Row = Nothing
                Try
                    row = table.GetNextRow()
                    Dim entryIdStr As String = ConvertEntryIDToString(row("http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"))
                    If Not String.IsNullOrEmpty(entryIdStr) Then
                        result.Add(entryIdStr)
                    End If
                Finally
                    If row IsNot Nothing Then
                        Runtime.InteropServices.Marshal.ReleaseComObject(row)
                    End If
                End Try
            Loop

            Debug.WriteLine($"GetAllMailsInConversation: 找到 {result.Count} 个邮件")

        Catch ex As System.Exception
            Debug.WriteLine($"GetAllMailsInConversation error: {ex.Message}")
        Finally
            ' 释放COM对象
            If table IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(table)
            End If
            If conversation IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(conversation)
            End If
            If mailItem IsNot Nothing Then
                Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
            End If
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 独立的联系人来往邮件信息处理方法，不依赖按钮状态
    ''' </summary>
    Private Sub ShowMailHistoryInfo()
        Debug.WriteLine("=== ShowContactMailHistoryInfo 开始 ===")
        Try
            ' 检查lvMails是否为Nothing
            If lvMails Is Nothing Then
                Debug.WriteLine("ShowContactMailHistoryInfo: lvMails is Nothing")
                Return
            End If

            If lvMails.SelectedItems.Count = 0 Then
                Debug.WriteLine("没有选中的邮件项")
                Return
            End If

            Debug.WriteLine($"选中邮件数量: {lvMails.SelectedItems.Count}")

            ' 直接从选中邮件获取发件人信息
            Dim selectedItem = lvMails.SelectedItems(0)
            Dim entryId As String = ConvertEntryIDToString(selectedItem.Tag)
            Dim senderName As String = If(selectedItem.SubItems.Count > 2, selectedItem.SubItems(2).Text, "未知发件人")

            Debug.WriteLine($"EntryID: {entryId}")
            Debug.WriteLine($"发件人姓名: {senderName}")
            Debug.WriteLine($"SubItems数量: {selectedItem.SubItems.Count}")

            ' 获取发件人邮箱地址
            Dim senderEmail As String = ""
            Try
                Debug.WriteLine("开始获取邮件对象...")
                Dim mailItem = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(entryId)
                Debug.WriteLine($"邮件对象获取结果: {If(mailItem Is Nothing, "Nothing", "成功")}")

                If mailItem IsNot Nothing Then
                    Debug.WriteLine($"邮件对象类型: {mailItem.GetType().Name}")
                    If TypeOf mailItem Is Outlook.MailItem Then
                        senderEmail = DirectCast(mailItem, Outlook.MailItem).SenderEmailAddress
                        Debug.WriteLine($"MailItem发件人邮箱: {senderEmail}")
                    ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                        senderEmail = DirectCast(mailItem, Outlook.MeetingItem).SenderEmailAddress
                        Debug.WriteLine($"MeetingItem发件人邮箱: {senderEmail}")
                    Else
                        Debug.WriteLine($"未知的邮件对象类型: {mailItem.GetType().Name}")
                    End If
                Else
                    Debug.WriteLine("SafeGetItemFromID 返回 Nothing")
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"获取发件人邮箱失败: {ex.Message}")
                Debug.WriteLine($"异常堆栈: {ex.StackTrace}")
                ' 不直接返回，而是使用发件人姓名作为邮箱地址的替代
                senderEmail = senderName
            End Try

            Debug.WriteLine($"最终获取的发件人邮箱: {senderEmail}")

            If String.IsNullOrEmpty(senderEmail) Then
                Debug.WriteLine("发件人邮箱为空，使用默认值")
                senderEmail = "unknown@unknown.com"
            End If

            ' 切换到来往邮件标签页
            Debug.WriteLine("开始切换到来往邮件标签页...")
            Dim foundMailInteractionTab As Boolean = False
            For Each tabPage As TabPage In tabControl.TabPages
                If tabPage.Text = "来往邮件" Then
                    tabControl.SelectedTab = tabPage
                    foundMailInteractionTab = True
                    Debug.WriteLine("成功切换到来往邮件标签页")
                    Exit For
                End If
            Next

            If Not foundMailInteractionTab Then
                Debug.WriteLine("警告：未找到来往邮件标签页")
            End If

            ' 确保找到来往邮件标签页中的TreeView
            Debug.WriteLine("开始查找来往邮件标签页中的TreeView...")
            Dim mailInteractionTreeView As TreeView = Nothing
            For Each tabPage As TabPage In tabControl.TabPages
                If tabPage.Text = "来往邮件" Then
                    ' 在来往邮件标签页中查找TreeView控件
                    For Each control As Control In tabPage.Controls
                        Dim treeView As TreeView = FindTreeViewInControl(control)
                        If treeView IsNot Nothing Then
                            mailInteractionTreeView = treeView
                            Debug.WriteLine("找到来往邮件标签页中的TreeView")
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next

            ' 确保使用正确的TreeView
            If mailInteractionTreeView IsNot Nothing Then
                contactInfoTree = mailInteractionTreeView
                Debug.WriteLine("成功设置contactInfoTree为来往邮件标签页的TreeView")
            Else
                Debug.WriteLine("警告：未找到来往邮件标签页中的TreeView")
            End If

            ' 直接调用联系人来往邮件信息数据获取方法，传递发件人信息
            Debug.WriteLine($"开始调用GetContactInfoTreeHandlerWithSender，发件人: {senderName}, 邮箱: {senderEmail}")
            GetContactInfoTreeHandlerWithSender(senderName, senderEmail)
            Debug.WriteLine("=== ShowContactMailHistoryInfo 结束 ===")
        Catch ex As System.Exception
            Debug.WriteLine($"ShowContactInfo error: {ex.Message}")
            Debug.WriteLine($"ShowContactInfo 异常堆栈: {ex.StackTrace}")
        End Try
    End Sub

    Private Sub GotoMailHistory_Click(sender As Object, e As EventArgs)
        ' 直接调用邮件历史信息处理方法（保持在UI线程）
        ShowMailHistoryInfo()
    End Sub

    Private Sub MailContextMenu_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs)
        Try
            ' 获取当前选中邮件的发件人姓名
            Dim senderName As String = "联系人"
            If lvMails.SelectedItems.Count > 0 Then
                Dim selectedItem = lvMails.SelectedItems(0)
                If selectedItem.SubItems.Count > 2 Then
                    senderName = selectedItem.SubItems(2).Text ' 发件人在第三列（索引2）
                End If
            End If

            ' 更新右键菜单中的菜单项
            Dim contextMenu As ContextMenuStrip = TryCast(sender, ContextMenuStrip)
            If contextMenu IsNot Nothing Then
                For Each item As ToolStripItem In contextMenu.Items
                    Dim menuItem As ToolStripMenuItem = TryCast(item, ToolStripMenuItem)
                    If menuItem IsNot Nothing Then
                        If Not menuItem.Text.StartsWith("自定义") Then
                            If menuItem.Text.Contains("待办邮件") Then
                                menuItem.Text = $"{senderName}待办邮件"
                            ElseIf menuItem.Text.Contains("往来邮件") Then
                                menuItem.Text = $"{senderName}往来邮件"
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"MailContextMenu_Opening error: {ex.Message}")
        End Try
    End Sub

    Private Sub PendingMails_Click(sender As Object, e As EventArgs)
        Task.Run(Sub() PendingMailsAsync())
    End Sub

    Private Sub HideContextMenuOnMouseDown(sender As Object, e As MouseEventArgs)
        Try
            If mailContextMenu IsNot Nothing AndAlso mailContextMenu.Visible Then
                mailContextMenu.Hide()
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub CustomMailHistory_Click(sender As Object, e As EventArgs)
        Dim email As String = ""
        Dim senderName As String = ""
        If Not PromptEmailAddress(email, senderName) Then Exit Sub
        Try
            Dim foundMailInteractionTab As Boolean = False
            For Each tabPage As TabPage In tabControl.TabPages
                If tabPage.Text = "来往邮件" Then
                    tabControl.SelectedTab = tabPage
                    foundMailInteractionTab = True
                    Exit For
                End If
            Next
            Dim mailInteractionTreeView As TreeView = Nothing
            For Each tabPage As TabPage In tabControl.TabPages
                If tabPage.Text = "来往邮件" Then
                    For Each control As Control In tabPage.Controls
                        Dim treeView As TreeView = FindTreeViewInControl(control)
                        If Not treeView Is Nothing Then
                            mailInteractionTreeView = treeView
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
            If Not mailInteractionTreeView Is Nothing Then
                contactInfoTree = mailInteractionTreeView
            End If
            GetContactInfoTreeHandlerWithSender(senderName, email)
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub CustomPendingMails_Click(sender As Object, e As EventArgs)
        Dim email As String = ""
        Dim senderName As String = ""
        If Not PromptEmailAddress(email, senderName) Then Exit Sub
        Try
            Dim foundTab As Boolean = False
            If tabControl IsNot Nothing Then
                For Each tabPage As TabPage In tabControl.TabPages
                    If tabPage.Text = "待办邮件" Then
                        tabControl.SelectedTab = tabPage
                        foundTab = True
                        Exit For
                    End If
                Next
            End If
            If Not foundTab Then Exit Sub
            ShowPendingMailsForEmail(email, senderName)
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub ShowPendingMailsForEmail(email As String, senderName As String)
        Dim pendingMailListView As ListView = GetPendingMailListView()
        If pendingMailListView Is Nothing Then Exit Sub
        If Me.InvokeRequired Then
            Me.Invoke(Sub()
                          pendingMailListView.Items.Clear()
                          Dim loadingItem As New ListViewItem("正在收集联系人任务邮件...")
                          loadingItem.SubItems.Add("")
                          loadingItem.SubItems.Add("")
                          pendingMailListView.Items.Add(loadingItem)
                      End Sub)
        Else
            pendingMailListView.Items.Clear()
            Dim loadingItem As New ListViewItem("正在收集联系人任务邮件...")
            loadingItem.SubItems.Add("")
            loadingItem.SubItems.Add("")
            pendingMailListView.Items.Add(loadingItem)
        End If
        Task.Run(Sub()
                     Dim result = GetTaskMailsForEmail(email, senderName)
                     If Me.InvokeRequired Then
                         Me.Invoke(Sub() PopulateContactTaskMails(result))
                     Else
                         PopulateContactTaskMails(result)
                     End If
                 End Sub)
    End Sub

    Private Function GetTaskMailsForEmail(email As String, senderName As String) As List(Of Object)
        Dim taskMails As New List(Of Object)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim outlookNameSpace As Microsoft.Office.Interop.Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            taskMails = GetTaskMailsUsingTable(outlookApp, outlookNameSpace, email, senderName)
        Catch ex As System.Exception
        End Try
        Return taskMails
    End Function

    Private Function PromptEmailAddress(ByRef email As String, ByRef displayName As String) As Boolean
        Dim input As String = ""
        While True
            Dim f As New Form()
            Dim tb As New TextBox()
            Dim lbl As New Label()
            Dim okBtn As New Button()
            Dim cancelBtn As New Button()
            f.Text = "输入邮件地址"
            f.FormBorderStyle = FormBorderStyle.FixedDialog
            f.StartPosition = FormStartPosition.CenterParent
            f.MinimizeBox = False
            f.MaximizeBox = False
            f.ClientSize = New Size(360, 140)
            lbl.Text = "邮件地址："
            lbl.Location = New Point(16, 20)
            lbl.Size = New Size(320, 20)
            tb.Location = New Point(16, 45)
            tb.Size = New Size(320, 24)
            okBtn.Text = "确认"
            okBtn.Location = New Point(170, 90)
            okBtn.Size = New Size(75, 28)
            okBtn.DialogResult = DialogResult.OK
            cancelBtn.Text = "取消"
            cancelBtn.Location = New Point(260, 90)
            cancelBtn.Size = New Size(75, 28)
            cancelBtn.DialogResult = DialogResult.Cancel
            f.AcceptButton = okBtn
            f.CancelButton = cancelBtn
            f.Controls.Add(lbl)
            f.Controls.Add(tb)
            f.Controls.Add(okBtn)
            f.Controls.Add(cancelBtn)
            Dim dr = f.ShowDialog(Me)
            If dr = DialogResult.OK Then
                input = tb.Text.Trim()
            Else
                f.Dispose()
                Return False
            End If
            f.Dispose()
            If IsValidEmail(input) Then
                email = ExtractEmailAddress(input)
                displayName = ExtractDisplayName(input)
                Return True
            Else
                MessageBox.Show("请输入有效的邮件地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End While
        Return False
    End Function

    Private Function IsValidEmail(addr As String) As Boolean
        If String.IsNullOrWhiteSpace(addr) Then Return False
        Dim pure As String = ExtractEmailAddress(addr)
        If String.IsNullOrWhiteSpace(pure) Then Return False
        Try
            Dim m = New System.Net.Mail.MailAddress(pure)
            Return String.Equals(m.Address, pure, StringComparison.OrdinalIgnoreCase)
        Catch ex As System.Exception
            Return False
        End Try
    End Function

    Private Function ExtractEmailAddress(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return String.Empty
        Dim t As String = text.Trim()
        Dim lt As Integer = t.IndexOf("<"c)
        Dim gt As Integer = t.LastIndexOf(">"c)
        If lt >= 0 AndAlso gt > lt Then
            Dim inner As String = t.Substring(lt + 1, gt - lt - 1).Trim()
            Return inner
        End If
        Try
            Dim m = New System.Net.Mail.MailAddress(t)
            Return m.Address
        Catch ex As System.Exception
            Return String.Empty
        End Try
    End Function

    Private Function ExtractDisplayName(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return String.Empty
        Dim t As String = text.Trim()
        Dim lt As Integer = t.IndexOf("<"c)
        If lt > 0 Then
            Dim namePart As String = t.Substring(0, lt).Trim()
            If namePart.StartsWith(""""c) AndAlso namePart.EndsWith(""""c) Then
                namePart = namePart.Substring(1, namePart.Length - 2)
            End If
            If Not String.IsNullOrWhiteSpace(namePart) Then Return namePart
        End If
        Try
            Dim m = New System.Net.Mail.MailAddress(t)
            If Not String.IsNullOrWhiteSpace(m.DisplayName) Then Return m.DisplayName.Trim()
        Catch ex As System.Exception
        End Try
        Return t
    End Function

    Private Sub PendingMailsAsync()
        Dim senderName As String = "联系人"
        Try
            ' 在UI线程上检查选中项
            Dim hasSelectedItems As Boolean = False
            Dim selectedEntryId As String = ""

            Me.Invoke(Sub()
                          ' 检查lvMails是否为Nothing
                          If lvMails Is Nothing Then
                              Debug.WriteLine("ContactTasksAsync: lvMails is Nothing")
                              Return
                          End If

                          If lvMails.SelectedItems.Count = 0 Then
                              MessageBox.Show("请先选择一封邮件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                              Return
                          End If
                          hasSelectedItems = True
                          Dim selectedItem = lvMails.SelectedItems(0)
                          selectedEntryId = ConvertEntryIDToString(selectedItem.Tag)
                          If selectedItem.SubItems.Count > 2 Then
                              senderName = selectedItem.SubItems(2).Text
                          End If
                      End Sub)

            If Not hasSelectedItems Then
                Return
            End If

            ' 在UI线程上切换到待办邮件标签页
            Dim foundTab As Boolean = False
            Me.Invoke(Sub()
                          If tabControl IsNot Nothing Then
                              For Each tabPage As TabPage In tabControl.TabPages
                                  If tabPage.Text = "待办邮件" Then
                                      tabControl.SelectedTab = tabPage
                                      foundTab = True
                                      Exit For
                                  End If
                              Next
                          End If
                      End Sub)

            If Not foundTab Then
                Me.Invoke(Sub() MessageBox.Show("未找到待办邮件页面", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error))
                Return
            End If

            ' 异步调用显示联系人任务邮件的方法
            ShowPendingMailsAsync()

        Catch ex As System.Exception
            Debug.WriteLine($"显示{senderName}待办邮件时出错: {ex.Message}")
            Me.Invoke(Sub() MessageBox.Show($"显示{senderName}待办邮件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error))
        End Try
    End Sub

    ' 获取待办邮件标签页中的ListView控件
    Private Function GetPendingMailListView() As ListView
        Try
            If tabControl IsNot Nothing Then
                For Each tabPage As TabPage In tabControl.TabPages
                    If tabPage.Text = "待办邮件" Then
                        ' 在待办邮件标签页中查找ListView控件
                        For Each control As Control In tabPage.Controls
                            If TypeOf control Is ListView Then
                                Return DirectCast(control, ListView)
                            End If
                        Next
                    End If
                Next
            End If
        Catch ex As System.Exception
            ' 忽略错误
        End Try
        Return Nothing
    End Function

    Private Async Sub ShowPendingMailsAsync()
        Try
            ' 查找待办邮件标签页中的ListView
            Dim pendingMailListView As ListView = Nothing
            If Me.InvokeRequired Then
                Me.Invoke(Sub() pendingMailListView = GetPendingMailListView())
            Else
                pendingMailListView = GetPendingMailListView()
            End If

            If pendingMailListView Is Nothing Then
                If Me.InvokeRequired Then
                    Me.Invoke(Sub() MessageBox.Show("待办邮件列表未找到", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error))
                Else
                    MessageBox.Show("待办邮件列表未找到", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return
            End If

            ' 显示进度指示器
            If Me.InvokeRequired Then
                Me.Invoke(Sub()
                              pendingMailListView.Items.Clear()
                              Dim loadingItem As New ListViewItem("正在收集联系人任务邮件...")
                              loadingItem.SubItems.Add("")
                              loadingItem.SubItems.Add("")
                              pendingMailListView.Items.Add(loadingItem)
                          End Sub)
            Else
                pendingMailListView.Items.Clear()
                Dim loadingItem As New ListViewItem("正在收集联系人任务邮件...")
                loadingItem.SubItems.Add("")
                loadingItem.SubItems.Add("")
                pendingMailListView.Items.Add(loadingItem)
            End If

            ' 异步收集联系人信息和任务邮件
            Dim result = Await Task.Run(Function() GetContactTaskMailsData())

            ' 在UI线程中更新显示
            If Me.InvokeRequired Then
                Me.Invoke(Sub() PopulateContactTaskMails(result))
            Else
                PopulateContactTaskMails(result)
            End If
        Catch ex As System.Exception
            If Me.InvokeRequired Then
                Me.Invoke(Sub() MessageBox.Show($"显示联系人任务邮件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error))
            Else
                MessageBox.Show($"显示联系人任务邮件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    Private Function GetContactTaskMailsData() As List(Of Object)
        Dim taskMails As New List(Of Object)
        Dim entryID As String = ""

        ' 获取选中邮件的EntryID
        If Me.InvokeRequired Then
            Me.Invoke(Sub()
                          If lvMails.SelectedItems.Count > 0 Then
                              entryID = lvMails.SelectedItems(0).Tag?.ToString()
                          End If
                      End Sub)
        Else
            If lvMails.SelectedItems.Count > 0 Then
                entryID = lvMails.SelectedItems(0).Tag?.ToString()
            End If
        End If

        If String.IsNullOrEmpty(entryID) Then
            Return taskMails
        End If

        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim outlookNameSpace As Microsoft.Office.Interop.Outlook.NameSpace = outlookApp.GetNamespace("MAPI")

            ' 获取选中邮件的发件人信息
            Dim selectedMail As Outlook.MailItem = Nothing
            Dim senderEmail As String = ""
            Dim senderName As String = ""

            Try
                selectedMail = TryCast(outlookNameSpace.GetItemFromID(entryID), Outlook.MailItem)
                If selectedMail IsNot Nothing Then
                    senderEmail = selectedMail.SenderEmailAddress
                    senderName = selectedMail.SenderName

                    ' 如果是Exchange格式，尝试获取SMTP地址
                    If Not String.IsNullOrEmpty(senderEmail) AndAlso senderEmail.StartsWith("/O=") Then
                        Try
                            If selectedMail.Sender IsNot Nothing Then
                                Dim exchangeUser = selectedMail.Sender.GetExchangeUser()
                                If exchangeUser IsNot Nothing Then
                                    senderEmail = exchangeUser.PrimarySmtpAddress
                                End If
                                If String.IsNullOrEmpty(senderEmail) Then
                                    senderEmail = selectedMail.Sender.Address
                                End If
                            End If
                        Catch
                            ' 如果获取SMTP地址失败，使用原始地址
                        End Try
                    End If
                End If
            Finally
                If selectedMail IsNot Nothing Then
                    Runtime.InteropServices.Marshal.ReleaseComObject(selectedMail)
                End If
            End Try

            If String.IsNullOrEmpty(senderEmail) Then
                Return taskMails
            End If

            ' 使用优化的搜索方法
            taskMails = GetTaskMailsUsingTable(outlookApp, outlookNameSpace, senderEmail, senderName)
        Catch ex As System.Exception
            ' 忽略错误
        End Try

        Return taskMails
    End Function

    ' 使用GetTable()和DASL查询的优化搜索方法
    Private Function GetTaskMailsUsingTable(outlookApp As Microsoft.Office.Interop.Outlook.Application,
                                           outlookNameSpace As Outlook.NameSpace,
                                           senderEmail As String,
                                           senderName As String) As List(Of Object)
        Dim taskMails As New List(Of Object)

        Try
            ' 获取指定的核心文件夹 - 使用现有的GetAllMailFolders方法
            Dim folders As New List(Of Outlook.Folder)

            ' 从根文件夹开始搜索所有核心文件夹
            Try
                Dim rootFolder As Outlook.Folder = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
                GetAllMailFolders(rootFolder, folders)

                ' 如果没有找到文件夹，至少添加收件箱作为备选
                If folders.Count = 0 Then
                    folders.Add(DirectCast(outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), Outlook.Folder))
                End If
            Catch ex As System.Exception
                ' 如果出错，至少添加收件箱
                Try
                    folders.Add(DirectCast(outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), Outlook.Folder))
                Catch
                    ' 忽略错误
                End Try
            End Try

            For Each folder As Outlook.Folder In folders
                Try

                    ' 使用官方文档推荐的DASL过滤器筛选标记为任务的邮件
                    ' 参考: https://learn.microsoft.com/en-us/office/vba/outlook/concepts/electronic-business-cards/setting-items-for-follow-up
                    Dim taskFilter As String = "@SQL=" & Chr(34) & "http://schemas.microsoft.com/mapi/proptag/0x0E2B0003" & Chr(34) & " = 1"

                    Dim table As Outlook.Table = folder.GetTable(taskFilter, Outlook.OlTableContents.olUserItems)

                    ' 添加需要的列（移除不支持的TaskSubject等任务特定列）
                    table.Columns.Add("Subject")
                    table.Columns.Add("ReceivedTime")
                    table.Columns.Add("SenderName")
                    table.Columns.Add("SenderEmailAddress")
                    table.Columns.Add("EntryID")
                    table.Columns.Add("FlagStatus")
                    ' 注意：TaskSubject, TaskDueDate, TaskCompletedDate 在Table中不被支持
                    ' 需要通过获取实际邮件项来检查这些属性

                    Dim folderTaskCount As Integer = 0

                    ' 遍历Table中的行
                    While Not table.EndOfTable
                        Dim row As Outlook.Row = table.GetNextRow()

                        Try
                            ' 获取发件人邮箱地址
                            Dim mailSenderEmail As String = If(row("SenderEmailAddress"), "").ToString()

                            ' 处理Exchange格式地址 - 需要获取实际邮件项来转换
                            If Not String.IsNullOrEmpty(mailSenderEmail) AndAlso mailSenderEmail.StartsWith("/O=") Then
                                Try
                                    Dim mailEntryID As String = row("EntryID").ToString()
                                    Dim mail As Outlook.MailItem = TryCast(outlookNameSpace.GetItemFromID(mailEntryID), Outlook.MailItem)
                                    If mail IsNot Nothing Then
                                        If mail.Sender IsNot Nothing Then
                                            Dim exchangeUser = mail.Sender.GetExchangeUser()
                                            If exchangeUser IsNot Nothing Then
                                                Dim smtpAddress As String = exchangeUser.PrimarySmtpAddress
                                                If Not String.IsNullOrEmpty(smtpAddress) Then
                                                    mailSenderEmail = smtpAddress
                                                Else
                                                    mailSenderEmail = mail.Sender.Address
                                                End If
                                            End If
                                        End If
                                        Runtime.InteropServices.Marshal.ReleaseComObject(mail)
                                    End If
                                Catch
                                    ' 使用原始地址
                                End Try
                            End If

                            ' 检查发件人是否匹配
                            If String.Equals(mailSenderEmail, senderEmail, StringComparison.OrdinalIgnoreCase) Then
                                ' 需要获取实际邮件项来检查任务完成状态
                                Try
                                    Dim mailEntryID As String = row("EntryID").ToString()
                                    Dim mail As Outlook.MailItem = TryCast(outlookNameSpace.GetItemFromID(mailEntryID), Outlook.MailItem)
                                    If mail IsNot Nothing Then
                                        ' 检查任务是否未完成
                                        ' TaskCompletedDate = 4501-01-01 表示任务未完成
                                        Dim isTaskIncomplete As Boolean = mail.IsMarkedAsTask AndAlso
                                            (mail.TaskCompletedDate = Nothing OrElse
                                             mail.TaskCompletedDate.Year = 4501)

                                        If isTaskIncomplete Then
                                            folderTaskCount += 1

                                            ' 安全获取各个字段值
                                            Dim subject As String = If(row("Subject"), "").ToString()
                                            Dim receivedTime As Object = row("ReceivedTime")
                                            Dim flagStatus As String = If(row("FlagStatus"), "").ToString()
                                            Dim entryID As String = row("EntryID").ToString()

                                            taskMails.Add(New With {
                                                .Subject = subject,
                                                .ReceivedTime = receivedTime,
                                                .SenderName = senderName,
                                                .SenderEmailAddress = mailSenderEmail,
                                                .EntryID = entryID,
                                                .TaskSubject = mail.TaskSubject,
                                                .TaskDueDate = mail.TaskDueDate,
                                                .FlagStatus = flagStatus,
                                                .IsMarkedAsTask = True,
                                                .TaskCompletedDate = mail.TaskCompletedDate
                                            })
                                        End If
                                        Runtime.InteropServices.Marshal.ReleaseComObject(mail)
                                    End If
                                Catch mailEx As System.Exception
                                    ' 忽略获取邮件项的错误
                                End Try
                            End If
                        Catch ex As System.Exception
                            ' 忽略处理行数据的错误
                        End Try
                    End While

                    ' 清理Table对象
                    Runtime.InteropServices.Marshal.ReleaseComObject(table)

                Catch ex As System.Exception
                    ' 忽略搜索文件夹的错误
                Finally
                    Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                End Try
            Next

        Catch ex As System.Exception
            ' 忽略搜索错误
        End Try

        Return taskMails
    End Function

    Private Sub PopulateContactTaskMails(taskMails As List(Of Object))
        ' 查找待办邮件标签页中的ListView
        Dim pendingMailListView As ListView = GetPendingMailListView()
        If pendingMailListView Is Nothing Then
            Return
        End If

        ' 确保ListView绑定了正确的事件处理程序
        Try
            ' 移除可能存在的旧事件处理程序
            RemoveHandler pendingMailListView.Click, AddressOf MailHistory_Click
            RemoveHandler pendingMailListView.DoubleClick, AddressOf MailHistory_DoubleClick
        Catch
            ' 忽略移除事件处理程序时的错误
        End Try

        ' 添加事件处理程序
        AddHandler pendingMailListView.Click, AddressOf MailHistory_Click
        AddHandler pendingMailListView.DoubleClick, AddressOf MailHistory_DoubleClick

        Try

            pendingMailListView.BeginUpdate()
            pendingMailListView.Items.Clear()

            ' 获取发件人名字用于列标题
            Dim senderName As String = "未知发件人"
            If taskMails.Count > 0 Then
                senderName = If(taskMails(0).SenderName?.ToString(), "未知发件人")
            End If

            If taskMails.Count > 0 Then
                ' 更新第一列标题显示发件人名字
                If pendingMailListView.Columns.Count > 0 Then
                    pendingMailListView.Columns(0).Text = $"时间 ({senderName})"
                End If

                ' 添加标题
                Dim headerItem As New ListViewItem($"{senderName}待办邮件")
                headerItem.SubItems.Add($"共找到 {taskMails.Count} 封任务邮件")
                headerItem.SubItems.Add("")
                headerItem.BackColor = SystemColors.Control
                headerItem.Font = New Font(headerItem.Font.Name, headerItem.Font.Size, FontStyle.Bold)
                pendingMailListView.Items.Add(headerItem)

                ' 添加任务邮件
                For Each taskMail In taskMails
                    ' 安全的日期格式化
                    Dim dateText As String = ""
                    Try
                        If taskMail.ReceivedTime IsNot Nothing Then
                            Dim receivedTime As DateTime = Convert.ToDateTime(taskMail.ReceivedTime)
                            dateText = receivedTime.ToString("yyyy-MM-dd HH:mm")
                        Else
                            dateText = "未知时间"
                        End If
                    Catch ex As System.Exception
                        dateText = "日期格式错误"
                        Debug.WriteLine($"日期格式化错误: {ex.Message}")
                    End Try

                    ' 获取Due Date（如果有的话）
                    Dim dueDateText As String = "无截止日期"
                    Try
                        If taskMail.TaskDueDate IsNot Nothing Then
                            Dim dueDate As DateTime = Convert.ToDateTime(taskMail.TaskDueDate)
                            dueDateText = dueDate.ToString("yyyy-MM-dd")
                        End If
                    Catch ex As System.Exception
                        dueDateText = "无截止日期"
                    End Try

                    ' 第一列显示时间，第二列显示主题，第三列显示Due Date
                    Dim item As New ListViewItem(dateText)
                    item.SubItems.Add(If(String.IsNullOrEmpty(taskMail.Subject), "(无主题)", taskMail.Subject.ToString()))
                    item.SubItems.Add(dueDateText)
                    item.Tag = If(taskMail.EntryID IsNot Nothing, taskMail.EntryID.ToString(), "")
                    item.BackColor = currentBackColor  ' 使用当前主题背景色
                    pendingMailListView.Items.Add(item)
                Next
            Else
                ' 没有找到任务邮件
                Dim noTaskItem As New ListViewItem($"{senderName}待办邮件")
                noTaskItem.SubItems.Add("该联系人没有标记为任务的邮件")
                noTaskItem.SubItems.Add("")

                ' 应用主题到无任务项目
                ApplyThemeToListViewItem(noTaskItem)

                pendingMailListView.Items.Add(noTaskItem)
            End If

        Finally
            If pendingMailListView IsNot Nothing Then
                pendingMailListView.EndUpdate()
            End If
        End Try
    End Sub



    ' 邮件历史和待办邮件ListView点击事件处理程序
    Private Sub MailHistory_Click(sender As Object, e As EventArgs)
        Try
            ' 确定是哪个ListView触发的事件
            Dim sourceListView As ListView = TryCast(sender, ListView)
            If sourceListView Is Nothing Then
                Debug.WriteLine("MailHistory_Click: sender is not a ListView")
                Return
            End If

            Debug.WriteLine($"MailHistory_Click: ListView名称={sourceListView.Name}, 选中项数量={sourceListView.SelectedItems.Count}")
            Debug.WriteLine($"MailHistory_Click: suppressWebViewUpdate当前值={suppressWebViewUpdate}")

            If sourceListView.SelectedItems.Count = 0 Then 
                Debug.WriteLine("MailHistory_Click: 没有选中项，退出")
                Return
            End If

            Dim item = sourceListView.SelectedItems(0)
            Dim mailId As String = TryCast(item.Tag, String)

            ' 只处理有EntryID的邮件项
            If String.IsNullOrEmpty(mailId) Then
                Debug.WriteLine("MailHistory_Click: EntryID为空")
                Return
            End If

            ' 始终更新高亮，不受suppressWebViewUpdate影响
            If Not mailId.Equals(currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                Debug.WriteLine($"MailHistory_Click: 邮件ID不同，准备更新高亮。旧邮件ID = {currentMailEntryID}, 新邮件ID = {mailId}")
                Dim oldMailId As String = currentMailEntryID
                currentMailEntryID = mailId
                UpdateHighlightByEntryID(oldMailId, mailId)
                sourceListView.Refresh() ' 强制刷新ListView以确保高亮立即显示
                Debug.WriteLine("MailHistory_Click: UpdateHighlightByEntryID 已调用。")

                ' 只有在非抑制模式下才加载WebView内容
                Debug.WriteLine($"MailHistory_Click: suppressWebViewUpdate = {suppressWebViewUpdate}")
                If suppressWebViewUpdate = 0 Then
                    Debug.WriteLine($"MailHistory_Click: 开始加载WebView内容，邮件ID = {mailId}")
                    Me.BeginInvoke(New Action(Of String)(AddressOf LoadMailContentDeferred), mailId)
                    
                    ' 立即应用主题到WebBrowser控件本身，确保主题色正确
                    Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
                    UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)
                    Debug.WriteLine($"MailHistory_Click: 已立即应用主题色到WebBrowser控件")
                Else
                    Debug.WriteLine($"MailHistory_Click: WebView更新被抑制，suppressWebViewUpdate = {suppressWebViewUpdate}")
                End If
            Else
                Debug.WriteLine($"MailHistory_Click: 邮件ID相同，跳过更新")
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"MailHistory_Click error: {ex.Message}")
        End Try
    End Sub

    ' 邮件历史和待办邮件ListView双击事件处理程序
    Private Sub MailHistory_DoubleClick(sender As Object, e As EventArgs)
        Try
            ' 确定是哪个ListView触发的事件
            Dim sourceListView As ListView = TryCast(sender, ListView)
            If sourceListView Is Nothing Then
                Debug.WriteLine("MailHistory_DoubleClick: sender is not a ListView")
                Return
            End If

            If sourceListView.SelectedItems.Count = 0 Then Return

            Dim item = sourceListView.SelectedItems(0)
            Dim entryId = TryCast(item.Tag, String)

            ' 只处理有EntryID的邮件项
            If Not String.IsNullOrEmpty(entryId) Then
                ' 更新高亮显示
                Dim oldEntryId As String = currentHighlightEntryID
                UpdateHighlightByEntryID(oldEntryId, entryId)
                
                ' 双击时在Outlook中打开邮件
                SafeOpenOutlookMail(entryId)
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            ' COM异常静默处理，只记录调试信息，绝不抛出
            Debug.WriteLine($"MailHistory_DoubleClick COM异常 (HRESULT: {ex.HResult:X8}): {ex.Message}")
        Catch ex As System.Exception
            ' 其他异常也静默处理，绝不抛出，避免弹出错误对话框
            Debug.WriteLine($"MailHistory_DoubleClick 异常: {ex.Message}")
        End Try
    End Sub

    ' 在WebView中显示邮件内容
    Private Sub DisplayMailInWebView(entryId As String)
        Try
            Me.BeginInvoke(New Action(Of String)(AddressOf LoadMailContentDeferred), entryId)
            Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
            UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)
        Catch ex As System.Exception
            Debug.WriteLine($"DisplayMailInWebView error: {ex.Message}")
        End Try
    End Sub

    ' 获取简洁的发件人姓名
    Private Function GetFriendlySenderName(mailItem As Outlook.MailItem) As String
        Try
            ' 首先尝试获取Sender的DisplayName
            If mailItem.Sender IsNot Nothing AndAlso Not String.IsNullOrEmpty(mailItem.Sender.Name) Then
                Return mailItem.Sender.Name
            End If

            ' 如果Sender为空，尝试从SenderName中提取
            If Not String.IsNullOrEmpty(mailItem.SenderName) Then
                Dim senderName As String = mailItem.SenderName

                ' 如果是Exchange DN格式（以/O=开头），尝试提取CN部分
                If senderName.StartsWith("/O=") Then
                    Dim cnIndex As Integer = senderName.LastIndexOf("/CN=")
                    If cnIndex >= 0 Then
                        Dim cnPart As String = senderName.Substring(cnIndex + 4)
                        ' 移除可能的后缀
                        Dim dashIndex As Integer = cnPart.IndexOf("-")
                        If dashIndex >= 0 Then
                            cnPart = cnPart.Substring(0, dashIndex)
                        End If
                        Return cnPart
                    End If
                End If

                ' 如果不是Exchange DN格式，直接返回
                Return senderName
            End If

            ' 最后尝试SenderEmailAddress
            If Not String.IsNullOrEmpty(mailItem.SenderEmailAddress) Then
                Dim emailParts() As String = mailItem.SenderEmailAddress.Split("@"c)
                If emailParts.Length > 0 Then
                    Return emailParts(0)
                End If
            End If

            Return "未知发件人"
        Catch ex As System.Exception
            Debug.WriteLine($"GetFriendlySenderName error: {ex.Message}")
            Return "未知发件人"
        End Try
    End Function

    ' 获取简洁的会议组织者姓名
    Private Function GetFriendlyOrganizerName(meetingItem As Outlook.MeetingItem) As String
        Try
            ' 首先尝试获取Sender的DisplayName
            If meetingItem.Sender IsNot Nothing AndAlso Not String.IsNullOrEmpty(meetingItem.Sender.Name) Then
                Return meetingItem.Sender.Name
            End If

            ' 如果Sender为空，尝试从SenderName中提取
            If Not String.IsNullOrEmpty(meetingItem.SenderName) Then
                Dim senderName As String = meetingItem.SenderName

                ' 如果是Exchange DN格式（以/O=开头），尝试提取CN部分
                If senderName.StartsWith("/O=") Then
                    Dim cnIndex As Integer = senderName.LastIndexOf("/CN=")
                    If cnIndex >= 0 Then
                        Dim cnPart As String = senderName.Substring(cnIndex + 4)
                        ' 移除可能的后缀
                        Dim dashIndex As Integer = cnPart.IndexOf("-")
                        If dashIndex >= 0 Then
                            cnPart = cnPart.Substring(0, dashIndex)
                        End If
                        Return cnPart
                    End If
                End If

                ' 如果不是Exchange DN格式，直接返回
                Return senderName
            End If

            Return "未知组织者"
        Catch ex As System.Exception
            Debug.WriteLine($"GetFriendlyOrganizerName error: {ex.Message}")
            Return "未知组织者"
        End Try
    End Function

    ''' <summary>
    ''' 递归查找控件中的TreeView
    ''' </summary>
    Private Function FindTreeViewInControl(control As Control) As TreeView
        ' 如果当前控件就是TreeView，直接返回
        If TypeOf control Is TreeView Then
            Return DirectCast(control, TreeView)
        End If

        ' 递归查找子控件
        For Each childControl As Control In control.Controls
            Dim treeView As TreeView = FindTreeViewInControl(childControl)
            If treeView IsNot Nothing Then
                Return treeView
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' 更新会话信息按钮标题，显示当前邮件数量
    ''' </summary>
    Private Sub UpdateConversationInfoButtonTitle()
        Try
            If conversationInfoButton IsNot Nothing Then
                If Me.InvokeRequired Then
                    Me.BeginInvoke(Sub() UpdateConversationInfoButtonTitle())
                Else
                    conversationInfoButton.Text = $"会话信息({lvMails.Items.Count})"
                    conversationInfoButton.TextAlign = ContentAlignment.MiddleLeft
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"更新会话信息按钮标题时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 公共方法：设置当前邮件的高亮状态
    ''' </summary>
    ''' <param name="entryID">要高亮的邮件EntryID</param>
    Public Sub SetCurrentHighlight(entryID As String)
        Try
            ' === 详细调试信息输出 ===
            Debug.WriteLine($"=== SetCurrentHighlight 调试开始 ===")
            Debug.WriteLine($"时间: {DateTime.Now:HH:mm:ss.fff}")
            Debug.WriteLine($"接收到EntryID: {entryID}")
            Debug.WriteLine($"EntryID长度: {If(entryID?.Length, 0)}")
            
            ' 确保在UI线程上执行
            If Me.InvokeRequired Then
                Debug.WriteLine("需要跨线程调用，使用Invoke")
                Me.Invoke(Sub() SetCurrentHighlight(entryID))
                Return
            End If
            
            ' 检查lvMails控件是否可用
            If lvMails Is Nothing Then
                Debug.WriteLine("错误: lvMails控件为null")
                Return
            End If
            
            ' 强制创建句柄并确保控件可见
            If Not lvMails.IsHandleCreated Then
                Debug.WriteLine("警告: lvMails控件句柄尚未创建")
                ' 尝试强制创建句柄
                Try
                    Dim handle = lvMails.Handle
                    Debug.WriteLine($"强制创建句柄结果: {handle}")
                    
                    ' 确保控件可见
                    If Not lvMails.Visible Then
                        lvMails.Visible = True
                        Debug.WriteLine("已设置lvMails为可见")
                    End If
                    
                    ' 强制布局更新
                    lvMails.PerformLayout()
                    Me.PerformLayout()
                Catch ex As Exception
                    Debug.WriteLine($"创建句柄失败: {ex.Message}")
                    ' 如果创建句柄失败，延迟重试
                    System.Threading.Tasks.Task.Delay(500).ContinueWith(Sub() 
                        If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
                            Me.BeginInvoke(Sub() SetCurrentHighlight(entryID))
                        End If
                    End Sub)
                    Return
                End Try
            End If
            
            ' 处理EntryID格式转换
            Dim normalizedEntryID As String = ""
            Dim shortEntryID As String = ""
            
            If Not String.IsNullOrEmpty(entryID) Then
                Try
                    normalizedEntryID = ConvertStringToHexFormat(entryID)
                    shortEntryID = OutlookMyList.Utils.OutlookUtils.GetShortEntryID(normalizedEntryID)
                    
                    Debug.WriteLine($"原始EntryID格式: {If(entryID.StartsWith("0000"), "长格式", If(entryID.StartsWith("EF"), "短格式", "未知格式"))}")
                    Debug.WriteLine($"规范化EntryID: {normalizedEntryID}")
                    Debug.WriteLine($"短格式EntryID: {shortEntryID}")
                Catch ex As Exception
                    Debug.WriteLine($"EntryID格式转换失败: {ex.Message}")
                    ' 使用原始EntryID继续
                    normalizedEntryID = entryID
                End Try
                
                ' 获取当前高亮状态
                Dim currentHighlight As String = GetCurrentHighlightEntryID()
                Debug.WriteLine($"当前高亮EntryID: {If(String.IsNullOrEmpty(currentHighlight), "无", currentHighlight)}")
                
                ' 检查lvMails中的项目总数
                Debug.WriteLine($"lvMails项目总数: {lvMails.Items.Count}")
                
                ' 直接在这里查找匹配项并设置高亮
                Dim foundItem As ListViewItem = Nothing
                
                ' 首先尝试使用短格式EntryID查找（最可靠的方式）
                If Not String.IsNullOrEmpty(shortEntryID) Then
                    For Each item As ListViewItem In lvMails.Items
                        If item.Tag IsNot Nothing Then
                            Try
                                Dim itemEntryID As String = item.Tag.ToString()
                                Dim itemShortEntryID As String = OutlookMyList.Utils.OutlookUtils.GetShortEntryID(itemEntryID)
                                
                                If String.Equals(itemShortEntryID, shortEntryID, StringComparison.OrdinalIgnoreCase) Then
                                    foundItem = item
                                    Debug.WriteLine($"通过短格式EntryID找到匹配项: {itemEntryID}")
                                    Exit For
                                End If
                            Catch ex As Exception
                                Debug.WriteLine($"短格式EntryID比较失败: {ex.Message}")
                                ' 继续下一项
                            End Try
                        End If
                    Next
                End If
                
                ' 如果通过短格式EntryID未找到，尝试其他方法
                If foundItem Is Nothing Then
                    For Each item As ListViewItem In lvMails.Items
                        If item.Tag IsNot Nothing Then
                            Dim itemEntryID As String = item.Tag.ToString()
                            Dim isMatched As Boolean = False
                            
                            ' 1. 尝试使用Outlook Session的CompareEntryIDs方法
                            Try
                                isMatched = Globals.ThisAddIn.Application.Session.CompareEntryIDs(itemEntryID, normalizedEntryID)
                                If isMatched Then
                                    Debug.WriteLine($"通过CompareEntryIDs找到匹配项: {itemEntryID}")
                                    foundItem = item
                                    Exit For
                                End If
                            Catch ex As Exception
                                Debug.WriteLine($"CompareEntryIDs失败: {ex.Message}")
                            End Try
                            
                            ' 2. 尝试直接比较规范化的EntryID
                            If Not isMatched Then
                                isMatched = String.Equals(itemEntryID, normalizedEntryID, StringComparison.OrdinalIgnoreCase)
                                If isMatched Then
                                    Debug.WriteLine($"通过规范化EntryID找到匹配项: {itemEntryID}")
                                    foundItem = item
                                    Exit For
                                End If
                            End If
                            
                            ' 3. 尝试比较原始EntryID的最后部分
                            If Not isMatched AndAlso itemEntryID.Length >= 20 AndAlso entryID.Length >= 20 Then
                                Dim itemLastPart As String = itemEntryID.Substring(itemEntryID.Length - 20)
                                Dim entryLastPart As String = entryID.Substring(entryID.Length - 20)
                                isMatched = String.Equals(itemLastPart, entryLastPart, StringComparison.OrdinalIgnoreCase)
                                If isMatched Then
                                    Debug.WriteLine($"通过EntryID尾部比较找到匹配项: {itemEntryID}")
                                    foundItem = item
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                
                ' 清除所有高亮
                For Each item As ListViewItem In lvMails.Items
                    If item.BackColor <> currentBackColor Then
                        Debug.WriteLine($"清除项目高亮: {If(item.Tag IsNot Nothing, item.Tag.ToString(), "无Tag")}")
                        item.BackColor = currentBackColor
                        item.ForeColor = currentForeColor
                        item.Font = New Font(item.Font, FontStyle.Regular)
                    End If
                Next
                
                ' 强制刷新清除高亮后的状态
                lvMails.Update()
                lvMails.Refresh()
                Application.DoEvents()
                
                ' 设置新的高亮项
                If foundItem IsNot Nothing Then
                    Debug.WriteLine($"直接设置高亮: {If(foundItem.Tag IsNot Nothing, foundItem.Tag.ToString(), "无Tag")}")
                    
                    ' 使用更明显的高亮效果
                    foundItem.BackColor = Color.FromArgb(65, 105, 225) ' 皇家蓝
                    foundItem.ForeColor = Color.White
                    foundItem.Font = New Font(foundItem.Font, FontStyle.Bold)
                    foundItem.Selected = True
                    foundItem.UseItemStyleForSubItems = True
                    
                    ' 确保项目可见
                    foundItem.EnsureVisible()
                    
                    ' 强制多次刷新以确保高亮显示
                    lvMails.Update()
                    lvMails.Refresh()
                    Application.DoEvents()
                    
                    ' 再次确保项目可见
                    foundItem.EnsureVisible()
                    
                    ' 再次强制刷新
                    lvMails.Update()
                    lvMails.Refresh()
                    Application.DoEvents()
                    
                    Debug.WriteLine("已设置高亮项")
                    
                    ' 延迟再次刷新，确保高亮效果显示
                     System.Threading.Tasks.Task.Delay(100).ContinueWith(Sub()
                         If Me.IsHandleCreated AndAlso Not Me.IsDisposed Then
                             Me.BeginInvoke(Sub()
                                 ' ListViewItem没有IsDisposed属性，改为检查foundItem是否为Nothing
                                 If foundItem IsNot Nothing Then
                                     foundItem.EnsureVisible()
                                     lvMails.Update()
                                     lvMails.Refresh()
                                     Application.DoEvents()
                                     Debug.WriteLine("延迟刷新高亮项完成")
                                 End If
                             End Sub)
                         End If
                     End Sub)
                Else
                    Debug.WriteLine("未找到匹配项，无法设置高亮")
                End If
            Else
                Debug.WriteLine("接收到空EntryID")
                
                ' 清除所有高亮
                For Each item As ListViewItem In lvMails.Items
                    item.BackColor = currentBackColor
                    item.ForeColor = currentForeColor
                    item.Font = New Font(item.Font, FontStyle.Regular)
                Next
                
                ' 强制刷新
                lvMails.Update()
                lvMails.Refresh()
            End If
            
            Debug.WriteLine($"=== SetCurrentHighlight 调试结束 ===")
        Catch ex As System.Exception
            Debug.WriteLine($"设置高亮时发生异常: {ex.Message}")
            Debug.WriteLine($"异常堆栈: {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' 公共方法：获取当前高亮的邮件EntryID
    ''' </summary>
    ''' <returns>当前高亮的邮件EntryID，如果没有高亮项则返回空字符串</returns>
    Public Function GetCurrentHighlightEntryID() As String
        Dim entryID As String = ""
        
        ' 获取当前高亮的邮件EntryID
        If Me.InvokeRequired Then
            Me.Invoke(Sub()
                          ' 查找高亮的项
                          For Each item As ListViewItem In lvMails.Items
                              If item.BackColor = Color.LightBlue OrElse item.BackColor = Color.LightSkyBlue Then
                                  entryID = item.Tag?.ToString()
                                  Exit For
                              End If
                          Next
                      End Sub)
        Else
            ' 查找高亮的项
            For Each item As ListViewItem In lvMails.Items
                If item.BackColor = Color.LightBlue OrElse item.BackColor = Color.LightSkyBlue Then
                    entryID = item.Tag?.ToString()
                    Exit For
                End If
            Next
        End If
        
        Return entryID
    End Function
End Class

