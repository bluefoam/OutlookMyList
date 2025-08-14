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


<ComVisible(True)>
Public Class MailThreadPane
    Inherits UserControl

    ' 添加类级别的字体缓存
    Private ReadOnly iconFont As Font
    Private ReadOnly defaultFont As Font
    Private ReadOnly highlightFont As Font
    Private ReadOnly normalFont As Font
    Private ReadOnly highlightColor As Color = Color.FromArgb(255, 255, 200)
    
    ' 主题颜色
    Private currentBackColor As Color = SystemColors.Window
    Private currentForeColor As Color = SystemColors.WindowText
    
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
    
            ' 应用到WebBrowser控件
            If wbContent IsNot Nothing AndAlso wbContent.Document IsNot Nothing Then
                ' 为WebBrowser设置背景色
                Dim bgColorHex As String = "#" & backgroundColor.R.ToString("X2") & backgroundColor.G.ToString("X2") & backgroundColor.B.ToString("X2")
                Dim fgColorHex As String = "#" & foregroundColor.R.ToString("X2") & foregroundColor.G.ToString("X2") & foregroundColor.B.ToString("X2")
                
                Try
                    ' 通过JavaScript设置背景色、文本颜色和CSS变量
                    Dim script As String = "" & _
                    "document.body.style.backgroundColor = '" & bgColorHex & "';" & _
                    "document.body.style.color = '" & fgColorHex & "';" & _
                    "document.documentElement.style.setProperty('--theme-color', '#0078d7');"
                    
                    wbContent.Document.InvokeScript("eval", New Object() {script})
                Catch ex As System.Exception
                    Debug.WriteLine("设置WebBrowser颜色出错: " & ex.Message)
                End Try
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
    Private wbContent As WebBrowser
    Private splitter1, splitter2 As SplitContainer
    Private tabControl As TabControl
    Private btnPanel As Panel
    Private currentConversationId As String = String.Empty
    Private currentMailEntryID As String = String.Empty
    Private currentSortColumn As Integer = 0
    Private currentSortOrder As SortOrder = SortOrder.Ascending
    Private currentHighlightEntryID As String

    Private mailItems As New List(Of (Index As Integer, EntryID As String))  ' 移到这里

    ' 在类级别添加一个字典来存储链接和EntryID的映射
    Private mailLinkMap As New Dictionary(Of String, String)

    ' 删除原来的 mailIndexMap

    Private Sub SetupControls()
        InitializeSplitContainers()
        SetupMailList()
        SetupMailContent()

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

        ' 然后添加第一个分隔控件到窗体
        Me.Controls.Add(splitter1)

        ' 添加尺寸改变事件处理
        AddHandler Me.SizeChanged, AddressOf Control_Resize
        AddHandler splitter1.Panel2.SizeChanged, AddressOf Panel2_SizeChanged
    End Sub

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

    Private Sub ExecuteJavaScript(script As String)
        Try
            If infoWebBrowser Is Nothing Then
                Debug.WriteLine("infoWebBrowser 是 null")
                Return
            End If

            If infoWebBrowser.Document Is Nothing Then
                Debug.WriteLine("Document 是 null")
                Return
            End If

            infoWebBrowser.Document.InvokeScript("eval", New Object() {script})
            Debug.WriteLine("JavaScript 脚本执行成功")
        Catch ex As System.Exception
            Debug.WriteLine($"执行 JavaScript 出错: {ex.Message}")
            Throw
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

    Private Function GetItemImageText(item As Object) As String
        Try
            Dim icons As New List(Of String)



            ' 检查项目类型
            If TypeOf item Is Outlook.MailItem Then
                icons.Add("✉️") '📧
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                icons.Add("📅")
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                icons.Add("👥")
            Else
                icons.Add("❓")
            End If

            ' 根据任务状态添加不同的图标
            Select Case CheckItemHasTask(item)
                Case TaskStatus.InProgress
                    icons.Add("🚩")
                Case TaskStatus.Completed
                    icons.Add("✔")   '✅
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
            .Sorting = SortOrder.Descending,
            .AllowColumnReorder = True,
            .OwnerDraw = True,  ' 启用自定义绘制
            .BackColor = currentBackColor,
            .ForeColor = currentForeColor
        }

        lvMails.Columns.Add("----", 40)  ' 增加宽度以适应更大的图标
        lvMails.Columns.Add("日期", 100)
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

        splitter1.Panel1.Controls.Add(lvMails)

        ' 添加绘制事件处理
        AddHandler lvMails.DrawColumnHeader, AddressOf ListView_DrawColumnHeader
        AddHandler lvMails.DrawSubItem, AddressOf ListView_DrawSubItem
    End Sub



    Private Sub ListView_DrawColumnHeader(sender As Object, e As DrawListViewColumnHeaderEventArgs)
        e.DrawDefault = True
    End Sub

    Private Sub ListView_DrawSubItem(sender As Object, e As DrawListViewSubItemEventArgs)
        ' 使用当前项的背景色
        Dim backBrush As Brush = New SolidBrush(e.Item.BackColor)
        e.Graphics.FillRectangle(backBrush, e.Bounds)

        ' 第一列使用 emoji 字体，其他列使用默认字体
        If e.ColumnIndex = 0 Then
            If e.SubItem.Text.Contains("🚩") Then
                ' 使用特殊颜色和字体
                Dim specialFont As New Font(iconFont, FontStyle.Bold)
                Dim specialBrush As Brush = Brushes.Red
                e.Graphics.DrawString(e.SubItem.Text, specialFont, specialBrush, e.Bounds)
            Else
                e.Graphics.DrawString(e.SubItem.Text, iconFont, Brushes.Black, e.Bounds)
            End If
        Else
            ' 根据是否高亮使用不同字体
            Dim font As Font = If(e.Item.BackColor = highlightColor, highlightFont, normalFont)
            e.Graphics.DrawString(e.SubItem.Text, font, Brushes.Black, e.Bounds)
        End If
        backBrush.Dispose()
    End Sub

    Private Sub SetupMailContent()
        wbContent = New WebBrowser With {
            .Dock = DockStyle.Fill,
            .ScrollBarsEnabled = True,
            .ScriptErrorsSuppressed = True,  ' 忽略脚本错误
            .AllowNavigation = True,
            .IsWebBrowserContextMenuEnabled = True,
            .WebBrowserShortcutsEnabled = True
        }

        Try
            wbContent.ObjectForScripting = Me
        Catch ex As System.Exception
            Debug.WriteLine($"设置 ObjectForScripting 失败: {ex.Message}")
        End Try

        splitter2.Panel1.Controls.Add(wbContent)
        ' 添加导航事件处理
        AddHandler wbContent.Navigating, AddressOf WebBrowser_Navigating
    End Sub

    Private Sub SetupTabPages()
        tabControl = New TabControl With {
            .Dock = DockStyle.Fill
        }
        splitter2.Panel2.Controls.Add(tabControl)

        ' 只初始化第一个标签页
        SetupNotesTab()

        ' 延迟加载其他标签页
        Task.Run(Sub()
                     Me.Invoke(Sub()
                                   SetupTasksTab()
                                   SetupActionsTab()
                                   tabControl.SelectedIndex = 0
                               End Sub)
                 End Sub)
    End Sub

    ' Add this new method
    <ComVisible(True)>
    Private Sub WebBrowser_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs)
        Try
            If e.Url.ToString() <> "about:blank" Then
                e.Cancel = True  ' Cancel default navigation
                Process.Start(New ProcessStartInfo With {
                    .FileName = e.Url.ToString(),
                    .UseShellExecute = True
                })
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"打开链接出错: {ex.Message}")
            MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。")
        End Try
    End Sub

    Private WithEvents infoWebBrowser As WebBrowser  ' 添加到类级别变量

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

        ' 替换原来的 infoWebBrowser
        infoWebBrowser = Nothing

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
            "Weekly"
        }

            ' 检查当前文件夹是否是邮件文件夹且在核心文件夹列表中
            If folder.DefaultItemType = Outlook.OlItemType.olMailItem AndAlso
           coreFolders.Contains(folder.Name) Then
                folderList.Add(folder)
            End If

            ' 只在核心文件夹中递归搜索
            For Each subFolder As Outlook.Folder In folder.Folders
                If coreFolders.Contains(subFolder.Name) Then
                    GetAllMailFolders(subFolder, folderList)
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"处理文件夹 {folder.Name} 时出错: {ex.Message}")
        End Try
    End Sub
    ' 添加一个新的辅助方法用于递归获取所有邮件文件夹
    Private Sub GetAllMailFoldersAll(folder As Outlook.Folder, folderList As List(Of Outlook.Folder))
        Try
            ' 添加当前文件夹（如果是邮件文件夹）
            If folder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                folderList.Add(folder)
            End If

            ' 递归处理子文件夹
            For Each subFolder As Outlook.Folder In folder.Folders
                GetAllMailFolders(subFolder, folderList)
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"处理文件夹 {folder.Name} 时出错: {ex.Message}")
        End Try
    End Sub

    Private Async Function GetContactInfoAsync() As Task(Of String)
        Try
            Dim info As New StringBuilder()
            Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
            If currentItem Is Nothing Then Return "未选择邮件项"

            Dim senderEmail As String = String.Empty
            Dim senderName As String = String.Empty

            ' 获取发件人信息
            If TypeOf currentItem Is Outlook.MailItem Then
                Dim mail = DirectCast(currentItem, Outlook.MailItem)
                senderEmail = mail.SenderEmailAddress
                senderName = mail.SenderName
            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                Dim meeting = DirectCast(currentItem, Outlook.MeetingItem)
                senderEmail = meeting.SenderEmailAddress
                senderName = meeting.SenderName
            End If

            If String.IsNullOrEmpty(senderEmail) Then Return "无法获取发件人信息"

            info.AppendLine($"发件人: {senderName}")
            info.AppendLine($"邮箱: {senderEmail}")
            info.AppendLine("----------------------------------------")

            ' 搜索联系人信息
            Dim contacts = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
            Dim filter = $"[Email1Address] = '{senderEmail}' OR [Email2Address] = '{senderEmail}' OR [Email3Address] = '{senderEmail}'"
            Dim matchingContacts = contacts.Items.Restrict(filter)

            If matchingContacts.Count > 0 Then
                Dim contact = DirectCast(matchingContacts(1), Outlook.ContactItem)
                info.AppendLine("联系人信息:")
                If Not String.IsNullOrEmpty(contact.BusinessTelephoneNumber) Then
                    info.AppendLine($"工作电话: {contact.BusinessTelephoneNumber}")
                End If
                If Not String.IsNullOrEmpty(contact.MobileTelephoneNumber) Then
                    info.AppendLine($"手机: {contact.MobileTelephoneNumber}")
                End If
                If Not String.IsNullOrEmpty(contact.Department) Then
                    info.AppendLine($"部门: {contact.Department}")
                End If
                If Not String.IsNullOrEmpty(contact.CompanyName) Then
                    info.AppendLine($"公司: {contact.CompanyName}")
                End If
                info.AppendLine("----------------------------------------")
            End If

            ' 统计会议信息
            Dim calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
            Dim startDate = DateTime.Now.AddMonths(-3)
            Dim endDate = DateTime.Now.AddMonths(1)
            Dim meetingFilter = $"[Start] >= '{startDate:MM/dd/yyyy}' AND [End] <= '{endDate:MM/dd/yyyy}'"
            Dim meetings = calendar.Items.Restrict(meetingFilter)

            Dim meetingStats As New Dictionary(Of String, Integer)
            Dim totalMeetings As Integer = 0
            Dim upcomingMeetings As New List(Of (MeetingDate As DateTime, Title As String))

            For i = meetings.Count To 1 Step -1
                Dim meeting = DirectCast(meetings(i), Outlook.AppointmentItem)
                If meeting.RequiredAttendees IsNot Nothing AndAlso
               (meeting.RequiredAttendees.Contains(senderEmail) OrElse
                meeting.OptionalAttendees?.Contains(senderEmail)) Then

                    totalMeetings += 1

                    ' 提取项目名称
                    Dim projectName = "其他"
                    Dim match = System.Text.RegularExpressions.Regex.Match(meeting.Subject, "\[(.*?)\]")
                    If match.Success Then
                        projectName = match.Groups(1).Value
                    End If

                    If meetingStats.ContainsKey(projectName) Then
                        meetingStats(projectName) += 1
                    Else
                        meetingStats.Add(projectName, 1)
                    End If

                    If meeting.Start > DateTime.Now Then
                        upcomingMeetings.Add((meeting.Start, meeting.Subject))
                    End If
                End If
            Next

            info.AppendLine($"会议统计 (近3个月):")
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

            ' 统计邮件往来
            ' 统计邮件往来
            Dim mailCount As Integer = 0
            Dim recentMails As New List(Of Outlook.MailItem)

            ' 获取所有邮件文件夹
            Dim folders As New List(Of Outlook.Folder)
            Dim store As Outlook.Store = Globals.ThisAddIn.Application.Session.DefaultStore
            GetAllMailFolders(store.GetRootFolder(), folders)

            ' 遍历所有文件夹搜索邮件
            For Each folder In folders
                Try
                    Dim mailFilter = $"[SenderEmailAddress] = '{senderEmail}'"
                    Dim folderMails = folder.Items.Restrict(mailFilter)
                    mailCount += folderMails.Count

                    ' 收集最近的邮件
                    For i = folderMails.Count To 1 Step -1
                        If recentMails.Count >= 30 Then Exit For
                        Dim mail = TryCast(folderMails(i), Outlook.MailItem)
                        If mail IsNot Nothing Then
                            recentMails.Add(mail)
                        End If
                    Next
                Catch ex As SystemException
                    Debug.WriteLine($"搜索文件夹 {folder.Name} 时出错: {ex.Message}")
                    Continue For
                End Try
            Next



            info.AppendLine($"邮件往来统计:")
            info.AppendLine($"总邮件数: {mailCount}")
            info.AppendLine("最近邮件:")

            ' 清除之前的映射
            mailLinkMap.Clear()

            ' 按时间排序并显示最近邮件，添加序号
            Dim sortedMails = recentMails.OrderByDescending(Function(m) m.ReceivedTime).Take(30).ToList()
            For i As Integer = 0 To sortedMails.Count - 1
                Dim mail = sortedMails(i)
                ' 创建唯一的链接ID
                Dim linkId = $"m_{i + 1}"
                ' 存储映射关系
                mailLinkMap(linkId) = mail.EntryID
                ' 添加序号，使用简短链接ID
                info.AppendLine($"- [{i + 1}] {mail.ReceivedTime:yyyy-MM-dd HH:mm} http://{linkId} {mail.Subject.Replace("[EXT]", "")}")
            Next

            Return info.ToString()  ' 添加返回语句
        Catch ex As System.Exception
            Debug.WriteLine($"获取联系人信息时出错: {ex.Message}")
            Return $"获取联系人信息时出错: {ex.Message}"
        End Try
    End Function

    ' 修改导航事件处理程序
    <ComVisible(True)>
    Private Sub infoWebBrowser_Navigating(sender As Object, e As WebBrowserNavigatingEventArgs) Handles infoWebBrowser.Navigating
        Try
            ' 检查是否是邮件链接
            If e.Url.ToString() <> "about:blank" Then
                e.Cancel = True  ' 取消 WebBrowser 的默认导航
                Debug.WriteLine($"正在尝试打开链接: {e.Url}")

                ' 检查是否是邮件链接
                If e.Url.ToString().StartsWith("outlook-mail:") Then
                    Dim mailEntryID = e.Url.ToString().Replace("outlook-mail:", "")
                    OpenOutlookMail(mailEntryID)
                Else
                    ' 普通链接，使用默认浏览器打开
                    Process.Start(New ProcessStartInfo With {
                        .FileName = e.Url.ToString(),
                        .UseShellExecute = True
                    })
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"打开链接出错: {ex.Message}")
            MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。")
        End Try
    End Sub

    ' 添加打开邮件的方法
    Private Sub OpenOutlookMail(entryID As String)
        Try
            ' 使用 Application.CreateItem 方法而不是直接获取项目
            ' 这可以避免一些 COM 互操作问题
            Dim mailItem = Globals.ThisAddIn.Application.Session.GetItemFromID(entryID)
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
            .Height = 40
        }

        ' 使用 RichTextBox 替代 TextBox
        Dim outputTextBox As New RichTextBox With {
            .Multiline = True,
            .ScrollBars = RichTextBoxScrollBars.Vertical,
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .DetectUrls = True  ' 启用URL检测
        }

        ' 添加链接点击事件
        AddHandler outputTextBox.LinkClicked, AddressOf OutputTextBox_LinkClicked

        ' 只创建按钮，不预先创建文本框
        Dim x As Integer = 10
        For i As Integer = 1 To 3
            Dim btn As New Button With {
                .Text = If(i = 1, "联系人信息", $"按钮 {i}"),
                .Location = New Point(x, 5),
                .Size = New Size(120, 30)
            }

            ' 特别处理第一个按钮 - 延迟初始化
            If i = 1 Then
                AddHandler btn.Click, Sub(s, e)
                                          GetContactInfoHandler(outputTextBox)
                                      End Sub
            Else
                AddHandler btn.Click, Sub(s, e)
                                          outputTextBox.Text = "正在获取会话信息..."
                                          Dim conversationTitle As String = "当前会话"
                                          outputTextBox.Text = $"当前会话ID: {currentConversationId}" & vbCrLf &
                                                                $"会话邮件数量: {lvMails.Items.Count}" & vbCrLf &
                                                                $"当前邮件ID: {currentMailEntryID}"
                                      End Sub
            End If

            btnPanel.Controls.Add(btn)
            x += 125
        Next

        ' 先添加文本框到主面板
        btnPanel.Controls.Add(outputTextBox)
        ' 再添加按钮面板到主面板
        btnPanel.Controls.Add(buttonPanel)

        tabPage3.Controls.Add(btnPanel)
        tabControl.TabPages.Add(tabPage3)
    End Sub

    ' 然后修改链接点击事件处理程序
    Private Sub OutputTextBox_LinkClicked(sender As Object, e As LinkClickedEventArgs)
        Try
            ' 检查是否是邮件链接
            If e.LinkText.StartsWith("http://m_") Then
                Dim linkId = e.LinkText.Replace("http://", "")
                If mailLinkMap.ContainsKey(linkId) Then
                    ' 使用 Control.Invoke 而不是 BeginInvoke
                    If Me.InvokeRequired Then
                        Me.Invoke(Sub() SafeOpenOutlookMail(mailLinkMap(linkId)))
                    Else
                        SafeOpenOutlookMail(mailLinkMap(linkId))
                    End If
                Else
                    MessageBox.Show("无法找到对应的邮件")
                End If
            Else
                ' 普通链接，使用默认浏览器打开
                Process.Start(New ProcessStartInfo With {
                    .FileName = e.LinkText,
                    .UseShellExecute = True
                })
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"处理链接点击时出错: {ex.Message}")
        End Try
    End Sub

    Private Sub SafeOpenOutlookMail(entryID As String)
        Try
            Debug.WriteLine($"尝试打开邮件，EntryID: {If(entryID?.Length > 10, entryID.Substring(0, 10) & "...", "null")}")

            ' 检查EntryID是否有效
            If String.IsNullOrEmpty(entryID) Then
                Debug.WriteLine("EntryID为空")
                Return ' 不显示错误消息
            End If

            ' 直接使用最简单的方法打开邮件
            Debug.WriteLine("直接使用简单方法打开邮件")

            ' 获取邮件项并直接显示
            Dim mailItem = Nothing
            Try
                mailItem = Globals.ThisAddIn.Application.Session.GetItemFromID(entryID)
                If mailItem IsNot Nothing Then
                    Debug.WriteLine("成功获取邮件项，尝试显示")

                    ' 直接调用Display方法
                    If TypeOf mailItem Is Outlook.MailItem Then
                        DirectCast(mailItem, Outlook.MailItem).Display(False)
                        Debug.WriteLine("邮件显示成功")
                    ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                        DirectCast(mailItem, Outlook.AppointmentItem).Display(False)
                        Debug.WriteLine("会议项显示成功")
                    ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                        DirectCast(mailItem, Outlook.MeetingItem).Display(False)
                        Debug.WriteLine("会议邮件显示成功")
                    ElseIf TypeOf mailItem Is Outlook.TaskItem Then
                        DirectCast(mailItem, Outlook.TaskItem).Display(False)
                        Debug.WriteLine("任务项显示成功")
                    Else
                        ' 对于其他类型，尝试通用方法
                        CallByName(mailItem, "Display", CallType.Method)
                        Debug.WriteLine("项目显示成功")
                    End If
                Else
                    Debug.WriteLine("GetItemFromID返回空")
                End If
            Catch itemEx As System.Exception
                Debug.WriteLine($"获取或显示邮件项时出错: {itemEx.Message}")
                ' 捕获错误但不显示给用户
            Finally
                If mailItem IsNot Nothing Then
                    Try
                        Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                        Debug.WriteLine("已释放邮件COM对象")
                    Catch releaseEx As System.Exception
                        Debug.WriteLine($"释放COM对象时出错: {releaseEx.Message}")
                    End Try
                End If
            End Try
        Catch ex As System.Exception
            Debug.WriteLine($"安全打开邮件时出错: {ex.Message}")
            Debug.WriteLine($"错误堆栈: {ex.StackTrace}")
            ' 不显示错误消息
        End Try
    End Sub

    ' 将异步逻辑移到单独的方法中
    ' 将异步逻辑移到单独的方法中
    Private Async Sub GetContactInfoHandler(outputTextBox As Control)
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
            Dim info = Await Task.Run(Function() GetContactInfoAsync().Result)

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
        End Try
    End Sub

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
            Dim noteList As New List(Of (CreateTime As String, Title As String, Link As String))
            ' 首先检查所有相关邮件的属性
            Try
                ' 获取当前会话的所有邮件

                Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
                Dim conversation As Outlook.Conversation = Nothing

                ' 获取 conversation 对象前先检查类型
                If TypeOf currentItem Is Outlook.MailItem Then
                    conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
                End If


                If conversation IsNot Nothing Then
                    Dim table As Outlook.Table = conversation.GetTable()

                    ' 遍历会话中的所有项目
                    Do Until table.EndOfTable
                        Dim item As Object = Nothing  ' Declare item at the beginning of the loop
                        Try
                            Dim row As Outlook.Row = table.GetNextRow()
                            item = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString())

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
        Catch ex As System.Exception
            Debug.WriteLine($"CheckWolaiRecord 执行出错: {ex.Message}")
            Return String.Empty
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
                            Dim item As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
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

        Catch ex As System.Exception
            Debug.WriteLine($"SaveToWolai 执行出错: {ex.Message}")
            MessageBox.Show($"保存失败: {ex.Message}")
            Return False
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
                                                            Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
                                                            If mailItem IsNot Nothing Then
                                                                ' 根据不同类型获取主题
                                                                If TypeOf mailItem Is Outlook.MailItem Then
                                                                    Return DirectCast(mailItem, Outlook.MailItem).Subject
                                                                ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                                                                    Return DirectCast(mailItem, Outlook.AppointmentItem).Subject
                                                                ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                                                                    Return DirectCast(mailItem, Outlook.MeetingItem).Subject
                                                                ElseIf TypeOf mailItem Is Outlook.TaskItem Then
                                                                    Return DirectCast(mailItem, Outlook.TaskItem).Subject
                                                                End If
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

    End Sub

    ' 添加类级别的防重复调用变量
    Private isUpdatingMailList As Boolean = False
    Private lastUpdateTime As DateTime = DateTime.MinValue
    Private Const UpdateThreshold As Integer = 500 ' 毫秒

    Public Async Sub UpdateMailList(conversationId As String, mailEntryID As String)
        Try

            ' 添加堆栈跟踪日志，查看谁调用了这个方法
            Debug.WriteLine($"UpdateMailList 被调用，调用堆栈: {Environment.StackTrace}")

            If String.IsNullOrEmpty(mailEntryID) Then
                lvMails?.Items.Clear()
                Return
            End If

            ' 记录开始时间，用于性能分析
            Dim startTime = DateTime.Now
            Debug.WriteLine($"开始更新邮件列表: {startTime}")

            ' 检查是否需要重新加载列表
            Dim needReload As Boolean = True
            If lvMails.Items.Count > 0 AndAlso Not String.IsNullOrEmpty(conversationId) AndAlso
           String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                needReload = False
            End If

            ' 单独处理无会话的邮件
            If Not String.IsNullOrEmpty(mailEntryID) AndAlso String.IsNullOrEmpty(conversationId) Then
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailEntryID)
                currentMailEntryID = mailEntryID
                Debug.WriteLine($"处理无会话邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Return
            End If

            If needReload Then
                ' 异步加载会话邮件，完全不阻塞主窗口
                Await LoadConversationMailsAsync(mailEntryID)

                ' 更新当前会话ID并检查笔记
                If Not String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                    currentConversationId = conversationId
                    Await CheckWolaiRecordAsync(currentConversationId)
                End If
            Else
                ' 只更新高亮和内容
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailEntryID)
                UpdateHighlightByEntryID(currentMailEntryID, mailEntryID)
            End If

            currentMailEntryID = mailEntryID
            Debug.WriteLine($"完成更新邮件列表，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateMailList error: {ex.Message}")
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
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailEntryID)
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
                    Await CheckWolaiRecordAsync(currentConversationId)
                End If


            Else
                ' 只更新高亮和内容
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailEntryID)
                UpdateHighlightByEntryID(currentMailEntryID, mailEntryID)
            End If
            currentMailEntryID = mailEntryID
            Debug.WriteLine($"完成更新邮件列表，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateMailList error: {ex.Message}")
        End Try

    End Sub

    Private Function GetIndexByEntryID(entryID As String) As Integer
        Return mailItems.FindIndex(Function(x) String.Equals(x.EntryID, entryID.Trim(), StringComparison.OrdinalIgnoreCase))
    End Function


    ' 新的异步方法，完全在后台线程执行耗时操作
    Private Async Function LoadConversationMailsAsync(currentMailEntryID As String) As Task
        If String.IsNullOrEmpty(currentMailEntryID) Then
            Return
        End If

        Dim startTime = DateTime.Now
        Debug.WriteLine($"开始异步加载会话邮件: {startTime}")

        ' 在UI线程中显示加载状态
        If Me.InvokeRequired Then
            Me.Invoke(Sub()
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
                           LoadConversationMailsBackground(currentMailEntryID, startTime)
                       End Sub)
    End Function

    ' 后台线程执行的邮件加载逻辑
    Private Sub LoadConversationMailsBackground(currentMailEntryID As String, startTime As DateTime)
        Dim currentItem As Object = Nothing
        Dim conversation As Outlook.Conversation = Nothing
        Dim table As Outlook.Table = Nothing
        Dim allItems As New List(Of ListViewItem)()
        Dim tempMailItems As New List(Of (Index As Integer, EntryID As String))()

        Try
            Try
                currentItem = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
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
                        End If
                    End With

                    allItems.Add(lvi)
                    tempMailItems.Add((0, entryId))

                    Debug.WriteLine($"处理单个邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Else
                    ' 使用批量处理方式加载会话邮件
                    table = conversation.GetTable()
                    Try
                        ' 设置需要的列
                        table.Columns.Add("EntryID")
                        table.Columns.Add("SentOn")
                        table.Columns.Add("ReceivedTime")
                        table.Columns.Add("SenderName")
                        table.Columns.Add("Subject")
                        table.Columns.Add("MessageClass")

                        ' 预分配容量，提高性能
                        Dim currentIndex As Integer = 0
                        Dim batchSize As Integer = 0

                        ' 一次性收集所有数据
                        Do Until table.EndOfTable
                            Dim row As Outlook.Row = table.GetNextRow()
                            Dim mailItem As Object = Nothing
                            Try
                                mailItem = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString())
                                If mailItem IsNot Nothing Then
                                    Dim entryId As String = GetPermanentEntryID(mailItem)

                                    ' 创建 ListViewItem
                                    Dim lvi As New ListViewItem(GetItemImageText(mailItem)) With {
                                    .Tag = entryId,
                                    .Name = currentIndex.ToString()
                                }

                                    ' 添加所有列
                                    With lvi.SubItems
                                        If TypeOf mailItem Is Outlook.MeetingItem Then
                                            Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                                            .Add(meeting.CreationTime.ToString("yyyy-MM-dd HH:mm"))
                                            .Add(meeting.SenderName)
                                            .Add(meeting.Subject)
                                        Else
                                            .Add(If(row("ReceivedTime") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row("ReceivedTime").ToString()),
                                            DateTime.Parse(row("ReceivedTime").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                            "Unknown Date"))
                                            .Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                                            .Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))
                                        End If
                                    End With

                                    ' 添加到临时列表
                                    allItems.Add(lvi)
                                    tempMailItems.Add((currentIndex, entryId))
                                    currentIndex += 1
                                    batchSize += 1
                                End If
                            Finally
                                If mailItem IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                                End If
                                If row IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                End If
                            End Try
                        Loop

                        Debug.WriteLine($"收集了 {batchSize} 封邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                    Finally
                        If table IsNot Nothing Then
                            Runtime.InteropServices.Marshal.ReleaseComObject(table)
                        End If
                    End Try
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"处理邮件时出错: {ex.Message}")
                ' 在UI线程中显示错误信息
                Me.Invoke(Sub()
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

        ' 在UI线程中更新界面
        Me.Invoke(Sub()
                      Try
                          lvMails.BeginUpdate()
                          lvMails.Items.Clear()
                          mailItems.Clear()
                          
                          If allItems.Count > 0 Then
                              lvMails.Items.AddRange(allItems.ToArray())
                              mailItems = tempMailItems
                              
                              ' 设置排序
                              lvMails.Sorting = SortOrder.Descending
                              lvMails.ListViewItemSorter = New ListViewItemComparer(1, SortOrder.Descending)
                              lvMails.Sort()
                              
                              ' 设置高亮并确保可见
                              UpdateHighlightByEntryID(String.Empty, currentMailEntryID)
                          End If
                          
                          Debug.WriteLine($"完成异步加载会话邮件，总耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                      Finally
                          lvMails.EndUpdate()
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
                currentItem = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
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
                        End If
                    End With

                    lvMails.Items.Add(lvi)
                    mailItems.Add((0, entryId))

                    Debug.WriteLine($"处理单个邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")
                Else
                    ' 使用批量处理方式加载会话邮件
                    table = conversation.GetTable()
                    Try
                        ' 设置需要的列
                        table.Columns.Add("EntryID")
                        table.Columns.Add("SentOn")
                        table.Columns.Add("ReceivedTime")
                        table.Columns.Add("SenderName")
                        table.Columns.Add("Subject")
                        table.Columns.Add("MessageClass")

                        ' 预分配容量，提高性能
                        Dim allItems As New List(Of ListViewItem)(100)
                        Dim tempMailItems As New List(Of (Index As Integer, EntryID As String))(100)
                        Dim currentIndex As Integer = 0
                        Dim batchSize As Integer = 0

                        ' 一次性收集所有数据
                        Do Until table.EndOfTable
                            Dim row As Outlook.Row = table.GetNextRow()
                            Dim mailItem As Object = Nothing
                            Try
                                mailItem = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString())
                                If mailItem IsNot Nothing Then
                                    Dim entryId As String = GetPermanentEntryID(mailItem)

                                    ' 创建 ListViewItem
                                    Dim lvi As New ListViewItem(GetItemImageText(mailItem)) With {
                                    .Tag = entryId,
                                    .Name = currentIndex.ToString()
                                }

                                    ' 添加所有列
                                    With lvi.SubItems
                                        If TypeOf mailItem Is Outlook.MeetingItem Then
                                            Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                                            .Add(meeting.CreationTime.ToString("yyyy-MM-dd HH:mm"))
                                            .Add(meeting.SenderName)
                                            .Add(meeting.Subject)
                                        Else
                                            .Add(If(row("ReceivedTime") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row("ReceivedTime").ToString()),
                                            DateTime.Parse(row("ReceivedTime").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                            "Unknown Date"))
                                            .Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                                            .Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))
                                        End If
                                    End With

                                    ' 添加到临时列表
                                    allItems.Add(lvi)
                                    tempMailItems.Add((currentIndex, entryId))
                                    currentIndex += 1
                                    batchSize += 1
                                End If
                            Finally
                                If mailItem IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                                End If
                                If row IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                End If
                            End Try
                        Loop

                        Debug.WriteLine($"收集了 {batchSize} 封邮件，耗时: {(DateTime.Now - startTime).TotalMilliseconds}ms")

                        ' 一次性添加所有项目
                        lvMails.Items.Clear()
                        mailItems.Clear()
                        lvMails.Items.AddRange(allItems.ToArray())
                        mailItems = tempMailItems

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
                currentItem = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
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
                        table.Columns.Add("EntryID")
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
                            Dim mailItem As Object = Nothing
                            Try
                                mailItem = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString())
                                If mailItem IsNot Nothing Then
                                    Dim entryId As String = GetPermanentEntryID(mailItem)

                                    ' 创建 ListViewItem
                                    Dim lvi As New ListViewItem(GetItemImageText(mailItem)) With {
                                    .Tag = entryId,
                                    .Name = currentIndex.ToString()
                                }

                                    ' 添加所有列
                                    With lvi.SubItems
                                        If TypeOf mailItem Is Outlook.MeetingItem Then
                                            Dim meeting As Outlook.MeetingItem = DirectCast(mailItem, Outlook.MeetingItem)
                                            .Add(meeting.CreationTime.ToString("yyyy-MM-dd HH:mm"))
                                            .Add(meeting.SenderName)
                                            .Add(meeting.Subject)
                                        Else
                                            .Add(If(row("ReceivedTime") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row("ReceivedTime").ToString()),
                                            DateTime.Parse(row("ReceivedTime").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                            "Unknown Date"))
                                            .Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                                            .Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))
                                        End If
                                    End With

                                    ' 添加到临时列表
                                    allItems.Add(lvi)
                                    tempMailItems.Add((currentIndex, entryId))
                                    currentIndex += 1
                                End If
                            Finally
                                If mailItem IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                                End If
                                If row IsNot Nothing Then
                                    Runtime.InteropServices.Marshal.ReleaseComObject(row)
                                End If
                            End Try
                        Loop

                        ' 一次性添加所有项目
                        lvMails.Items.Clear()
                        mailItems.Clear()
                        lvMails.Items.AddRange(allItems.ToArray())
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

                ' 2. 检查是否被标记为任务
                If mail.IsMarkedAsTask Then
                    ' 使用 FlagStatus 检查任务是否完成
                    If mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete Then
                        Debug.WriteLine($"任务已完成: {mail.Subject}")
                        Return TaskStatus.Completed
                    Else
                        Debug.WriteLine($"任务进行中: {mail.Subject}")
                        Return TaskStatus.InProgress
                    End If
                End If


                ' 1. 检查邮件自身的任务属性
                'If mail.TaskCompletedDate <> DateTime.MinValue OrElse
                '   mail.TaskDueDate <> DateTime.MinValue OrElse
                '   mail.TaskStartDate <> DateTime.MinValue OrElse
                '   mail.IsMarkedAsTask Then
                '    Return True
                'End If

                ' 2. 检查邮件的标志状态
                'If mail.FlagStatus <> Outlook.OlFlagStatus.olNoFlag OrElse
                '   mail.FlagIcon <> Outlook.OlFlagIcon.olNoFlagIcon Then
                '    Return True
                'End If

                ' 3. 检查是否有关联的任务项
                'Try
                'Dim taskFolder As Outlook.Folder = DirectCast(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks), Outlook.Folder)
                'Dim filter As String = $"[MessageClass]='IPM.Task' AND [ConversationID]='{mail.ConversationID}'"
                'Dim tasks As Outlook.Items = taskFolder.Items.Restrict(filter)
                'If tasks.Count > 0 Then
                '    Return True
                'End If
                'Catch ex As System.Exception
                '    Debug.WriteLine($"检查关联任务时出错: {ex.Message}")
                'End Try

                ' 4. 检查自定义属性（如果有使用）
                Try
                    For Each prop As Outlook.UserProperty In mail.UserProperties
                        If prop.Name.StartsWith("Task") Then
                            Return True
                        End If
                    Next
                Catch ex As System.Exception
                    Debug.WriteLine($"检查自定义任务属性时出错: {ex.Message}")
                End Try
            End If

            Return TaskStatus.None
        Catch ex As System.Exception
            Debug.WriteLine($"检查任务标记出错: {ex.Message}")
            Return TaskStatus.None
        End Try
    End Function

    Public Sub New()
        ' 这个调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 之后添加任何初始化代码
        defaultFont = SystemFonts.DefaultFont
        'iconFont = New Font("Segoe UI Emoji", 10)  ' 使用 Segoe UI Emoji 字体以获得更好的 emoji 显示效果
        iconFont = New Font("Segoe UI Emoji", 8, FontStyle.Regular)
        normalFont = New Font(defaultFont, FontStyle.Regular)
        highlightFont = New Font(defaultFont, FontStyle.Bold)  ' 使用 defaultFont 作为基础字体

        ' 最后设置控件
        SetupControls()
    End Sub

    Private Sub UpdateHighlightByEntryID(oldEntryID As String, newEntryID As String)
        Try
            lvMails.BeginUpdate()
            ' 清除所有项的高亮状态
            For Each item As ListViewItem In lvMails.Items
                SetItemHighlight(item, False)
            Next

            ' 设置新的高亮
            If Not String.IsNullOrEmpty(newEntryID) Then
                ' 直接在 ListView 中查找匹配的项
                For Each item As ListViewItem In lvMails.Items
                    If String.Equals(item.Tag.ToString(), newEntryID.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        SetItemHighlight(item, True)
                        item.EnsureVisible()
                        currentHighlightEntryID = newEntryID
                        Exit For
                    End If
                Next
            End If
        Finally
            lvMails.EndUpdate()
        End Try
    End Sub


    Private Sub SetItemHighlight(item As ListViewItem, isHighlighted As Boolean)
        If isHighlighted Then
            item.BackColor = highlightColor
            item.Font = highlightFont
            item.Selected = True
        Else
            item.BackColor = SystemColors.Window
            item.Font = normalFont

        End If
    End Sub
    Private Function GetPermanentEntryID(item As Object) As String
        Try
            If TypeOf item Is Outlook.MailItem Then
                Return DirectCast(item, Outlook.MailItem).EntryID
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Return DirectCast(item, Outlook.AppointmentItem).EntryID
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                Return DirectCast(item, Outlook.MeetingItem).EntryID
            End If
            Return String.Empty
        Catch ex As System.Exception
            Debug.WriteLine($"GetPermanentEntryID error: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Sub lvMails_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then Return

            Dim mailId As String = lvMails.SelectedItems(0).Tag.ToString()
            If String.IsNullOrEmpty(mailId) Then Return

            ' 更新高亮和内容
            If Not mailId.Equals(currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                UpdateHighlightByEntryID(currentMailEntryID, mailId)
                currentMailEntryID = mailId

                ' 异步加载邮件内容，避免阻塞UI
                LoadMailContentAsync(mailId)
            Else
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_SelectedIndexChanged error: {ex.Message}")
        End Try
    End Sub

    ' 异步加载邮件内容的方法
    Private Async Sub LoadMailContentAsync(mailId As String)
        Try
            ' 在UI线程显示加载状态
            wbContent.DocumentText = "<html><body>正在加载邮件内容...</body></html>"

            ' 在后台线程中执行耗时的Outlook操作
            Dim content As String = Await Task.Run(Function()
                                                        Try
                                                            Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailId)
                                                            If TypeOf currentItem Is Outlook.MailItem Then
                                                                Return MailHandler.DisplayMailContent(mailId)
                                                            ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                                                                Return MailHandler.DisplayMailContent(mailId)
                                                            ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                                                                Return MailHandler.DisplayMailContent(mailId)
                                                            Else
                                                                Return MailHandler.DisplayMailContent(mailId)
                                                            End If
                                                        Catch ex As System.Exception
                                                            Debug.WriteLine($"LoadMailContentAsync background error: {ex.Message}")
                                                            Return $"<html><body>加载邮件内容时出错: {ex.Message}</body></html>"
                                                        End Try
                                                    End Function)

            ' 回到UI线程更新内容
            If Me.InvokeRequired Then
                Me.Invoke(Sub() wbContent.DocumentText = content)
            Else
                wbContent.DocumentText = content
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"LoadMailContentAsync error: {ex.Message}")
            If Me.InvokeRequired Then
                Me.Invoke(Sub() wbContent.DocumentText = $"<html><body>加载邮件内容时出错: {ex.Message}</body></html>")
            Else
                wbContent.DocumentText = $"<html><body>加载邮件内容时出错: {ex.Message}</body></html>"
            End If
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
                Dim mailId As String = selectedItem.Tag.ToString()
                If Not String.IsNullOrEmpty(mailId) Then
                    Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailId)
                    If mailItem IsNot Nothing Then
                        mailItem.Display()
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
                Dim taskId As String = selectedItem.Tag.ToString()
                If Not String.IsNullOrEmpty(taskId) Then
                    Dim taskItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(taskId)
                    If taskItem IsNot Nothing Then
                        taskItem.Display()
                    End If
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine("TaskList_DoubleClick error: " & ex.Message)
        End Try
    End Sub
    Private Sub BtnAddTask_Click(sender As Object, e As EventArgs)
        Try
            If String.IsNullOrEmpty(currentConversationId) Then
                MessageBox.Show("请先选择一封邮件")
                Return
            End If

            OutlookAddIn3.Handlers.TaskHandler.CreateNewTask(currentConversationId, currentMailEntryID)
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

End Class
