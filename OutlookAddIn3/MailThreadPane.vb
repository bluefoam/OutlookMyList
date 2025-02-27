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


    Private WithEvents lvMails As ListView
    Private WithEvents taskList As ListView
    Private wbContent As WebBrowser
    Private splitter1, splitter2 As SplitContainer
    Private tabControl As TabControl
    Private btnPanel As Panel
    Private currentConversationId As String = String.Empty
    Private currentMailEntryID As String = String.Empty
    Private currentHighlightEntryID As String

    Private mailItems As New List(Of (Index As Integer, EntryID As String))  ' 移到这里
    ' 删除原来的 mailIndexMap

    Private Sub SetupControls()
        InitializeSplitContainers()
        SetupMailList()
        SetupMailContent()
        ' 延迟加载标签页
        Task.Run(Sub()
                     Threading.Thread.Sleep(500)  ' 给主界面一些加载时间
                     Me.Invoke(Sub()
                                   SetupTabPages()
                                   BindEvents()
                               End Sub)
                 End Sub)
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
            .OwnerDraw = True  ' 启用自定义绘制
        }

        lvMails.Columns.Add("----", 60)  ' 增加宽度以适应更大的图标
        lvMails.Columns.Add("日期", 100)
        lvMails.Columns.Add("发件人", 100)
        lvMails.Columns.Add("主题", 200)
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
                        If recentMails.Count >= 20 Then Exit For
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

            ' 按时间排序并显示最近5封
            For Each mail In recentMails.OrderByDescending(Function(m) m.ReceivedTime).Take(20)
                info.AppendLine($"- {mail.ReceivedTime:yyyy-MM-dd HH:mm} {mail.Subject}")
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
            If e.Url.ToString() <> "about:blank" Then
                e.Cancel = True  ' 取消 WebBrowser 的默认导航
                Debug.WriteLine($"正在尝试打开链接: {e.Url}")
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
            .Dock = DockStyle.Fill
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

        ' 只创建按钮，不预先创建文本框
        Dim x As Integer = 10
        For i As Integer = 1 To 3
            Dim btn As New Button With {
                .Text = If(i = 1, "查看联系人信息", $"按钮 {i}"),
                .Location = New Point(x, 10),
                .Size = New Size(120, 30)
            }

            ' 特别处理第一个按钮 - 延迟初始化
            If i = 1 Then
                AddHandler btn.Click, Sub(s, e)
                                          ' 第一次点击时才创建文本框
                                          If Not btnPanel.Controls.OfType(Of TextBox)().Any() Then
                                              Dim outputTextBox As New TextBox With {
                                                .Multiline = True,
                                                .ScrollBars = ScrollBars.Vertical,
                                                .Location = New Point(10, 45),
                                                .Size = New Size(350, 200),
                                                .ReadOnly = True
                                            }
                                              btnPanel.Controls.Add(outputTextBox)
                                          End If
                                          ' 获取文本框并执行操作
                                          Dim textBox = btnPanel.Controls.OfType(Of TextBox)().FirstOrDefault()
                                          If textBox IsNot Nothing Then
                                              GetContactInfoHandler(textBox)
                                          End If
                                      End Sub
            Else
                AddHandler btn.Click, Sub(s, e)
                                          Dim conversationTitle As String = "获取会话标题的逻辑"
                                          MessageBox.Show($"当前会话ID: {currentConversationId} 和 标题: {conversationTitle}")
                                      End Sub
            End If

            btnPanel.Controls.Add(btn)
            x += 125
        Next

        tabPage3.Controls.Add(btnPanel)
        tabControl.TabPages.Add(tabPage3)
    End Sub

    ' 将异步逻辑移到单独的方法中
    Private Async Sub GetContactInfoHandler(outputTextBox As TextBox)
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub() outputTextBox.Text = "正在收集联系人信息...")
            Else
                outputTextBox.Text = "正在收集联系人信息..."
            End If

            Dim info = Await GetContactInfoAsync()

            If Me.InvokeRequired Then
                Me.Invoke(Sub()
                              If Not String.IsNullOrEmpty(info) Then
                                  outputTextBox.Text = info
                              Else
                                  outputTextBox.Text = "未能获取联系人信息"
                              End If
                          End Sub)
            Else
                If Not String.IsNullOrEmpty(info) Then
                    outputTextBox.Text = info
                Else
                    outputTextBox.Text = "未能获取联系人信息"
                End If
            End If
        Catch ex As System.Exception
            If Me.InvokeRequired Then
                Me.Invoke(Sub() outputTextBox.Text = $"获取联系人信息时出错: {ex.Message}")
            Else
                outputTextBox.Text = $"获取联系人信息时出错: {ex.Message}"
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
            If Not IsNetworkAvailable() Then
                Debug.WriteLine("网络不可用，跳过网络查询")
                UpdateNoteList(noteList)
                Return String.Empty
            End If

            Using client As New HttpClient()
                ' 获取 token
                Dim tokenData As New JObject()
                tokenData.Add("appId", "2NdHab5WdUG995izevb69b")
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
                    {"property", "会话ID"},
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
                            Dim title = row.Parent.Parent("标题")?.ToString()
                            Dim createTime = row.Parent.Parent("创建时间")?.ToString()
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
                row.Add("标题", conversationTitle)
                row.Add("网址", "undefined")
                row.Add("会话ID", conversationId)
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
        'If Not String.IsNullOrEmpty(currentConversationId) Then
        Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
        Dim subject As String = ""

        If mailItem IsNot Nothing Then
            ' 根据不同类型获取主题
            If TypeOf mailItem Is Outlook.MailItem Then
                subject = DirectCast(mailItem, Outlook.MailItem).Subject
            ElseIf TypeOf mailItem Is Outlook.AppointmentItem Then
                subject = DirectCast(mailItem, Outlook.AppointmentItem).Subject
            ElseIf TypeOf mailItem Is Outlook.MeetingItem Then
                subject = DirectCast(mailItem, Outlook.MeetingItem).Subject
            ElseIf TypeOf mailItem Is Outlook.TaskItem Then
                subject = DirectCast(mailItem, Outlook.TaskItem).Subject
            End If
        End If

        Await SaveToWolaiAsync(currentConversationId, subject)
        'Else
        'MessageBox.Show("请先选择一封邮件")
        'End If
    End Sub

    Private Sub BindEvents()
        AddHandler lvMails.SelectedIndexChanged, AddressOf lvMails_SelectedIndexChanged
        AddHandler lvMails.ColumnClick, AddressOf lvMails_ColumnClick
        AddHandler lvMails.DoubleClick, AddressOf lvMails_DoubleClick

    End Sub

    Public Async Sub UpdateMailList(conversationId As String, mailEntryID As String)
        If lvMails Is Nothing Then
            SetupControls()
        End If

        Try
            If String.IsNullOrEmpty(mailEntryID) Then
                lvMails.Items.Clear()
                Return
            End If

            ' 检查是否需要重新加载列表
            Dim needReload As Boolean = True
            If lvMails.Items.Count > 0 AndAlso Not String.IsNullOrEmpty(conversationId) AndAlso
               String.Equals(conversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                needReload = False
            End If

            If Not String.IsNullOrEmpty(mailEntryID) AndAlso String.IsNullOrEmpty(conversationId) Then
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailEntryID)
            End If

            If needReload Then
                ' 暂时移除事件处理器，避免重复触发
                'RemoveHandler lvMails.SelectedIndexChanged, AddressOf lvMails_SelectedIndexChanged

                LoadConversationMails(mailEntryID)
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
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateMailList error: {ex.Message}")
        End Try
        currentMailEntryID = mailEntryID
    End Sub

    Private Function GetIndexByEntryID(entryID As String) As Integer
        Return mailItems.FindIndex(Function(x) String.Equals(x.EntryID, entryID.Trim(), StringComparison.OrdinalIgnoreCase))
    End Function

    ' 在listview_Mailist添加构造列表
    Private Sub LoadConversationMails(currentMailEntryID As String)
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
        iconFont = New Font("Segoe UI Emoji", 10)  ' 使用 Segoe UI Emoji 字体以获得更好的 emoji 显示效果
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

                ' 获取当前选中项的内容
                Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailId)
                If TypeOf currentItem Is Outlook.MailItem Then
                    wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
                ElseIf TypeOf currentItem Is Outlook.MeetingItem Then
                    ' 对于会议项目，尝试获取关联的邮件
                    'Dim meetingItem = DirectCast(currentItem, Outlook.MeetingItem)
                    'Dim associatedMail As Outlook.MailItem = meetingItem.GetAssociatedItem()
                    'If associatedMail IsNot Nothing Then
                    '    wbContent.DocumentText = MailHandler.DisplayMailContent(associatedMail.EntryID)
                    'Else
                    wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
                    'End If
                ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                    ' 对于约会项目，尝试查找相关的会议请求邮件
                    'Dim appointmentItem = DirectCast(currentItem, Outlook.AppointmentItem)
                    'Try
                    '    ' 使用 GetAssociatedAppointment 方法获取关联的会议请求
                    '    Dim namespace1 = Globals.ThisAddIn.Application.Session
                    '    Dim inbox = namespace1.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                    '    ' 使用会议组织者和主题来查找相关邮件
                    '    Dim filter = $"[MessageClass]='IPM.Schedule.Meeting.Request' AND " &
                    '               $"[Subject] LIKE '%{appointmentItem.Subject}%' AND " &
                    '               $"[SenderName] = '{appointmentItem.Organizer}'"
                    '    Dim items = inbox.Items.Restrict(filter)
                    '
                    '    If items.Count > 0 Then
                    '        Dim meetingMail As Outlook.MailItem = items.GetFirst()
                    '        wbContent.DocumentText = MailHandler.DisplayMailContent(meetingMail.EntryID)
                    '    Else
                    '        wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
                    '    End If
                    'Catch ex As System.Exception
                    '    Debug.WriteLine($"查找会议邮件时出错: {ex.Message}")
                    wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
                End If
            Else
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_SelectedIndexChanged error: {ex.Message}")
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

    Private Sub lvMails_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        Try
            Dim lv As ListView = DirectCast(sender, ListView)

            ' 切换排序方向
            lv.Sorting = If(lv.Sorting = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)

            ' 使用自定义排序器
            lv.ListViewItemSorter = New ListViewItemComparer(e.Column, lv.Sorting)
            lv.Sort()

            ' 更新高亮
            If Not String.IsNullOrEmpty(currentMailEntryID) Then
                UpdateHighlightByEntryID(String.Empty, currentMailEntryID)
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_ColumnClick error: {ex.Message}")
        End Try
    End Sub

    Private Sub lvMails_DoubleClick(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count > 0 Then
                Dim mailId As String = lvMails.SelectedItems(0).Tag.ToString()
                If Not String.IsNullOrEmpty(mailId) Then
                    Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailId)
                    If mailItem IsNot Nothing Then
                        mailItem.Display()
                    End If
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_DoubleClick error: {ex.Message}")
        End Try
    End Sub

    Private Sub TaskList_DoubleClick(sender As Object, e As EventArgs)
        Try
            If taskList.SelectedItems.Count > 0 Then
                Dim taskId As String = taskList.SelectedItems(0).Tag.ToString()
                If Not String.IsNullOrEmpty(taskId) Then
                    Dim taskItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(taskId)
                    If taskItem IsNot Nothing Then
                        taskItem.Display()
                    End If
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"TaskList_DoubleClick error: {ex.Message}")
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
            Debug.WriteLine($"BtnAddTask_Click error: {ex.Message}")
            MessageBox.Show($"创建任务时出错: {ex.Message}")
        End Try
    End Sub



End Class
