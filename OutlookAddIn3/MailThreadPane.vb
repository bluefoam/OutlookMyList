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

    Private WithEvents lvMails As ListView
    Private WithEvents taskList As ListView
    Private wbContent As WebBrowser
    Private splitter1, splitter2 As SplitContainer
    Private tabControl As TabControl
    Private btnPanel As Panel
    Private currentConversationId As String = String.Empty
    Private currentMailEntryID As String = String.Empty
    Private currentHighlightIndex As Integer = -1  ' Add this line
    ' 删除这行，因为已经使用 mailItems 了
    Private mailIndexMap As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
    Private mailItems As New List(Of (Index As Integer, EntryID As String))  ' 移到这里
    ' 删除原来的 mailIndexMap

    Private Sub SetupControls()
        InitializeSplitContainers()
        SetupMailList()
        SetupMailContent()
        SetupTabPages()
        BindEvents()
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

    Private Sub SetupMailList()
        lvMails = New ListView With {
            .Dock = DockStyle.Fill,
            .View = Windows.Forms.View.Details,
            .FullRowSelect = True,
            .Sorting = SortOrder.Descending,
            .AllowColumnReorder = True,
            .SmallImageList = New ImageList()  ' 添加 ImageList
        }

        ' 设置图标列表属性
        lvMails.SmallImageList.ColorDepth = ColorDepth.Depth32Bit
        lvMails.SmallImageList.ImageSize = New Size(16, 16)

        Try
            ' 从资源加载图标

            Using bitmap As New Bitmap(My.Resources.mail_icon)
                Dim icon As Icon = Icon.FromHandle(bitmap.GetHicon())
                lvMails.SmallImageList.Images.Add("mail", icon)
            End Using

            Using bitmap As New Bitmap(My.Resources.calendar_icon)
                Dim icon As Icon = Icon.FromHandle(bitmap.GetHicon())
                lvMails.SmallImageList.Images.Add("calendar", icon)
            End Using

            Using bitmap As New Bitmap(My.Resources.meeting_icon)
                Dim icon As Icon = Icon.FromHandle(bitmap.GetHicon())
                lvMails.SmallImageList.Images.Add("meeting", icon)
            End Using

            Using bitmap As New Bitmap(My.Resources.other_icon)
                Dim icon As Icon = Icon.FromHandle(bitmap.GetHicon())
                lvMails.SmallImageList.Images.Add("other", icon)
            End Using

        Catch ex As System.Exception
            ' 如果资源加载失败，使用系统图标作为后备
            Debug.WriteLine($"从资源加载图标失败: {ex.Message}")
            lvMails.SmallImageList.Images.Add("mail", SystemIcons.Information)
            lvMails.SmallImageList.Images.Add("calendar", SystemIcons.Warning)
            lvMails.SmallImageList.Images.Add("meeting", SystemIcons.Application)
            lvMails.SmallImageList.Images.Add("other", SystemIcons.Question)
        End Try


        lvMails.Columns.Add("类型", 24)  ' 添加类型列，宽度刚好放置图标
        lvMails.Columns.Add("日期", 100)
        lvMails.Columns.Add("发件人", 100)
        lvMails.Columns.Add("主题", 200)
        splitter1.Panel1.Controls.Add(lvMails)
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

        SetupNotesTab()
        SetupTasksTab()
        SetupActionsTab()

        tabControl.SelectedIndex = 0
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
        AddHandler btnNewNote.Click, Async Sub(s, e)
                                         If Not String.IsNullOrEmpty(currentConversationId) Then
                                             Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
                                             Dim subject As String = ""
                                             If mailItem IsNot Nothing AndAlso TypeOf mailItem Is MailItem Then
                                                 subject = DirectCast(mailItem, MailItem).Subject
                                             End If
                                             Await SaveToWolaiAsync(currentConversationId, subject)
                                         Else
                                             MessageBox.Show("请先选择一封邮件")
                                         End If
                                     End Sub
        buttonPanel.Controls.Add(btnNewNote)

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

    Private Sub SetupNotesTab1()
        ' 首先检查 ComVisible 特性
        Dim isComVisible As Boolean = CheckComVisibleAttribute()
        If Not isComVisible Then
            Debug.WriteLine("警告: 当前类未标记 ComVisible(True)，这可能导致网页链接无法正常打开")
            ' 可以考虑添加一个提示
            MessageBox.Show("网页链接功能可能受限", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

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
        ' For the anonymous async handler
        AddHandler btnNewNote.Click, Async Sub(s, e)
                                         If Not String.IsNullOrEmpty(currentConversationId) Then
                                             Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
                                             Dim subject As String = ""
                                             If mailItem IsNot Nothing AndAlso TypeOf mailItem Is MailItem Then
                                                 subject = DirectCast(mailItem, MailItem).Subject
                                             End If
                                             Await SaveToWolaiAsync(currentConversationId, subject)
                                         Else
                                             MessageBox.Show("请先选择一封邮件")
                                         End If
                                     End Sub
        buttonPanel.Controls.Add(btnNewNote)

        ' 创建 WebBrowser
        infoWebBrowser = New WebBrowser With {
            .Dock = DockStyle.Fill,
            .ScrollBarsEnabled = True,
            .ScriptErrorsSuppressed = True,
            .AllowNavigation = True,
            .IsWebBrowserContextMenuEnabled = True,  ' 允许上下文菜单
            .WebBrowserShortcutsEnabled = True      ' 允许快捷键
        }

        ' 设置安全性和隐私设置
        Try
            infoWebBrowser.ObjectForScripting = Me
        Catch ex As System.Exception
            Debug.WriteLine($"设置 ObjectForScripting 失败: {ex.Message}")
        End Try

        Dim htmlContent As String = $"<html><body style='font-family: Arial; padding: 10px;'>" &
                                  $"<div id='entryId' style='margin-bottom: 10px;'></div>" &
                                  $"<div><a href='https://www.wolai.com/autolab/pLEYWMtYy4xFRzTyLEewrX' target='_blank' " &
                                  $"onclick='window.open(this.href); return false;'>所有笔记</a></div>" &
                                  $"</body></html>"
        infoWebBrowser.DocumentText = htmlContent
        ' 按正确的顺序添加控件
        containerPanel.Controls.Add(infoWebBrowser)
        containerPanel.Controls.Add(buttonPanel)
        tabPage1.Controls.Add(containerPanel)
        tabControl.TabPages.Add(tabPage1)
    End Sub

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

        Dim y As Integer = 10
        For i As Integer = 1 To 3
            Dim btn As New Button With {
                .Text = $"按钮 {i}",
                .Location = New Point(10, y),
                .Size = New Size(100, 30)
            }
            AddHandler btn.Click, Async Sub(s, e)
                                      Dim conversationTitle As String = "获取会话标题的逻辑"
                                      MessageBox.Show($"当前会话ID: {currentConversationId} 和 标题: {conversationTitle}")

                                  End Sub
            btnPanel.Controls.Add(btn)
            y += 40
        Next

        tabPage3.Controls.Add(btnPanel)
        tabControl.TabPages.Add(tabPage3)
    End Sub

    Private Function IsNetworkAvailable() As Boolean
        Try
            Return System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable()
        Catch ex As System.Exception
            Debug.WriteLine($"检查网络连接出错: {ex.Message}")
            Return False
        End Try
    End Function

    Private Async Function CheckWolaiRecordAsync_backup(conversationId As String) As Task(Of String)
        Try
            Using client As New HttpClient()
                ' 获取 token
                Dim tokenData As New JObject()
                tokenData.Add("appId", "2NdHab5WdUG995izevb69b")
                tokenData.Add("appSecret", "ffa888d4ebd73bae77a77abebcacf80001654b3f19d4ffbbcc3c41cbe0bed645")

                Dim tokenContent = New StringContent(tokenData.ToString(), Encoding.UTF8, "application/json")
                Dim tokenResponse = Await client.PostAsync("https://openapi.wolai.com/v1/token", tokenContent)

                If Not tokenResponse.IsSuccessStatusCode Then
                    Return String.Empty
                End If

                Dim tokenResult = Await tokenResponse.Content.ReadAsStringAsync()
                Dim tokenJson = JObject.Parse(tokenResult)
                Dim appToken = tokenJson.SelectToken("data.app_token")?.ToString()

                If String.IsNullOrEmpty(appToken) Then
                    Return String.Empty
                End If

                ' 查询数据
                client.DefaultRequestHeaders.Clear()
                client.DefaultRequestHeaders.Add("Authorization", appToken)

                ' 构建查询参数，直接过滤会话ID
                Dim queryData As New JObject()
                queryData.Add("filter", New JObject From {
                    {"property", "会话ID"},
                    {"value", conversationId},
                    {"type", "text"},
                    {"operator", "equals"}
                })

                Dim queryContent = New StringContent(queryData.ToString(), Encoding.UTF8, "application/json")
                Dim queryResponse = Await client.PostAsync("https://openapi.wolai.com/v1/databases/pLEYWMtYy4xFRzTyLEewrX/query", queryContent)

                Debug.WriteLine($"查询参数: {queryData}")

                If queryResponse.IsSuccessStatusCode Then
                    Dim responseContent = Await queryResponse.Content.ReadAsStringAsync()
                    Debug.WriteLine($"查询响应: {responseContent}")
                    Dim responseJson = JObject.Parse(responseContent)
                    Dim rows = responseJson.SelectToken("data")

                    If rows IsNot Nothing AndAlso rows.HasValues Then
                        ' 构建 HTML 表格
                        Dim htmlContent As New StringBuilder()
                        htmlContent.AppendLine("<html><body style='font-family: Arial; padding: 10px;'>")
                        'htmlContent.AppendLine("<h3>已存在的笔记记录：</h3>")
                        htmlContent.AppendLine("<table style='width: 100%; border-collapse: collapse; margin-bottom: 20px;'>")
                        htmlContent.AppendLine("<tr style='background-color: #f2f2f2;'>")
                        htmlContent.AppendLine("<th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>创建日期</th>")
                        htmlContent.AppendLine("<th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>标题</th>")
                        htmlContent.AppendLine("<th style='padding: 8px; border: 1px solid #ddd; text-align: left;'>操作</th>")
                        htmlContent.AppendLine("</tr>")

                        For Each row In rows
                            Dim pageId = row.ToString().Split("/"c).Last()
                            Dim wolaiLink = $"https://www.wolai.com/{pageId}"
                            Dim title = row.Parent.Parent("标题")?.ToString()
                            Dim createTime = row.Parent.Parent("创建时间")?.ToString()

                            htmlContent.AppendLine("<tr>")
                            htmlContent.AppendLine($"<td style='padding: 8px; border: 1px solid #ddd;'>{If(createTime, DateTime.Now.ToString("yyyy-MM-dd HH:mm"))}</td>")
                            htmlContent.AppendLine($"<td style='padding: 8px; border: 1px solid #ddd;'>{If(title, "无标题")}</td>")
                            htmlContent.AppendLine($"<td style='padding: 8px; border: 1px solid #ddd;'>")
                            htmlContent.AppendLine($"<a href='{wolaiLink}' target='_blank' onclick='window.open(this.href); return false;'>打开笔记</a>")
                            htmlContent.AppendLine("</td>")
                            htmlContent.AppendLine("</tr>")
                        Next

                        htmlContent.AppendLine("</table>")
                        htmlContent.AppendLine($"<div style='margin-top: 10px;'><a href='https://www.wolai.com/autolab/pLEYWMtYy4xFRzTyLEewrX' target='_blank' onclick='window.open(this.href); return false;'>所有笔记</a></div>")
                        htmlContent.AppendLine("</body></html>")

                        Return htmlContent.ToString()
                    End If
                End If

                Return String.Empty
            End Using
        Catch ex As System.Exception
            Debug.WriteLine($"CheckWolaiRecord 执行出错: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Private Async Function CheckWolaiRecordAsync(conversationId As String) As Task(Of String)
        Try
            Dim noteList As New List(Of (CreateTime As String, Title As String, Link As String))
            ' 首先检查所有相关邮件的属性
            Try
                ' 获取当前会话的所有邮件
                Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
                If TypeOf currentItem Is Outlook.MailItem Then
                    Dim conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
                    If conversation IsNot Nothing Then
                        Dim table As Outlook.Table = conversation.GetTable()
                        table.Columns.Add("EntryID")

                        ' 遍历会话中的所有邮件
                        Do Until table.EndOfTable
                            Dim row As Outlook.Row = table.GetNextRow()
                            Dim mailItem As Outlook.MailItem = DirectCast(Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString()), Outlook.MailItem)

                            If mailItem IsNot Nothing Then
                                Dim wolaiProp = mailItem.UserProperties.Find("WolaiNoteLink")
                                If wolaiProp IsNot Nothing Then
                                    Dim wolaiLink = wolaiProp.Value.ToString()
                                    Debug.WriteLine($"从邮件属性中找到 Wolai 链接: {wolaiLink}")
                                    ' 避免重复添加相同的链接
                                    If Not noteList.Any(Function(n) n.Link = wolaiLink) Then
                                        noteList.Add((DateTime.Now.ToString("yyyy-MM-dd HH:mm"), mailItem.Subject, wolaiLink))
                                    End If
                                End If
                            End If
                        Loop
                    End If
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"检查邮件属性时出错: {ex.Message}")
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
                            Dim mailItem As Outlook.MailItem = DirectCast(Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID), Outlook.MailItem)
                            If mailItem IsNot Nothing Then
                                mailItem.UserProperties.Add("WolaiNoteLink", Outlook.OlUserPropertyType.olText, True, Outlook.OlFormatText.olFormatTextText)
                                mailItem.UserProperties("WolaiNoteLink").Value = wolaiLink
                                mailItem.Save()
                                Debug.WriteLine($"已保存 Wolai 链接到邮件属性: {wolaiLink}")
                            End If
                        Catch ex As System.Exception
                            Debug.WriteLine($"保存链接到邮件属性时出错: {ex.Message}")
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
        If Not String.IsNullOrEmpty(currentConversationId) Then
            Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
            Dim subject As String = ""
            
            If mailItem IsNot Nothing AndAlso TypeOf mailItem Is MailItem Then
                subject = DirectCast(mailItem, MailItem).Subject
            End If
            Await SaveToWolaiAsync(currentConversationId, subject)
        Else
            MessageBox.Show("请先选择一封邮件")
        End If
    End Sub

    Private Sub SaveToWolai1(conversationId As String, conversationTitle As String)
        Try
            If infoWebBrowser.Document Is Nothing Then
                MessageBox.Show("WebBrowser 控件未准备就绪")
                Return
            End If

            Dim testScript As String = "
                function testFetch() {
                    return fetch('https://openapi.wolai.com/v1/token', {
                        'method': 'GET'
                    })
                    .then(response => 'Fetch API 可用')
                    .catch(error => 'Fetch API 不可用: ' + error);
                }
                testFetch().then(result => console.log(result));"

            ExecuteJavaScript(testScript)
            Debug.WriteLine($"JavaScript测试通过")

            ' 添加调试信息
            Debug.WriteLine($"开始执行保存操作 - 会话ID: {conversationId}, 标题: {conversationTitle}")

            Dim script As String = "javascript:(function (){" &
            "var newDiv = window.document.createElement('div');" &
            "newDiv.style.position = 'fixed';" &
            "newDiv.style.top = '10px';" &
            "newDiv.style.right = '10px';" &
            "newDiv.style.width = '200px';" &
            "newDiv.style.textAlign = 'center';" &
            "newDiv.style.height = '60px';" &
            "newDiv.style.padding = '10px';" &
            "newDiv.style.backgroundColor = 'white';" &
            "newDiv.style.border = '1px solid #ccc';" &
            "newDiv.style.boxShadow = '0 0 10px rgba(0,0,0,0.1)';" &
            "newDiv.style.zIndex = '9999';" &
            "newDiv.innerHTML = '保存到 Wolai 中...';" &
            "document.body.appendChild(newDiv);" &
            "login();" &
            "function login(){" &
            "var data = {" &
            "'appId': '2NdHab5WdUG995izevb69b'," &
            "'appSecret': 'ffa888d4ebd73bae77a77abebcacf80001654b3f19d4ffbbcc3c41cbe0bed645'" &
            "};" &
            "fetch('https://openapi.wolai.com/v1/token', {" &
            "'method': 'POST'," &
            "'headers': { 'Content-Type': 'application/json' }," &
            "'body': JSON.stringify(data)" &
            "})" &
            ".then(response => response.json())" &
            ".then(data => {" &
            "if(data?.data?.app_token) {" &
            "save(data.data.app_token);" &
            "}" &
            "})" &
            ".catch((error) => { alert('获取令牌失败'); });" &
            "}" &
            "function save(token){" &
            "var title = '" & conversationTitle & "';" &
            "var url = 'undefined';" &
            "var ConvetID = '" & conversationId & "';" &
            "const rows = { 'rows': [{'标题': title, '网址': url, '会话ID': ConvetID}] };" &
            "fetch('https://openapi.wolai.com/v1/databases/pLEYWMtYy4xFRzTyLEewrX/rows', {" &
            "'method': 'POST'," &
            "'headers': { 'Content-Type': 'application/json', 'Authorization': token }," &
            "'body': JSON.stringify(rows)" &
            "})" &
            ".then(response => response.json())" &
            ".then(data => {" &
            "console.log('保存成功:', data);" &
            "newDiv.innerHTML = '保存成功';" &
            "setTimeout(() => { newDiv.remove(); }, 2000);" &
            "})" &
            ".catch((error) => {" &
            "console.error('保存失败:', error);" &
            "newDiv.innerHTML = '保存失败: ' + error.message;" &
            "});" &
            "}" &
            "})();"

            Debug.WriteLine($"JavaScript 脚本内容:  {script}")

            ExecuteJavaScript(script)
            ' 添加执行完成的反馈
            Debug.WriteLine("JavaScript 脚本已执行")

        Catch ex As System.Exception
            Debug.WriteLine($"SaveToWolai 执行出错: {ex.Message}")
            MessageBox.Show($"保存失败: {ex.Message}")
        End Try
    End Sub

    Private Sub BindEvents()
        AddHandler lvMails.SelectedIndexChanged, AddressOf lvMails_SelectedIndexChanged
        AddHandler lvMails.ColumnClick, AddressOf lvMails_ColumnClick
        AddHandler lvMails.DoubleClick, AddressOf lvMails_DoubleClick
        AddHandler taskList.DoubleClick, AddressOf TaskList_DoubleClick
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

            If needReload Then
                LoadConversationMails(mailEntryID)
                Try
                    Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(mailEntryID)
                    If mailItem IsNot Nothing AndAlso TypeOf mailItem Is MailItem Then
                        ' 只在会话ID变化时更新笔记
                        Dim newConversationId = DirectCast(mailItem, MailItem).ConversationID
                        If Not String.Equals(newConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                            currentConversationId = newConversationId
                            ' 检查并更新笔记链接
                            Await CheckWolaiRecordAsync(currentConversationId)
                        End If
                    ElseIf mailItem IsNot Nothing AndAlso TypeOf mailItem Is AppointmentItem Then
                        Dim appointment As Outlook.AppointmentItem = DirectCast(mailItem, Outlook.AppointmentItem)
                        ' 使用GlobalAppointmentID作为会话ID
                        Dim newConversationId = appointment.GlobalAppointmentID
                        If Not String.Equals(newConversationId, currentConversationId, StringComparison.OrdinalIgnoreCase) Then
                            currentConversationId = newConversationId
                            ' 检查并更新笔记链接
                            Await CheckWolaiRecordAsync(currentConversationId)
                        End If

                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"更新会话ID时出错: {ex.Message}")
                End Try
            Else
                ' 只更新高亮和内容
                wbContent.DocumentText = MailHandler.DisplayMailContent(mailEntryID)
                UpdateHighlightByEntryID(currentMailEntryID, mailEntryID)
                currentHighlightIndex = GetIndexByEntryID(mailEntryID)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateMailList error: {ex.Message}")
        End Try
        currentMailEntryID = mailEntryID
    End Sub

    Private Function GetIndexByEntryID(entryID As String) As Integer
        Return mailItems.FindIndex(Function(x) String.Equals(x.EntryID, entryID.Trim(), StringComparison.OrdinalIgnoreCase))
    End Function

    Private Sub UpdateHighlightByEntryID(oldEntryID As String, newEntryID As String)
        Dim oldIndex As Integer = If(String.IsNullOrEmpty(oldEntryID), -1, GetIndexByEntryID(oldEntryID))
        Dim newIndex As Integer = If(String.IsNullOrEmpty(newEntryID), -1, GetIndexByEntryID(newEntryID))
        UpdateHighlightByMailId(oldIndex, newIndex)
    End Sub

    ' 在类级别添加字段
    Private Sub LoadConversationMails(currentMailEntryID As String)
        lvMails.BeginUpdate()
        Try
            lvMails.Items.Clear()
            mailItems.Clear()  ' 清空映射列表
            Dim currentItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(currentMailEntryID)
            Dim conversation As Outlook.Conversation = Nothing

            If TypeOf currentItem Is Outlook.MailItem Then
                conversation = DirectCast(currentItem, Outlook.MailItem).GetConversation()
            ElseIf TypeOf currentItem Is Outlook.AppointmentItem Then
                conversation = DirectCast(currentItem, Outlook.AppointmentItem).GetConversation()
            End If

            If conversation IsNot Nothing Then
                Dim table As Outlook.Table = conversation.GetTable()
                table.Columns.Add("EntryID")
                table.Columns.Add("SentOn")
                table.Columns.Add("ReceivedTime")  ' 添加接收时间
                table.Columns.Add("SenderName")
                table.Columns.Add("Subject")
                table.Columns.Add("MessageClass")  ' 添加消息类型
                table.Sort("[ReceivedTime]", True)  ' 修改为升序排序（最早的在前）
                Dim items As New List(Of ListViewItem)()
                currentHighlightIndex = -1

                ' 先收集所有项目
                Dim allItems As New List(Of (EntryID As String, ListItem As ListViewItem))
                Dim tempMailItems As New List(Of (Index As Integer, EntryID As String))

                Do Until table.EndOfTable
                    Dim row As Outlook.Row = table.GetNextRow()
                    Dim mailItem As Object = Globals.ThisAddIn.Application.Session.GetItemFromID(row("EntryID").ToString())

                    ' 跳过会议类型的项目
                    'If TypeOf mailItem Is Outlook.AppointmentItem Then
                    '    Continue Do
                    'End If

                    Dim entryId As String = GetPermanentEntryID(mailItem)
                    Dim currentIndex As Integer = allItems.Count

                    ' 创建 ListViewItem，第一列为空（仅显示图标）
                    Dim lvi As New ListViewItem("") With {
                        .Tag = currentIndex,
                        .ImageIndex = GetItemImageIndex(mailItem)  ' 设置图标
                    }

                    ' 添加时间列（作为排序索引）
                    Dim timeStr = If(row("ReceivedTime") IsNot Nothing AndAlso Not String.IsNullOrEmpty(row("ReceivedTime").ToString()),
                                   DateTime.Parse(row("ReceivedTime").ToString()).ToString("yyyy-MM-dd HH:mm"),
                                   "Unknown Date")
                    lvi.SubItems.Add(timeStr)

                    ' 添加其他列
                    lvi.SubItems.Add(If(row("SenderName") IsNot Nothing, row("SenderName").ToString(), "Unknown Sender"))
                    lvi.SubItems.Add(If(row("Subject") IsNot Nothing, row("Subject").ToString(), "Unknown Subject"))

                    allItems.Add((entryId, lvi))
                    tempMailItems.Add((currentIndex, entryId))

                    ' 检查是否是当前邮件
                    If String.Equals(entryId, currentMailEntryID.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        currentHighlightIndex = currentIndex
                    End If
                Loop
                ' 清空现有列表
                lvMails.Items.Clear()
                mailItems.Clear()

                ' 一次性添加所有项目
                lvMails.Items.AddRange(allItems.Select(Function(x) x.ListItem).ToArray())
                mailItems = tempMailItems

                ' 设置高亮
                If currentHighlightIndex >= 0 Then
                    UpdateHighlightByMailId(-1, currentHighlightIndex)
                End If
            End If
        Finally
            lvMails.EndUpdate()
        End Try
    End Sub

    ' 添加新方法用于确定图标索引
    Private Function GetItemImageIndex(item As Object) As Integer
        Try
            If TypeOf item Is Outlook.MailItem Then
                Return lvMails.SmallImageList.Images.IndexOfKey("mail")
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Return lvMails.SmallImageList.Images.IndexOfKey("calendar")
            ElseIf TypeOf item Is Outlook.MeetingItem Then
                Return lvMails.SmallImageList.Images.IndexOfKey("meeting")
            Else
                Return lvMails.SmallImageList.Images.IndexOfKey("other")
            End If
        Catch ex As system.Exception
            Debug.WriteLine($"获取图标索引出错: {ex.Message}")
            Return 0  ' 返回默认索引
        End Try
    End Function

    ' 添加类级别的字体缓存
    Private ReadOnly highlightFont As Font
    Private ReadOnly normalFont As Font
    Private ReadOnly highlightColor As Color = Color.FromArgb(255, 255, 200)

    Public Sub New()
        ' 这个调用是 Windows 窗体设计器所必需的。
        InitializeComponent()

        ' 在 InitializeComponent() 之后添加任何初始化代码
        normalFont = New Font(DefaultFont, FontStyle.Regular)
        highlightFont = New Font(DefaultFont, FontStyle.Bold)

        ' 最后设置控件
        SetupControls()
    End Sub

    Private Sub UpdateHighlightByMailId(oldIndex As Integer, newIndex As Integer)
        Try
            lvMails.BeginUpdate()

            ' 清除所有项的高亮状态
            For Each item As ListViewItem In lvMails.Items
                SetItemHighlight(item, False)
            Next

            ' 设置新的高亮
            If newIndex >= 0 AndAlso newIndex < lvMails.Items.Count Then
                Dim item = lvMails.Items(newIndex)
                SetItemHighlight(item, True)
                item.EnsureVisible()
            End If

        Finally
            lvMails.EndUpdate()
        End Try
    End Sub

    Private Sub UpdateHighlightByMailId1(oldIndex As Integer, newIndex As Integer)
        ' 如果索引无效或相同，快速返回
        If (oldIndex < 0 AndAlso newIndex < 0) OrElse
               (oldIndex >= lvMails.Items.Count AndAlso newIndex >= lvMails.Items.Count) Then
            Return
        End If

        ' 只有在真正需要更新时才使用 BeginUpdate
        If oldIndex <> newIndex Then
            lvMails.BeginUpdate()
        End If

        Try
            ' 清除旧的高亮
            If oldIndex >= 0 AndAlso oldIndex < lvMails.Items.Count Then
                SetItemHighlight(lvMails.Items(oldIndex), False)
            End If

            ' 设置新的高亮
            If newIndex >= 0 AndAlso newIndex < lvMails.Items.Count Then
                Dim item = lvMails.Items(newIndex)
                SetItemHighlight(item, True)
                item.EnsureVisible()
            End If
        Finally
            If oldIndex <> newIndex Then
                lvMails.EndUpdate()
            End If
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
            item.Selected = False
        End If
    End Sub
    Private Function GetPermanentEntryID(item As Object) As String
        Try
            If TypeOf item Is Outlook.MailItem Then
                Return DirectCast(item, Outlook.MailItem).EntryID
            ElseIf TypeOf item Is Outlook.AppointmentItem Then
                Return DirectCast(item, Outlook.AppointmentItem).EntryID
            End If
            Return String.Empty
        Catch ex As System.Exception
            Debug.WriteLine($"GetPermanentEntryID error: {ex.Message}")
            Return String.Empty
        End Try
    End Function
    Private Sub lvMails_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count = 0 Then
                Return
            End If



            Dim selectedIndex As Integer = CInt(lvMails.SelectedItems(0).Tag)
            If selectedIndex >= 0 AndAlso selectedIndex < mailItems.Count Then
                Dim mailId As String = mailItems(selectedIndex).EntryID
                If String.IsNullOrEmpty(mailId) Then
                    Return
                End If
                'Debug.WriteLine($"- EntryID: {mailId}")
                ' 更新高亮状态
                UpdateHighlightByMailId(currentHighlightIndex, selectedIndex)

                ' 如果不是同一封邮件，更新内容
                If Not mailId.Equals(currentMailEntryID, StringComparison.OrdinalIgnoreCase) Then
                    currentMailEntryID = mailId
                    wbContent.DocumentText = MailHandler.DisplayMailContent(mailId)
                End If

                currentHighlightIndex = selectedIndex
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_SelectedIndexChanged error: {ex.Message}")
        End Try
    End Sub
    Private Sub lvMails_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        Try
            Dim lv As ListView = DirectCast(sender, ListView)

            ' 记住当前的 EntryID
            Dim currentEntryID As String = If(currentHighlightIndex >= 0, mailItems(currentHighlightIndex).EntryID, String.Empty)

            lv.Sorting = If(lv.Sorting = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
            lv.ListViewItemSorter = New ListViewItemComparer(e.Column, lv.Sorting)

            ' 更新索引映射
            Dim newMailItems As New List(Of (Index As Integer, EntryID As String))
            For i As Integer = 0 To lv.Items.Count - 1
                Dim oldIndex As Integer = CInt(lv.Items(i).Tag)
                newMailItems.Add((i, mailItems(oldIndex).EntryID))
                lv.Items(i).Tag = i
            Next
            mailItems = newMailItems
            ' 根据 EntryID 查找新的索引位置
            If Not String.IsNullOrEmpty(currentEntryID) Then
                Dim newIndex = mailItems.FindIndex(Function(x) String.Equals(x.EntryID, currentEntryID, StringComparison.OrdinalIgnoreCase))
                If newIndex >= 0 Then
                    UpdateHighlightByMailId(currentHighlightIndex, newIndex)
                    currentHighlightIndex = newIndex
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"lvMails_ColumnClick error: {ex.Message}")
        End Try
    End Sub
    Private Sub lvMails_DoubleClick(sender As Object, e As EventArgs)
        Try
            If lvMails.SelectedItems.Count > 0 Then
                Dim index As Integer = CInt(lvMails.SelectedItems(0).Tag)
                If index >= 0 AndAlso index < mailItems.Count Then
                    Dim mailId As String = mailItems(index).EntryID
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
