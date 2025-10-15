Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms   ' 添加这行

Namespace OutlookMyList.Handlers
    Public Class MailHandler
    ''' <summary>
    ''' 将ListView项目的Tag转换为EntryID字符串
    ''' </summary>
    ''' <param name="tag">ListView项目的Tag对象</param>
    ''' <returns>EntryID字符串</returns>
    Private Shared Function ConvertEntryIDToString(tag As Object) As String
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
    Public Shared Function GetPermanentEntryID(item As Object) As String
        Try
            If TypeOf item Is MailItem Then
                Return DirectCast(item, MailItem).EntryID
            ElseIf TypeOf item Is AppointmentItem Then
                Return DirectCast(item, AppointmentItem).EntryID
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"GetPermanentEntryID error: {ex.Message}")
        End Try
        Return String.Empty
    End Function

    Public Shared Sub UpdateHighlight(lvMails As ListView, currentMailEntryID As String, Optional backgroundColor As Color = Nothing)
        If backgroundColor = Nothing Then backgroundColor = SystemColors.Window
        
        lvMails.BeginUpdate()
        Try
            For Each item As ListViewItem In lvMails.Items
                Dim itemEntryID As String = ConvertEntryIDToString(item.Tag).Trim()
                Dim currentEntryID As String = currentMailEntryID.Trim()
                If String.Equals(itemEntryID, currentEntryID, StringComparison.OrdinalIgnoreCase) Then
                    item.BackColor = System.Drawing.Color.FromArgb(255, 255, 200)
                    item.Font = New System.Drawing.Font(lvMails.Font, System.Drawing.FontStyle.Bold)
                    item.Selected = True
                    item.EnsureVisible()
                Else
                    item.BackColor = backgroundColor  ' 使用传入的背景色
                    item.Font = lvMails.Font
                    item.Selected = False
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"UpdateHighlight error: {ex.Message}")
        Finally
            lvMails.EndUpdate()
        End Try
    End Sub

    Private Shared Function GetTaskStatus(status As OlTaskStatus) As String
        Select Case status
            Case OlTaskStatus.olTaskComplete
                Return "已完成"
            Case OlTaskStatus.olTaskDeferred
                Return "已推迟"
            Case OlTaskStatus.olTaskInProgress
                Return "进行中"
            Case OlTaskStatus.olTaskNotStarted
                Return "未开始"
            Case OlTaskStatus.olTaskWaiting
                Return "等待中"
            Case Else
                Return "未知状态"
        End Select
    End Function

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

    Public Shared Function DisplayMailContent(mailEntryID As String) As String
        Debug.WriteLine($"[DisplayMailContent] 开始处理邮件，EntryID: {mailEntryID}")
        Try
            Dim mail As Object = OutlookMyList.Utils.OutlookUtils.SafeGetItemFromID(mailEntryID)
            If mail Is Nothing Then
                Debug.WriteLine("[DisplayMailContent] 无法获取邮件对象")
                Return "<html><body><p>无法获取邮件</p></body></html>"
            End If
            Debug.WriteLine($"[DisplayMailContent] 成功获取邮件对象，类型: {mail.GetType().Name}")

            If TypeOf mail Is MailItem Then
                Dim mailItem As MailItem = DirectCast(mail, MailItem)
                
                ' 使用全局主题变量
                Dim bgColor As String = MailThreadPane.globalThemeBackgroundColor
                Dim fgColor As String = MailThreadPane.globalThemeForegroundColor
                Dim accentColor As String = MailThreadPane.globalThemeAccentColor
                
                ' 检查是否为默认的白色主题，如果是则使用安全的默认值，避免在后台线程中刷新主题
                If bgColor = "#ffffff" AndAlso fgColor = "#000000" AndAlso MailThreadPane.globalThemeLastUpdate = DateTime.MinValue Then
                    Debug.WriteLine("[DisplayMailContent] 检测到默认主题，使用安全的默认主题值，避免在后台线程中刷新主题")
                    ' 使用安全的默认主题值，避免在后台线程中调用RefreshTheme导致COM异常
                    ' 主题刷新应该在主线程中进行，这里只使用默认值
                    bgColor = "#ffffff"  ' 白色背景
                    fgColor = "#000000"  ' 黑色文字
                    accentColor = "#0078d4"  ' 蓝色强调色
                    Debug.WriteLine($"[DisplayMailContent] 使用默认主题值: 背景={bgColor}, 前景={fgColor}")
                    
                    ' 异步通知主线程进行主题刷新，但不等待结果
                    Try
                        System.Threading.Tasks.Task.Run(Sub()
                             Try
                                 ' 在主线程中异步刷新主题
                                 If Globals.ThisAddIn.MailThreadPaneInstance IsNot Nothing AndAlso Globals.ThisAddIn.MailThreadPaneInstance.IsHandleCreated Then
                                     Globals.ThisAddIn.MailThreadPaneInstance.BeginInvoke(Sub()
                                         Try
                                             Globals.ThisAddIn.RefreshTheme()
                                         Catch ex As System.Exception
                                             Debug.WriteLine($"[DisplayMailContent] 异步主题刷新失败: {ex.Message}")
                                         End Try
                                     End Sub)
                                 End If
                             Catch ex As System.Exception
                                 Debug.WriteLine($"[DisplayMailContent] 异步主题刷新调度失败: {ex.Message}")
                             End Try
                         End Sub)
                    Catch ex As System.Exception
                        Debug.WriteLine($"[DisplayMailContent] 启动异步主题刷新失败: {ex.Message}")
                    End Try
                End If
                
                ' 添加调试信息
                Debug.WriteLine($"[DisplayMailContent] 全局主题变量值:")
                Debug.WriteLine($"  背景色: {bgColor}")
                Debug.WriteLine($"  前景色: {fgColor}")
                Debug.WriteLine($"  强调色: {accentColor}")
                Debug.WriteLine($"  最后更新时间: {MailThreadPane.globalThemeLastUpdate}")
                
                ' 检查邮件项是否已完全加载
                If Not OutlookUtils.IsMailItemReady(mailItem) Then
                    Debug.WriteLine("邮件项未完全加载，跳过内容提取")
                    Return "<html><body><p>邮件内容加载中...</p></body></html>"
                End If

                Try
                    Dim subject As String = If(String.IsNullOrEmpty(mailItem.Subject), "无主题", mailItem.Subject)
                    Dim senderName As String = If(String.IsNullOrEmpty(mailItem.SenderName), "未知", mailItem.SenderName)
                    Dim receivedTime As String = If(mailItem.ReceivedTime = DateTime.MinValue, "未知", mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim htmlBody As String = If(String.IsNullOrEmpty(mailItem.HTMLBody), "", ReplaceTableTag(mailItem.HTMLBody))
                    
                    Dim htmlContent As String = $"<html><head><style>
                        /* 强制覆盖所有元素的背景和文字颜色 */
                        * {{
                            background-color: {bgColor} !important;
                            color: {fgColor} !important;
                        }}
                        body {{
                            background-color: {bgColor} !important;
                            color: {fgColor} !important;
                            font-family: Arial, sans-serif;
                            padding: 10px;
                            font-size: 12px;
                            margin: 0;
                        }}
                        /* 覆盖所有可能的文本元素 */
                        p, div, span, td, th, li, ul, ol, a, em, i, b, strong, h1, h2, h3, h4, h5, h6 {{
                            color: {fgColor} !important;
                            background-color: transparent !important;
                        }}
                        /* 标题和强调文本使用强调色 */
                        h1, h2, h3, h4, h5, h6 {{
                            color: {accentColor} !important;
                            margin-top: 0;
                        }}
                        strong, b {{
                            color: {accentColor} !important;
                        }}
                        /* 链接颜色 */
                        a, a:visited, a:hover, a:active {{
                            color: {accentColor} !important;
                        }}
                        /* 边框颜色 */
                        div, table, td, th {{
                            border-color: {accentColor} !important;
                        }}
                        /* 隐藏不需要的元素 */
                        .hidden-table {{
                            display: none;
                        }}
                        img {{
                            display: none;
                        }}
                        /* 文本选择样式 - 提供良好的对比度 */
                        ::selection {{
                            background-color: {fgColor} !important;
                            color: {bgColor} !important;
                        }}
                        ::-moz-selection {{
                            background-color: {fgColor} !important;
                            color: {bgColor} !important;
                        }}
                        /* 覆盖表格样式 */
                        table, tbody, thead, tfoot, tr, td, th {{
                            background-color: transparent !important;
                            color: {fgColor} !important;
                        }}
                        /* 覆盖列表样式 */
                        ul, ol, li {{
                            color: {fgColor} !important;
                        }}
                    </style></head><body>" &
                           $"<h4>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px; font-size: 12px;'>" &
                           $"<strong>发件人:</strong> {senderName}<br/>" &
                           $"<strong>时间:</strong> {receivedTime}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid {accentColor}; padding-top: 10px;'>" &
                           $"{htmlBody}" &
                           $"</div>" &
                           "</body></html>"
                    Debug.WriteLine($"[DisplayMailContent] 返回邮件HTML内容，长度: {htmlContent.Length}")
                    If htmlContent.Length > 0 Then
                        Dim preview = If(htmlContent.Length > 200, htmlContent.Substring(0, 200), htmlContent)
                        Debug.WriteLine($"[DisplayMailContent] HTML内容预览: {preview}")
                    End If
                    Return htmlContent
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问邮件属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问邮件属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问邮件属性时发生异常: {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问邮件属性</body></html>"
                End Try

                       '"<div style='border-top: 1px solid #ccc; padding-top: 10px;'> <style> .hidden-table {display: none;} img {display: none;}</style>"
            ElseIf TypeOf mail Is AppointmentItem Then
                Dim appointment As AppointmentItem = DirectCast(mail, AppointmentItem)
                
                ' 使用全局主题变量
                Dim bgColor As String = MailThreadPane.globalThemeBackgroundColor
                Dim fgColor As String = MailThreadPane.globalThemeForegroundColor
                Dim accentColor As String = MailThreadPane.globalThemeAccentColor
                
                Try
                    Dim subject As String = If(String.IsNullOrEmpty(appointment.Subject), "无主题", appointment.Subject)
                    Dim organizer As String = If(String.IsNullOrEmpty(appointment.Organizer), "未知", appointment.Organizer)
                    Dim startTime As String = If(appointment.Start = DateTime.MinValue, "未设置", appointment.Start.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim endTime As String = If(appointment.End = DateTime.MinValue, "未设置", appointment.End.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim location As String = If(String.IsNullOrEmpty(appointment.Location), "未设置", appointment.Location)
                    Dim body As String = If(String.IsNullOrEmpty(appointment.Body), "", appointment.Body.Replace(vbCrLf, "<br>").Replace("<br><br>", "<br>"))
                    
                    Return $"<html><head><style>
                        body {{
                            background-color: {bgColor} !important;
                            color: {fgColor} !important;
                            font-family: Arial, sans-serif;
                            padding: 10px;
                            font-size: 12px;
                            margin: 0;
                        }}
                        h4 {{
                            color: {accentColor} !important;
                            margin-top: 0;
                        }}
                        strong {{
                            color: {accentColor} !important;
                        }}
                        div {{
                            border-color: {accentColor} !important;
                        }}
                    </style></head><body>" &
                           $"<h4>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px; font-size: 12px;'>" &
                           $"<strong>组织者:</strong> {organizer}<br/>" &
                           $"<strong>开始时间:</strong> {startTime}<br/>" &
                           $"<strong>结束时间:</strong> {endTime}<br/>" &
                           $"<strong>地点:</strong> {location}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid {accentColor}; padding-top: 10px; font-size: 12px;'>" &
                           $"{body}" &
                           $"</div></body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问会议属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问会议属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问会议属性时发生异常: {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问会议属性</body></html>"
                End Try
            ElseIf TypeOf mail Is TaskItem Then
                Dim task As TaskItem = DirectCast(mail, TaskItem)
                
                ' 使用全局主题变量
                Dim bgColor As String = MailThreadPane.globalThemeBackgroundColor
                Dim fgColor As String = MailThreadPane.globalThemeForegroundColor
                Dim accentColor As String = MailThreadPane.globalThemeAccentColor
                
                Try
                    Dim subject As String = If(String.IsNullOrEmpty(task.Subject), "无主题", task.Subject)
                    Dim startDate As String = If(task.StartDate = DateTime.MinValue, "未设置", task.StartDate.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim dueDate As String = If(task.DueDate = DateTime.MinValue, "未设置", task.DueDate.ToString("yyyy-MM-dd HH:mm:ss"))
                    Dim percentComplete As Integer = task.PercentComplete
                    Dim status As String = GetTaskStatus(task.Status)
                    Dim body As String = If(String.IsNullOrEmpty(task.Body), "", task.Body.Replace(vbCrLf, "<br>").Replace("<br><br>", "<br>"))
                    
                    Return $"<html><head><style>
                        body {{
                            background-color: {bgColor} !important;
                            color: {fgColor} !important;
                            font-family: Arial, sans-serif;
                            padding: 10px;
                            font-size: 12px;
                            margin: 0;
                        }}
                        h4 {{
                            color: {accentColor} !important;
                            margin-top: 0;
                        }}
                        strong {{
                            color: {accentColor} !important;
                        }}
                        div {{
                            border-color: {accentColor} !important;
                        }}
                    </style></head><body>" &
                           $"<h4>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px; font-size: 12px;'>" &
                           $"<strong>开始时间:</strong> {startDate}<br/>" &
                           $"<strong>结束时间:</strong> {dueDate}<br/>" &
                           $"<strong>完成度:</strong> {percentComplete}%<br/>" &
                           $"<strong>状态:</strong> {status}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid {accentColor}; padding-top: 10px; font-size: 12px;'>" &
                           $"{body}" &
                           $"</div></body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问任务属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问任务属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问任务属性时发生异常: {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问任务属性</body></html>"
                End Try
            ElseIf TypeOf mail Is MeetingItem Then
                Dim meeting As MeetingItem = DirectCast(mail, MeetingItem)
                
                ' 使用全局主题变量
                Dim bgColor As String = MailThreadPane.globalThemeBackgroundColor
                Dim fgColor As String = MailThreadPane.globalThemeForegroundColor
                Dim accentColor As String = MailThreadPane.globalThemeAccentColor

                ' 获取关联的约会项目以获取时间信息
                Dim associatedAppointment As AppointmentItem = meeting.GetAssociatedAppointment(False)

                ' 获取会议状态信息
                Dim meetingStatus As String = "会议邀请"
                Select Case meeting.MessageClass
                    Case "IPM.Schedule.Meeting.Canceled"
                        meetingStatus = "会议已取消"
                        Return "<html><body><p>会议已取消, 无法显示内容</p></body></html>"
                    Case "IPM.Schedule.Meeting.Request"
                        meetingStatus = "会议邀请"
                    Case "IPM.Schedule.Meeting.Resp.Pos"
                        meetingStatus = "已接受"
                    Case "IPM.Schedule.Meeting.Resp.Neg"
                        meetingStatus = "已拒绝"
                    Case "IPM.Schedule.Meeting.Resp.Tent"
                        meetingStatus = "暂定"
                End Select


                Try
                    Dim subject As String = If(String.IsNullOrEmpty(meeting.Subject), "无主题", meeting.Subject)
                    Dim senderName As String = If(String.IsNullOrEmpty(meeting.SenderName), "未知", meeting.SenderName)
                    Dim body As String = If(String.IsNullOrEmpty(meeting.Body), "", meeting.Body.Replace(vbCrLf, "<br>").Replace(" <br><br>", "<br>"))
                    
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; font-size: 12px; }}</style></head><body>" &
                           $"<h4>{subject}</h4>" &
                           $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                           $"<strong>状态:</strong> {meetingStatus}<br/>" &
                           $"<strong>发件人:</strong> {senderName}<br/>" &
                           $"<strong>开始时间:</strong> {If(associatedAppointment IsNot Nothing AndAlso associatedAppointment.Start <> DateTime.MinValue,
                                                          associatedAppointment.Start.ToString("yyyy-MM-dd HH:mm:ss"),
                                                          "未设置")}<br/>" &
                           $"<strong>结束时间:</strong> {If(associatedAppointment IsNot Nothing AndAlso associatedAppointment.End <> DateTime.MinValue,
                                                          associatedAppointment.End.ToString("yyyy-MM-dd HH:mm:ss"),
                                                          "未设置")}<br/>" &
                           $"<strong>地点:</strong> {If(associatedAppointment IsNot Nothing AndAlso Not String.IsNullOrEmpty(associatedAppointment.Location),
                                                     associatedAppointment.Location,
                                                     "未设置")}" &
                           $"</div>" &
                           $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size:12px;'>" &
                           $"{body}" &
                           $"</div></body></html>"
                Catch ex As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"COM异常访问会议属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问会议属性</body></html>"
                Catch ex As System.Exception
                    Debug.WriteLine($"访问会议属性时发生异常: {ex.Message}")
                    Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问会议属性</body></html>"
                End Try

                If 0 Then
                    Return $"<html><body style='font-family: Arial; padding: 10px;Font-size:12px;'>" &
                       $"<h4>{meeting.Subject}</h4>" &
                       $"<div style='margin-bottom: 10px;Font-size:12px;'>" &
                       $"<strong>开始时间:</strong> {meeting.Start:yyyy-MM-dd HH:mm:ss}<br/>" &
                       $"<strong>结束时间:</strong> {meeting.End:yyyy-MM-dd HH:mm:ss}" &
                       $"</div>" &
                       $"<div style='border-top: 1px solid #ccc; padding-top: 10px;Font-size:12px;'>" &
                       $"{meeting.HTMLBody.Replace(vbCrLf, "<br>").Replace(" <br><br>", "<br>")}" &
                       $"</div></body></html>"
                End If
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"显示邮件内容时出错: {ex.Message}")
        End Try
        ' 使用全局主题变量作为默认值
        Dim defaultBgColor As String = MailThreadPane.globalThemeBackgroundColor
        Dim defaultFgColor As String = MailThreadPane.globalThemeForegroundColor
        Dim defaultHtml As String = $"<html><head><style>body {{ background-color: {defaultBgColor} !important; color: {defaultFgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body><p>无法显示内容</p></body></html>"
        Debug.WriteLine($"[DisplayMailContent] 返回默认HTML内容，长度: {defaultHtml.Length}")
        Return defaultHtml
    End Function
    Public Shared Sub OpenLink(url As String)
        Try
            Process.Start(New ProcessStartInfo With {
                .FileName = url,
                .UseShellExecute = True
            })
        Catch ex As System.Exception
            Debug.WriteLine($"打开链接出错: {ex.Message}")
            If ErrorNotificationSettings.Instance.ShowErrorDialogs Then
                MessageBox.Show("无法打开链接，请手动复制链接地址到浏览器中打开。")
            End If
        End Try
    End Sub
End Class
End Namespace