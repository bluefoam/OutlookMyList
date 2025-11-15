Imports Microsoft.Office.Interop.Outlook
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports OutlookMyList.Utils

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

                    ' 实时获取当前系统主题色，不依赖全局变量
                    Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentSystemThemeColors()
                    Dim bgColor As String = $"#{currentThemeColors.backgroundColor.R:X2}{currentThemeColors.backgroundColor.G:X2}{currentThemeColors.backgroundColor.B:X2}"
                    Dim fgColor As String = $"#{currentThemeColors.foregroundColor.R:X2}{currentThemeColors.foregroundColor.G:X2}{currentThemeColors.foregroundColor.B:X2}"
                    Dim accentColor As String = "#0078d4" ' 使用默认的蓝色强调色

                    Debug.WriteLine($"[DisplayMailContent] 实时获取主题色: 背景={bgColor}, 前景={fgColor}")

                    ' 添加调试信息
                    Debug.WriteLine($"[DisplayMailContent] 全局主题变量值:")
                    Debug.WriteLine($"  背景色: {bgColor}")
                    Debug.WriteLine($"  前景色: {fgColor}")
                    Debug.WriteLine($"  强调色: {accentColor}")
                    Debug.WriteLine($"  最后更新时间: {MailThreadPane.globalThemeLastUpdate}")

                    ' 检查邮件项是否已完全加载
                    If Not OutlookMyList.Utils.OutlookUtils.IsMailItemReady(mailItem) Then
                        Debug.WriteLine("邮件项未完全加载，跳过内容提取")
                        Return "<html><body><p>邮件内容加载中...</p></body></html>"
                    End If

                    Try
                        Dim subject As String = If(String.IsNullOrEmpty(mailItem.Subject), "无主题", mailItem.Subject)
                        Dim senderName As String = If(String.IsNullOrEmpty(mailItem.SenderName), "未知", mailItem.SenderName)
                        Dim receivedTime As String = If(mailItem.ReceivedTime = DateTime.MinValue, "未知", mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss"))
                        Dim rawHtml As String = If(String.IsNullOrEmpty(mailItem.HTMLBody), "", ReplaceTableTag(mailItem.HTMLBody))
                        Dim htmlBody As String = SanitizeHtmlBody(rawHtml)

                        Dim htmlContent As String = "<html><head><style>html, body { background-color: " & bgColor & " !important; color: " & fgColor & " !important; font-family: Arial, sans-serif; padding: 10px; font-size: 12px; margin: 0 !important; } * { background-color: " & bgColor & " !important; color: " & fgColor & " !important; } h4 { color: " & accentColor & " !important; margin-top: 0; } </style></head><body>" & _
                            "<h4 style='color: " & accentColor & " !important;'>" & subject & "</h4>" & _
                            "<div style='margin-bottom: 10px; font-size: 12px; color: " & fgColor & " !important;'>" & _
                            "<strong>发件人:</strong> " & senderName & "<br/>" & _
                            "<strong>时间:</strong> " & receivedTime & _
                            "</div>" & _
                            "<div style='border-top: 1px solid " & accentColor & "; padding-top: 10px; color: " & fgColor & " !important;'>" & _
                            htmlBody & _
                            "</div></body></html>"
                        Return htmlContent
                    Catch ex As System.Runtime.InteropServices.COMException
                        Debug.WriteLine($"COM异常访问会议属性 (HRESULT: {ex.HResult:X8}): {ex.Message}")
                        Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问会议属性</body></html>"
                    Catch ex As System.Exception
                        Debug.WriteLine($"访问会议属性时发生异常: {ex.Message}")
                        Return $"<html><head><style>body {{ background-color: {bgColor} !important; color: {fgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body>无法访问会议属性</body></html>"
                    End Try
                ElseIf TypeOf mail Is TaskItem Then
                    Dim task As TaskItem = DirectCast(mail, TaskItem)

                    ' 实时获取当前系统主题色
                    Dim themeColors2 As (backgroundColor As Color, foregroundColor As Color) = GetCurrentSystemThemeColors()
                    Dim bgColor As String = $"#{themeColors2.backgroundColor.R:X2}{themeColors2.backgroundColor.G:X2}{themeColors2.backgroundColor.B:X2}"
                    Dim fgColor As String = $"#{themeColors2.foregroundColor.R:X2}{themeColors2.foregroundColor.G:X2}{themeColors2.foregroundColor.B:X2}"
                    Dim accentColor As String = "#0078d4"

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
                        /* 自定义滚动条样式 */
                        ::-webkit-scrollbar {{
                            width: 8px;
                        }}
                        ::-webkit-scrollbar-track {{
                            background: {bgColor};
                        }}
                        ::-webkit-scrollbar-thumb {{
                            background: {accentColor};
                            border-radius: 4px;
                        }}
                        ::-webkit-scrollbar-thumb:hover {{
                            background: {accentColor};
                            opacity: 0.8;
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

                    ' 实时获取当前系统主题色
                    Dim themeColors3 As (backgroundColor As Color, foregroundColor As Color) = GetCurrentSystemThemeColors()
                    Dim bgColor As String = $"#{themeColors3.backgroundColor.R:X2}{themeColors3.backgroundColor.G:X2}{themeColors3.backgroundColor.B:X2}"
                    Dim fgColor As String = $"#{themeColors3.foregroundColor.R:X2}{themeColors3.foregroundColor.G:X2}{themeColors3.foregroundColor.B:X2}"
                    Dim accentColor As String = "#0078d4"

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
            ' 实时获取当前系统主题色作为默认值
            Dim defaultThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentSystemThemeColors()
            Dim defaultBgColor As String = $"#{defaultThemeColors.backgroundColor.R:X2}{defaultThemeColors.backgroundColor.G:X2}{defaultThemeColors.backgroundColor.B:X2}"
            Dim defaultFgColor As String = $"#{defaultThemeColors.foregroundColor.R:X2}{defaultThemeColors.foregroundColor.G:X2}{defaultThemeColors.foregroundColor.B:X2}"
            Dim defaultHtml As String = $"<html><head><style>body {{ background-color: {defaultBgColor} !important; color: {defaultFgColor} !important; font-family: Arial; padding: 10px; }}</style></head><body><p>无法显示内容</p></body></html>"
            Debug.WriteLine($"[DisplayMailContent] 返回默认HTML内容，长度: {defaultHtml.Length}")
            Return defaultHtml
        End Function

        Private Shared Function SanitizeHtmlBody(html As String) As String
            Try
                If String.IsNullOrEmpty(html) Then Return html
                Dim result As String = html
                result = Regex.Replace(result, "<\s*body[^>]*>", "<div>", RegexOptions.IgnoreCase)
                result = Regex.Replace(result, "<\s*/\s*body\s*>", "</div>", RegexOptions.IgnoreCase)
                result = Regex.Replace(result, "\sbgcolor\s*=\s*""?#?[A-Za-z0-9]+""?", "", RegexOptions.IgnoreCase)
                result = Regex.Replace(result, "background\s*:\s*[^;]+;?", "background: transparent !important;", RegexOptions.IgnoreCase)
                result = Regex.Replace(result, "background-color\s*:\s*[^;]+;?", "background-color: transparent !important;", RegexOptions.IgnoreCase)
                Return result
            Catch ex As System.Exception
                Debug.WriteLine($"SanitizeHtmlBody error: {ex.Message}")
                Return html
            End Try
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

        ' 实时获取当前系统主题色
        Private Shared Function GetCurrentSystemThemeColors() As (backgroundColor As Color, foregroundColor As Color)
            Try
                '直接从ThisAddIn获取实时主题颜色
                If Globals.ThisAddIn IsNot Nothing Then
                    Dim colors = Globals.ThisAddIn.GetCurrentThemeColors()
                    Debug.WriteLine($"[GetCurrentSystemThemeColors] 获取到的主题色: 背景={colors.backgroundColor}, 前景={colors.foregroundColor}")
                    If colors.backgroundColor = SystemColors.Window AndAlso colors.foregroundColor = SystemColors.WindowText Then
                        Debug.WriteLine("[警告] 检测到回退到系统默认颜色，可能是主题检测失败")
                    End If
                    Return colors
                Else
                    Debug.WriteLine("[GetCurrentSystemThemeColors] ThisAddIn为空，使用系统颜色")
                    Return (SystemColors.Window, SystemColors.WindowText)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[GetCurrentSystemThemeColors] 获取主题颜色失败: {ex.Message}")
                '返回系统默认颜色
                Return (SystemColors.Window, SystemColors.WindowText)
            End Try
        End Function
    End Class
End Namespace
