Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook

''' <summary>
''' 直接合并邮件到同一会话的辅助类
''' </summary>
Namespace OutlookMyList
    Public Class DirectMergeHelper
    Private Const CustomConversationPropertyName As String = "CustomConversationId"

    ''' <summary>
    ''' 直接合并邮件到同一会话
    ''' </summary>
    ''' <param name="selection">选中的邮件集合</param>
    ''' <returns>合并结果信息</returns>
    Public Shared Function MergeConversation(selection As Selection) As (success As Boolean, processedCount As Integer, errorCount As Integer, targetConversationId As String)
        Dim processedCount As Integer = 0
        Dim errorCount As Integer = 0
        Dim targetConversationId As String = String.Empty

        Try
            If selection Is Nothing OrElse selection.Count < 2 Then
                Return (False, 0, 0, String.Empty)
            End If

            ' 收集所有需要处理的邮件（包括选中邮件所属会话组中的所有邮件）
            Dim allMailsToProcess As New List(Of Object)
            Dim processedConversationIds As New HashSet(Of String)

            ' 首先检查所有选中的邮件，查找是否有任何一个已存在自定义会话ID
            For i As Integer = 1 To selection.Count
                Try
                    Dim mailItem As Object = selection(i)
                    Dim customId As String = ReadCustomConversationIdFromItem(mailItem)
                    If Not String.IsNullOrEmpty(customId) Then
                        targetConversationId = customId
                        Debug.WriteLine($"找到自定义会话ID: {targetConversationId}")
                        Exit For
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"检查邮件 {i} 的自定义会话ID时出错: {ex.Message}")
                End Try
            Next

            ' 如果没有找到任何自定义会话ID，则使用第一个邮件的原始ConversationID
            If String.IsNullOrEmpty(targetConversationId) Then
                Dim firstMailItem As Object = selection(1)
                If TypeOf firstMailItem Is MailItem Then
                    targetConversationId = DirectCast(firstMailItem, MailItem).ConversationID
                    Debug.WriteLine($"使用第一封邮件的原始ConversationID: {targetConversationId}")
                ElseIf TypeOf firstMailItem Is AppointmentItem Then
                    targetConversationId = DirectCast(firstMailItem, AppointmentItem).ConversationID
                    Debug.WriteLine($"使用第一个约会的原始ConversationID: {targetConversationId}")
                ElseIf TypeOf firstMailItem Is MeetingItem Then
                    targetConversationId = DirectCast(firstMailItem, MeetingItem).ConversationID
                    Debug.WriteLine($"使用第一个会议的原始ConversationID: {targetConversationId}")
                End If
            End If

            ' 如果仍然没有有效的会话ID，则返回失败
            If String.IsNullOrEmpty(targetConversationId) Then
                Debug.WriteLine("无法确定目标会话ID")
                Return (False, 0, 0, String.Empty)
            End If

            Debug.WriteLine($"开始合并操作，目标会话ID: {targetConversationId}")

            ' 收集选中邮件所属的所有会话组中的邮件
            For i As Integer = 1 To selection.Count
                Try
                    Dim mailItem As Object = selection(i)
                    Dim conversationId As String = String.Empty

                    ' 获取邮件的会话ID
                    If TypeOf mailItem Is MailItem Then
                        conversationId = DirectCast(mailItem, MailItem).ConversationID
                    ElseIf TypeOf mailItem Is AppointmentItem Then
                        conversationId = DirectCast(mailItem, AppointmentItem).ConversationID
                    ElseIf TypeOf mailItem Is MeetingItem Then
                        conversationId = DirectCast(mailItem, MeetingItem).ConversationID
                    End If

                    ' 如果这个会话ID还没有处理过，则收集该会话中的所有邮件
                    If Not String.IsNullOrEmpty(conversationId) AndAlso Not processedConversationIds.Contains(conversationId) Then
                        processedConversationIds.Add(conversationId)
                        Dim conversationMails = GetAllMailsInConversation(conversationId)
                        allMailsToProcess.AddRange(conversationMails)
                        Debug.WriteLine($"收集会话 {conversationId} 中的 {conversationMails.Count} 封邮件")
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"收集邮件 {i} 的会话邮件时出错: {ex.Message}")
                End Try
            Next

            Debug.WriteLine($"总共需要处理 {allMailsToProcess.Count} 封邮件")

            ' 遍历所有收集到的邮件，设置自定义会话ID
            For i As Integer = 0 To allMailsToProcess.Count - 1
                Try
                    Dim mailItem As Object = allMailsToProcess(i)
                    Debug.WriteLine($"处理邮件 {i + 1}/{allMailsToProcess.Count}，类型: {mailItem.GetType().Name}")

                    ' 直接设置自定义属性
                    Dim success As Boolean = SetCustomConversationIdToItem(mailItem, targetConversationId)
                    If success Then
                        Debug.WriteLine($"成功设置邮件 {i + 1} 的自定义会话ID")
                        processedCount += 1
                    Else
                        Debug.WriteLine($"设置邮件 {i + 1} 的自定义会话ID失败")
                        errorCount += 1
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"处理邮件 {i + 1} 时出错: {ex.Message}")
                    errorCount += 1
                End Try
            Next

            Return (True, processedCount, errorCount, targetConversationId)
        Catch ex As System.Exception
            Debug.WriteLine($"MergeConversation错误: {ex.Message}")
            Return (False, processedCount, errorCount, targetConversationId)
        End Try
    End Function

    ''' <summary>
    ''' 从邮件项中读取自定义会话ID
    ''' </summary>
    Private Shared Function ReadCustomConversationIdFromItem(mailItem As Object) As String
        Try
            If mailItem Is Nothing Then Return String.Empty

            Dim userProps As UserProperties = Nothing
            If TypeOf mailItem Is MailItem Then
                userProps = DirectCast(mailItem, MailItem).UserProperties
            ElseIf TypeOf mailItem Is AppointmentItem Then
                userProps = DirectCast(mailItem, AppointmentItem).UserProperties
            ElseIf TypeOf mailItem Is MeetingItem Then
                userProps = DirectCast(mailItem, MeetingItem).UserProperties
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

    ''' <summary>
    ''' 获取指定会话ID中的所有邮件
    ''' </summary>
    ''' <param name="conversationId">会话ID</param>
    ''' <returns>会话中的所有邮件列表</returns>
    Private Shared Function GetAllMailsInConversation(conversationId As String) As List(Of Object)
        Dim mailsInConversation As New List(Of Object)
        
        Try
            If String.IsNullOrEmpty(conversationId) Then
                Return mailsInConversation
            End If

            ' 获取Outlook应用程序
            Dim outlookApp = Globals.ThisAddIn.Application
            If outlookApp Is Nothing Then
                Debug.WriteLine("无法获取Outlook应用程序")
                Return mailsInConversation
            End If

            ' 获取所有邮件存储
            For Each store As Store In outlookApp.Session.Stores
                Try
                    ' 获取根文件夹
                    Dim rootFolder As Folder = store.GetRootFolder()
                    If rootFolder IsNot Nothing Then
                        ' 获取所有核心邮件文件夹
                        Dim allMailFolders As New List(Of Folder)
                        GetAllMailFolders(rootFolder, allMailFolders)
                        
                        ' 在每个文件夹中搜索指定会话ID的邮件
                        For Each folder As Folder In allMailFolders
                            Try
                                ' 使用过滤器查找会话中的邮件
                                Dim filter As String = $"[ConversationID] = '{conversationId}'"
                                Dim items = folder.Items.Restrict(filter)
                                
                                For Each item As Object In items
                                    If TypeOf item Is MailItem OrElse TypeOf item Is AppointmentItem OrElse TypeOf item Is MeetingItem Then
                                        mailsInConversation.Add(item)
                                        Debug.WriteLine($"在文件夹 {folder.Name} 中找到会话邮件: {item.GetType().Name}")
                                    End If
                                Next
                            Catch ex As System.Exception
                                Debug.WriteLine($"搜索文件夹 {folder.Name} 时出错: {ex.Message}")
                            End Try
                        Next
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"处理存储 {store.DisplayName} 时出错: {ex.Message}")
                End Try
            Next

            Debug.WriteLine($"在会话 {conversationId} 中找到 {mailsInConversation.Count} 封邮件")
        Catch ex As System.Exception
            Debug.WriteLine($"GetAllMailsInConversation错误: {ex.Message}")
        End Try

        Return mailsInConversation
    End Function

    ''' <summary>
    ''' 获取所有核心邮件文件夹（复制自MailThreadPane的GetAllMailFolders方法）
    ''' </summary>
    Private Shared Sub GetAllMailFolders(folder As Folder, folderList As List(Of Folder))
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
            If folder.DefaultItemType = OlItemType.olMailItem AndAlso coreFolders.Contains(folder.Name) Then
                folderList.Add(folder)
                Debug.WriteLine($"添加邮件文件夹: {folder.Name}")
            End If

            ' 只在核心文件夹中递归搜索
            If folder.Folders IsNot Nothing Then
                For Each subFolder As Folder In folder.Folders
                    If coreFolders.Contains(subFolder.Name) Then
                        GetAllMailFolders(subFolder, folderList)
                    End If
                Next
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"处理文件夹 {folder.Name} 时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 设置邮件项的自定义会话ID
    ''' </summary>
    Private Shared Function SetCustomConversationIdToItem(mailItem As Object, conversationId As String) As Boolean
        Try
            If mailItem Is Nothing Then
                Return False
            End If
            
            ' 注意：允许conversationId为空字符串，这表示要清除自定义会话ID
            If conversationId Is Nothing Then
                Return False
            End If

            Dim userProps As UserProperties = Nothing
            If TypeOf mailItem Is MailItem Then
                userProps = DirectCast(mailItem, MailItem).UserProperties
            ElseIf TypeOf mailItem Is AppointmentItem Then
                userProps = DirectCast(mailItem, AppointmentItem).UserProperties
            ElseIf TypeOf mailItem Is MeetingItem Then
                userProps = DirectCast(mailItem, MeetingItem).UserProperties
            End If

            If userProps Is Nothing Then
                Debug.WriteLine("无法获取UserProperties")
                Return False
            End If

            ' 如果conversationId为空字符串，表示要清除自定义会话ID
            If String.IsNullOrEmpty(conversationId) Then
                Try
                    Dim existingProp = userProps.Find(CustomConversationPropertyName)
                    If existingProp IsNot Nothing Then
                        existingProp.Delete()
                        Debug.WriteLine("自定义会话ID属性已删除")
                        
                        ' 保存邮件项
                        If TypeOf mailItem Is MailItem Then
                            DirectCast(mailItem, MailItem).Save()
                        ElseIf TypeOf mailItem Is AppointmentItem Then
                            DirectCast(mailItem, AppointmentItem).Save()
                        ElseIf TypeOf mailItem Is MeetingItem Then
                            DirectCast(mailItem, MeetingItem).Save()
                        End If
                        
                        Return True
                    Else
                        Debug.WriteLine("自定义会话ID属性不存在，无需删除")
                        Return True
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"删除自定义会话ID属性时出错: {ex.Message}")
                    Return False
                End Try
            Else
                ' 设置或更新自定义会话ID
                ' 删除现有属性（如果存在）
                Try
                    Dim existingProp = userProps.Find(CustomConversationPropertyName)
                    If existingProp IsNot Nothing Then
                        existingProp.Delete()
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"删除现有属性时出错: {ex.Message}")
                    ' 继续尝试添加新属性
                End Try

                ' 添加新属性
                Try
                    Dim prop = userProps.Add(CustomConversationPropertyName, OlUserPropertyType.olText)
                    prop.Value = conversationId

                    ' 保存邮件项
                    If TypeOf mailItem Is MailItem Then
                        DirectCast(mailItem, MailItem).Save()
                    ElseIf TypeOf mailItem Is AppointmentItem Then
                        DirectCast(mailItem, AppointmentItem).Save()
                    ElseIf TypeOf mailItem Is MeetingItem Then
                        DirectCast(mailItem, MeetingItem).Save()
                    End If

                    ' 验证保存是否成功
                    Dim verifyProp = userProps.Find(CustomConversationPropertyName)
                    If verifyProp IsNot Nothing AndAlso verifyProp.Value.ToString() = conversationId Then
                        Debug.WriteLine("属性设置成功并已验证")
                        Return True
                    Else
                        Debug.WriteLine("属性设置后验证失败")
                        Return False
                    End If
                Catch ex As System.Exception
                    Debug.WriteLine($"设置属性时出错: {ex.Message}")
                    Return False
                End Try
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"SetCustomConversationIdToItem error: {ex.Message}")
            Return False
        End Try
    End Function
End Class
End Namespace