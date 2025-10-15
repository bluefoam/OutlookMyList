Imports System.IO
Imports System.Xml.Serialization
Imports System.Environment

''' <summary>
''' 错误提醒配置设置类
''' </summary>
Public Class ErrorNotificationSettings
    ''' <summary>
    ''' 是否显示错误提醒对话框
    ''' </summary>
    Public Property ShowErrorDialogs As Boolean = False
    
    ''' <summary>
    ''' 是否只显示第一次错误提醒
    ''' </summary>
    Public Property ShowOnlyFirstError As Boolean = True
    
    ''' <summary>
    ''' 是否记录错误到调试输出
    ''' </summary>
    Public Property LogErrorsToDebug As Boolean = True
    
    ''' <summary>
    ''' COM异常是否显示提醒（通常COM异常是临时性的）
    ''' </summary>
    Public Property ShowCOMErrorDialogs As Boolean = False
    
    Private Shared _instance As ErrorNotificationSettings
    Private Shared ReadOnly _lock As New Object()
    
    ''' <summary>
    ''' 单例实例
    ''' </summary>
    Public Shared ReadOnly Property Instance As ErrorNotificationSettings
        Get
            If _instance Is Nothing Then
                SyncLock _lock
                    If _instance Is Nothing Then
                        _instance = LoadSettings()
                    End If
                End SyncLock
            End If
            Return _instance
        End Get
    End Property
    
    ''' <summary>
    ''' 配置文件路径
    ''' </summary>
    Private Shared ReadOnly Property SettingsFilePath As String
        Get
            Dim appDataPath = GetFolderPath(SpecialFolder.ApplicationData)
            Dim folderPath = Path.Combine(appDataPath, "OutlookMyList")
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
            End If
            Return Path.Combine(folderPath, "ErrorNotificationSettings.xml")
        End Get
    End Property
    
    ''' <summary>
    ''' 加载设置
    ''' </summary>
    Private Shared Function LoadSettings() As ErrorNotificationSettings
        Try
            If File.Exists(SettingsFilePath) Then
                Dim serializer As New XmlSerializer(GetType(ErrorNotificationSettings))
                Using reader As New StreamReader(SettingsFilePath)
                    Return DirectCast(serializer.Deserialize(reader), ErrorNotificationSettings)
                End Using
            End If
        Catch ex As Exception
            ' 如果加载失败，使用默认设置
            System.Diagnostics.Debug.WriteLine($"加载错误提醒设置失败: {ex.Message}")
        End Try
        
        ' 返回默认设置
        Return New ErrorNotificationSettings()
    End Function
    
    ''' <summary>
    ''' 保存设置
    ''' </summary>
    Public Sub SaveSettings()
        Try
            Dim serializer As New XmlSerializer(GetType(ErrorNotificationSettings))
            Using writer As New StreamWriter(SettingsFilePath)
                serializer.Serialize(writer, Me)
            End Using
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"保存错误提醒设置失败: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' 重置为默认设置
    ''' </summary>
    Public Sub ResetToDefaults()
        ShowErrorDialogs = False
        ShowOnlyFirstError = True
        LogErrorsToDebug = True
        ShowCOMErrorDialogs = False
        SaveSettings()
    End Sub
End Class