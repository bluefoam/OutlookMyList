Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' 错误提醒设置窗体
''' </summary>
Public Class ErrorNotificationSettingsForm
    Inherits Form
    
    Private chkShowErrorDialogs As CheckBox
    Private chkShowOnlyFirstError As CheckBox
    Private chkLogErrorsToDebug As CheckBox
    Private chkShowCOMErrorDialogs As CheckBox
    Private btnOK As Button
    Private btnCancel As Button
    Private btnReset As Button
    
    Public Sub New()
        InitializeComponent()
        LoadCurrentSettings()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "错误提醒设置"
        Me.Size = New Size(400, 300)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        
        ' 创建控件
        Dim lblTitle As New Label With {
            .Text = "配置错误提醒选项：",
            .Location = New Point(20, 20),
            .Size = New Size(350, 20),
            .Font = New Font("Microsoft YaHei", 9, FontStyle.Bold)
        }
        
        chkShowErrorDialogs = New CheckBox With {
            .Text = "显示错误提醒对话框",
            .Location = New Point(30, 50),
            .Size = New Size(300, 20),
            .Checked = True
        }
        
        chkShowOnlyFirstError = New CheckBox With {
            .Text = "只显示第一次错误提醒（避免重复骚扰）",
            .Location = New Point(50, 80),
            .Size = New Size(300, 20),
            .Checked = True
        }
        
        chkLogErrorsToDebug = New CheckBox With {
            .Text = "记录错误到调试输出（用于开发调试）",
            .Location = New Point(30, 110),
            .Size = New Size(300, 20),
            .Checked = True
        }
        
        chkShowCOMErrorDialogs = New CheckBox With {
            .Text = "显示COM异常提醒（通常COM异常是临时性的）",
            .Location = New Point(30, 140),
            .Size = New Size(300, 20),
            .Checked = False
        }
        
        ' 说明标签
        Dim lblNote As New Label With {
            .Text = "注意：更改设置后需要重启Outlook才能完全生效。",
            .Location = New Point(20, 180),
            .Size = New Size(350, 20),
            .ForeColor = Color.Gray,
            .Font = New Font("Microsoft YaHei", 8)
        }
        
        ' 按钮
        btnOK = New Button With {
            .Text = "确定",
            .Location = New Point(150, 220),
            .Size = New Size(75, 30),
            .DialogResult = DialogResult.OK
        }
        
        btnCancel = New Button With {
            .Text = "取消",
            .Location = New Point(235, 220),
            .Size = New Size(75, 30),
            .DialogResult = DialogResult.Cancel
        }
        
        btnReset = New Button With {
            .Text = "重置默认",
            .Location = New Point(20, 220),
            .Size = New Size(80, 30)
        }
        
        ' 添加控件到窗体
        Me.Controls.AddRange({lblTitle, chkShowErrorDialogs, chkShowOnlyFirstError, 
                             chkLogErrorsToDebug, chkShowCOMErrorDialogs, lblNote,
                             btnOK, btnCancel, btnReset})
        
        ' 事件处理
        AddHandler chkShowErrorDialogs.CheckedChanged, AddressOf ChkShowErrorDialogs_CheckedChanged
        AddHandler btnOK.Click, AddressOf BtnOK_Click
        AddHandler btnReset.Click, AddressOf BtnReset_Click
        
        Me.AcceptButton = btnOK
        Me.CancelButton = btnCancel
    End Sub
    
    ''' <summary>
    ''' 加载当前设置
    ''' </summary>
    Private Sub LoadCurrentSettings()
        Dim settings = ErrorNotificationSettings.Instance
        chkShowErrorDialogs.Checked = settings.ShowErrorDialogs
        chkShowOnlyFirstError.Checked = settings.ShowOnlyFirstError
        chkLogErrorsToDebug.Checked = settings.LogErrorsToDebug
        chkShowCOMErrorDialogs.Checked = settings.ShowCOMErrorDialogs
        
        ' 更新控件状态
        UpdateControlStates()
    End Sub
    
    ''' <summary>
    ''' 更新控件状态
    ''' </summary>
    Private Sub UpdateControlStates()
        ' 如果不显示错误对话框，则禁用相关选项
        chkShowOnlyFirstError.Enabled = chkShowErrorDialogs.Checked
        chkShowCOMErrorDialogs.Enabled = chkShowErrorDialogs.Checked
    End Sub
    
    ''' <summary>
    ''' 显示错误对话框选项变化事件
    ''' </summary>
    Private Sub ChkShowErrorDialogs_CheckedChanged(sender As Object, e As EventArgs)
        UpdateControlStates()
    End Sub
    
    ''' <summary>
    ''' 确定按钮点击事件
    ''' </summary>
    Private Sub BtnOK_Click(sender As Object, e As EventArgs)
        Try
            ' 保存设置
            Dim settings = ErrorNotificationSettings.Instance
            settings.ShowErrorDialogs = chkShowErrorDialogs.Checked
            settings.ShowOnlyFirstError = chkShowOnlyFirstError.Checked
            settings.LogErrorsToDebug = chkLogErrorsToDebug.Checked
            settings.ShowCOMErrorDialogs = chkShowCOMErrorDialogs.Checked
            settings.SaveSettings()
            
            ' 重置全局错误标志，让新设置生效
            ThisAddIn.ResetErrorFlags()
            
            MessageBox.Show("设置已保存。部分设置可能需要重启Outlook后才能完全生效。", 
                          "设置保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"保存设置时出错：{ex.Message}", 
                          "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    
    ''' <summary>
    ''' 重置按钮点击事件
    ''' </summary>
    Private Sub BtnReset_Click(sender As Object, e As EventArgs)
        If MessageBox.Show("确定要重置为默认设置吗？", "确认重置", 
                          MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            ErrorNotificationSettings.Instance.ResetToDefaults()
            LoadCurrentSettings()
        End If
    End Sub
End Class