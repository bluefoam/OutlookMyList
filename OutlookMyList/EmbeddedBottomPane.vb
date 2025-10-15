Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel

Public Class EmbeddedBottomPane
    Inherits UserControl

    Private titleLabel As Label
    Private contentPanel As Panel
    Private statusLabel As Label
    Private isEmbedded As Boolean = False
    Private outlookMainWindow As IntPtr = IntPtr.Zero
    Private readingPane As IntPtr = IntPtr.Zero
    Private originalParent As IntPtr = IntPtr.Zero

    Public Sub New()
        InitializeComponent()
        SetupControls()
    End Sub

    Private Sub InitializeComponent()
        Me.Size = New Size(600, 150)
        Me.BackColor = SystemColors.Control
        Me.BorderStyle = BorderStyle.FixedSingle
        Me.Dock = DockStyle.None
    End Sub

    Private Sub SetupControls()
        ' 创建标题栏
        titleLabel = New Label()
        titleLabel.Text = "插件面板 - 嵌入式"
        titleLabel.Font = New Font("Microsoft YaHei UI", 9, FontStyle.Bold)
        titleLabel.ForeColor = Color.DarkBlue
        titleLabel.BackColor = Color.LightGray
        titleLabel.TextAlign = ContentAlignment.MiddleLeft
        titleLabel.Padding = New Padding(10, 0, 0, 0)
        titleLabel.Dock = DockStyle.Top
        titleLabel.Height = 25
        Me.Controls.Add(titleLabel)

        ' 创建内容面板
        contentPanel = New Panel()
        contentPanel.BackColor = Color.White
        contentPanel.Dock = DockStyle.Fill
        contentPanel.Padding = New Padding(5)
        Me.Controls.Add(contentPanel)

        ' 创建状态标签
        statusLabel = New Label()
        statusLabel.Text = "状态：准备就绪"
        statusLabel.Font = New Font("Microsoft YaHei UI", 8)
        statusLabel.ForeColor = Color.Gray
        statusLabel.BackColor = Color.LightGray
        statusLabel.TextAlign = ContentAlignment.MiddleLeft
        statusLabel.Padding = New Padding(10, 0, 0, 0)
        statusLabel.Dock = DockStyle.Bottom
        statusLabel.Height = 20
        Me.Controls.Add(statusLabel)

        ' 添加一些示例内容
        AddSampleContent()
    End Sub

    Private Sub AddSampleContent()
        Dim infoLabel As New Label()
        infoLabel.Text = "这是嵌入到Outlook阅读窗格下方的自定义面板。" & vbCrLf & 
                        "面板位置：主邮件内容区域下方" & vbCrLf & 
                        "实现方式：Win32 API窗口嵌入"
        infoLabel.Font = New Font("Microsoft YaHei UI", 9)
        infoLabel.ForeColor = Color.Black
        infoLabel.AutoSize = False
        infoLabel.Size = New Size(580, 80)
        infoLabel.Location = New Point(10, 10)
        infoLabel.TextAlign = ContentAlignment.TopLeft
        contentPanel.Controls.Add(infoLabel)

        ' 添加一个按钮
        Dim testButton As New Button()
        testButton.Text = "测试功能"
        testButton.Size = New Size(100, 30)
        testButton.Location = New Point(10, 95)
        testButton.BackColor = SystemColors.Control
        AddHandler testButton.Click, AddressOf TestButton_Click
        contentPanel.Controls.Add(testButton)
    End Sub

    Private Sub TestButton_Click(sender As Object, e As EventArgs)
        statusLabel.Text = "状态：按钮已点击 - " & DateTime.Now.ToString("HH:mm:ss")
        MessageBox.Show("嵌入式底部面板功能测试成功！", "测试", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' 尝试嵌入到Outlook窗口
    Public Function TryEmbedInOutlook() As Boolean
        Try
            ' 查找Outlook主窗口
            outlookMainWindow = Win32Helper.FindOutlookMainWindow()
            If outlookMainWindow = IntPtr.Zero Then
                statusLabel.Text = "状态：未找到Outlook主窗口，使用备用方案"
                Return TryPositionNearOutlook()
            End If

            ' 查找阅读窗格
            readingPane = Win32Helper.FindReadingPane(outlookMainWindow)
            If readingPane = IntPtr.Zero Then
                statusLabel.Text = "状态：未找到阅读窗格，使用备用方案"
                Return TryPositionNearOutlook()
            End If

            ' 尝试嵌入面板
            If Win32Helper.CreateBottomPanel(outlookMainWindow, readingPane, Me, originalParent) Then
                isEmbedded = True
                statusLabel.Text = "状态：已成功嵌入到Outlook"
                Return True
            Else
                statusLabel.Text = "状态：嵌入失败，使用备用方案"
                Return TryPositionNearOutlook()
            End If

        Catch ex As Exception
            statusLabel.Text = "状态：嵌入出错，使用备用方案 - " & ex.Message
            System.Diagnostics.Debug.WriteLine("嵌入底部面板失败: " & ex.Message)
            Return TryPositionNearOutlook()
        End Try
    End Function

    ' 备用方案：将面板定位到Outlook窗口附近
    Private Function TryPositionNearOutlook() As Boolean
        Try
            If outlookMainWindow = IntPtr.Zero Then
                outlookMainWindow = Win32Helper.FindOutlookMainWindow()
            End If

            If outlookMainWindow <> IntPtr.Zero Then
                ' 获取Outlook主窗口位置
                Dim outlookRect As Win32Helper.RECT
                If Win32Helper.GetWindowRect(outlookMainWindow, outlookRect) Then
                    ' 创建独立窗口但定位在Outlook窗口下方
                    Dim form As New Form()
                    form.Text = "插件面板 - 独立窗口"
                    form.Size = New Size(600, 150)
                    form.StartPosition = FormStartPosition.Manual
                    form.Location = New Point(outlookRect.Left + 50, outlookRect.Bottom - 200)
                    form.TopMost = False
                    form.ShowInTaskbar = False
                    form.FormBorderStyle = FormBorderStyle.FixedToolWindow
                    
                    ' 将当前控件添加到窗口
                    Me.Dock = DockStyle.Fill
                    form.Controls.Add(Me)
                    
                    ' 显示窗口
                    form.Show()
                    
                    statusLabel.Text = "状态：已显示为独立窗口（位于Outlook附近）"
                    Return True
                End If
            End If
            
            ' 最后的备用方案：显示在屏幕中央
            Dim centerForm As New Form()
            centerForm.Text = "插件面板 - 独立窗口"
            centerForm.Size = New Size(600, 150)
            centerForm.StartPosition = FormStartPosition.CenterScreen
            centerForm.TopMost = False
            centerForm.ShowInTaskbar = False
            centerForm.FormBorderStyle = FormBorderStyle.FixedToolWindow
            
            Me.Dock = DockStyle.Fill
            centerForm.Controls.Add(Me)
            centerForm.Show()
            
            statusLabel.Text = "状态：已显示为独立窗口（屏幕中央）"
            Return True
            
        Catch ex As Exception
            statusLabel.Text = "状态：备用方案也失败 - " & ex.Message
            System.Diagnostics.Debug.WriteLine("备用定位方案失败: " & ex.Message)
            Return False
        End Try
    End Function

    ' 检查是否已嵌入
    Public ReadOnly Property IsEmbeddedInOutlook As Boolean
        Get
            Return isEmbedded
        End Get
    End Property

    ' 更新状态
    Public Sub UpdateStatus(status As String)
        If statusLabel IsNot Nothing Then
            statusLabel.Text = "状态：" & status
        End If
    End Sub

    ' 清理资源
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            ' 如果已嵌入，尝试恢复原始状态
            If isEmbedded Then
                Try
                    ' 恢复原始父窗口关系
                    If originalParent <> IntPtr.Zero AndAlso Me.Handle <> IntPtr.Zero Then
                        Win32Helper.SetParent(Me.Handle, originalParent)
                        System.Diagnostics.Debug.WriteLine("已恢复原始父窗口关系")
                    End If
                    
                    ' 恢复阅读窗格的原始大小
                    If readingPane <> IntPtr.Zero Then
                        Dim rect As Win32Helper.RECT
                        If Win32Helper.GetWindowRect(readingPane, rect) Then
                            Win32Helper.SetWindowPos(readingPane, Win32Helper.HWND_TOP,
                                                    rect.Left, rect.Top,
                                                    rect.Width, rect.Height + 150,
                                                    Win32Helper.SWP_NOZORDER Or Win32Helper.SWP_NOACTIVATE)
                            System.Diagnostics.Debug.WriteLine("已恢复阅读窗格大小")
                        End If
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("恢复原始状态失败: " & ex.Message)
                End Try
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' 应用主题
    Public Sub ApplyTheme(backgroundColor As Color, foregroundColor As Color)
        Try
            Me.BackColor = backgroundColor
            contentPanel.BackColor = backgroundColor
            
            ' 更新标题栏颜色
            titleLabel.BackColor = If(backgroundColor = SystemColors.Window, Color.LightGray, Color.FromArgb(backgroundColor.R - 20, backgroundColor.G - 20, backgroundColor.B - 20))
            titleLabel.ForeColor = foregroundColor
            
            ' 更新状态栏颜色
            statusLabel.BackColor = titleLabel.BackColor
            statusLabel.ForeColor = foregroundColor
            
            ' 应用主题到所有控件
            ApplyThemeToControls(contentPanel, backgroundColor, foregroundColor)
        Catch ex As Exception
            ' 忽略主题应用错误
        End Try
    End Sub

    Private Sub ApplyThemeToControls(parent As Control, backgroundColor As Color, foregroundColor As Color)
        For Each ctrl As Control In parent.Controls
            If TypeOf ctrl Is Button Then
                ' 为按钮应用主题颜色
                Dim btn As Button = DirectCast(ctrl, Button)
                btn.BackColor = backgroundColor
                btn.ForeColor = foregroundColor
                btn.FlatStyle = FlatStyle.Flat
                btn.FlatAppearance.BorderColor = foregroundColor
                btn.FlatAppearance.BorderSize = 1
            ElseIf TypeOf ctrl Is Label Then
                ctrl.ForeColor = foregroundColor
            Else
                ctrl.BackColor = backgroundColor
                ctrl.ForeColor = foregroundColor
            End If
            
            ' 递归应用到子控件
            If ctrl.HasChildren Then
                ApplyThemeToControls(ctrl, backgroundColor, foregroundColor)
            End If
        Next
    End Sub
End Class