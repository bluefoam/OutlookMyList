Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Public Class BottomPane
    Inherits UserControl

    Private titleBar As Panel
    Private titleLabel As Label
    Private minimizeButton As Button
    Private contentPanel As Panel
    Private _isMinimized As Boolean = False
    Private normalHeight As Integer = 200
    Private minimizedHeight As Integer = 30

    Public Sub New()
        InitializeComponent()
        SetupControls()
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        
        ' 设置用户控件基本属性
        Me.Name = "BottomPane"
        Me.Size = New Size(800, normalHeight)
        Me.BackColor = SystemColors.Control
        Me.BorderStyle = BorderStyle.FixedSingle
        
        Me.ResumeLayout(False)
    End Sub

    Private Sub SetupControls()
        ' 创建标题栏
        titleBar = New Panel()
        titleBar.Dock = DockStyle.Top
        titleBar.Height = 25
        titleBar.BackColor = SystemColors.ActiveCaption
        titleBar.BorderStyle = BorderStyle.None
        
        ' 创建标题标签
        titleLabel = New Label()
        titleLabel.Text = "插件面板"
        titleLabel.ForeColor = SystemColors.ActiveCaptionText
        titleLabel.Font = New Font("Microsoft YaHei", 9, FontStyle.Bold)
        titleLabel.AutoSize = False
        titleLabel.Size = New Size(200, 25)
        titleLabel.Location = New Point(5, 0)
        titleLabel.TextAlign = ContentAlignment.MiddleLeft
        
        ' 创建最小化按钮
        minimizeButton = New Button()
        minimizeButton.Text = "−"
        minimizeButton.Size = New Size(25, 23)
        minimizeButton.Location = New Point(titleBar.Width - 30, 1)
        minimizeButton.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        minimizeButton.FlatStyle = FlatStyle.Flat
        minimizeButton.FlatAppearance.BorderSize = 0
        minimizeButton.BackColor = SystemColors.ActiveCaption
        minimizeButton.ForeColor = SystemColors.ActiveCaptionText
        minimizeButton.Font = New Font("Microsoft YaHei", 12, FontStyle.Bold)
        AddHandler minimizeButton.Click, AddressOf MinimizeButton_Click
        
        ' 创建内容面板
        contentPanel = New Panel()
        contentPanel.Dock = DockStyle.Fill
        contentPanel.BackColor = SystemColors.Window
        contentPanel.Padding = New Padding(5)
        
        ' 添加示例内容
        Dim sampleLabel As New Label()
        sampleLabel.Text = "这是插件面板的内容区域。\n您可以在这里添加各种功能控件。"
        sampleLabel.AutoSize = True
        sampleLabel.Location = New Point(10, 10)
        sampleLabel.Font = New Font("Microsoft YaHei", 9)
        contentPanel.Controls.Add(sampleLabel)
        
        ' 将控件添加到用户控件
        titleBar.Controls.Add(titleLabel)
        titleBar.Controls.Add(minimizeButton)
        Me.Controls.Add(contentPanel)
        Me.Controls.Add(titleBar)
    End Sub

    Private Sub MinimizeButton_Click(sender As Object, e As EventArgs)
        ToggleMinimize()
    End Sub

    Public Sub ToggleMinimize()
        _isMinimized = Not _isMinimized
        
        If _isMinimized Then
            ' 最小化
            Me.Height = minimizedHeight
            contentPanel.Visible = False
            minimizeButton.Text = "□"
            titleLabel.Text = "插件面板 (已最小化)"
        Else
            ' 还原
            Me.Height = normalHeight
            contentPanel.Visible = True
            minimizeButton.Text = "−"
            titleLabel.Text = "插件面板"
        End If
        
        ' 触发大小改变事件
        OnSizeChanged(EventArgs.Empty)
    End Sub

    Public ReadOnly Property IsMinimized As Boolean
        Get
            Return _isMinimized
        End Get
    End Property

    ' 应用主题
    Public Sub ApplyTheme(backgroundColor As Color, foregroundColor As Color)
        Try
            Me.BackColor = backgroundColor
            contentPanel.BackColor = backgroundColor
            
            ' 更新标题栏颜色
            titleBar.BackColor = If(backgroundColor = SystemColors.Window, SystemColors.ActiveCaption, Color.FromArgb(backgroundColor.R - 20, backgroundColor.G - 20, backgroundColor.B - 20))
            titleLabel.ForeColor = foregroundColor
            minimizeButton.BackColor = titleBar.BackColor
            minimizeButton.ForeColor = foregroundColor
        Catch ex As Exception
            ' 忽略主题应用错误
        End Try
    End Sub
End Class