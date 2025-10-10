Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Diagnostics

''' <summary>
''' 鼠标显示问题修复类
''' 用于解决Outlook插件中鼠标光标不显示的问题
''' </summary>
Public Class MouseFix
    ' Win32 API 声明
    <DllImport("user32.dll")>
    Private Shared Function ShowCursor(bShow As Boolean) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetCursorInfo(ByRef pci As CURSORINFO) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetCursor(hCursor As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function LoadCursor(hInstance As IntPtr, lpCursorName As IntPtr) As IntPtr
    End Function

    ' 常量
    Private Const IDC_ARROW As Integer = 32512

    ' 结构体
    <StructLayout(LayoutKind.Sequential)>
    Private Structure CURSORINFO
        Public cbSize As Integer
        Public flags As Integer
        Public hCursor As IntPtr
        Public ptScreenPos As POINT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINT
        Public x As Integer
        Public y As Integer
    End Structure

    ''' <summary>
    ''' 检查鼠标光标是否可见
    ''' </summary>
    ''' <returns>True如果鼠标可见，False如果不可见</returns>
    Public Shared Function IsCursorVisible() As Boolean
        Try
            Dim cursorInfo As New CURSORINFO()
            cursorInfo.cbSize = Marshal.SizeOf(cursorInfo)
            
            If GetCursorInfo(cursorInfo) Then
                ' flags = 1 表示光标可见，0 表示隐藏
                Return (cursorInfo.flags And 1) = 1
            End If
        Catch ex As Exception
            Debug.WriteLine($"检查鼠标可见性时出错: {ex.Message}")
        End Try
        Return True ' 默认假设可见
    End Function

    ''' <summary>
    ''' 强制显示鼠标光标
    ''' </summary>
    Public Shared Sub ForceShowCursor()
        Try
            ' 确保鼠标光标可见
            Dim showCount As Integer = ShowCursor(True)
            Debug.WriteLine($"ShowCursor调用结果: {showCount}")
            
            ' 如果显示计数小于0，继续调用直到>=0
            While showCount < 0
                showCount = ShowCursor(True)
                Debug.WriteLine($"继续ShowCursor调用，结果: {showCount}")
            End While
            
            ' 设置默认箭头光标
            Dim arrowCursor As IntPtr = LoadCursor(IntPtr.Zero, New IntPtr(IDC_ARROW))
            If arrowCursor <> IntPtr.Zero Then
                SetCursor(arrowCursor)
                Debug.WriteLine("已设置默认箭头光标")
            End If
            
        Catch ex As Exception
            Debug.WriteLine($"强制显示鼠标光标时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 重置鼠标光标到默认状态
    ''' </summary>
    Public Shared Sub ResetCursor()
        Try
            ' 重置为默认光标
            Cursor.Current = Cursors.Default
            
            ' 强制刷新光标
            ForceShowCursor()
            
            Debug.WriteLine("鼠标光标已重置到默认状态")
        Catch ex As Exception
            Debug.WriteLine($"重置鼠标光标时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 修复控件的鼠标光标问题
    ''' </summary>
    ''' <param name="control">要修复的控件</param>
    Public Shared Sub FixControlCursor(control As Control)
        Try
            If control Is Nothing Then Return
            
            ' 确保控件使用默认光标
            control.Cursor = Cursors.Default
            
            ' 递归修复子控件
            For Each childControl As Control In control.Controls
                FixControlCursor(childControl)
            Next
            
        Catch ex As Exception
            Debug.WriteLine($"修复控件鼠标光标时出错: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' 全面的鼠标修复方法
    ''' </summary>
    Public Shared Sub ComprehensiveMouseFix()
        Try
            Debug.WriteLine("开始全面鼠标修复...")
            
            ' 1. 检查当前鼠标状态
            Dim isVisible As Boolean = IsCursorVisible()
            Debug.WriteLine($"当前鼠标可见性: {isVisible}")
            
            ' 2. 强制显示鼠标
            ForceShowCursor()
            
            ' 3. 重置光标
            ResetCursor()
            
            ' 4. 刷新应用程序
            Application.DoEvents()
            
            Debug.WriteLine("全面鼠标修复完成")
            
        Catch ex As Exception
            Debug.WriteLine($"全面鼠标修复时出错: {ex.Message}")
        End Try
    End Sub
End Class