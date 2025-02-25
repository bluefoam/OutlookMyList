Imports System.Collections
Imports System.Windows.Forms

Public Class ListViewItemComparer
    Implements IComparer(Of ListViewItem)

    Private _columnIndex As Integer
    Private _sortOrder As SortOrder

    Public Sub New(columnIndex As Integer, sortOrder As SortOrder)
        _columnIndex = columnIndex
        _sortOrder = sortOrder
    End Sub

    Public Function Compare(x As ListViewItem, y As ListViewItem) As Integer Implements IComparer(Of ListViewItem).Compare
        Try
            Dim result As Integer

            ' 获取要比较的值
            Dim xValue As String = If(x.SubItems.Count > _columnIndex, x.SubItems(_columnIndex).Text, "")
            Dim yValue As String = If(y.SubItems.Count > _columnIndex, y.SubItems(_columnIndex).Text, "")

            ' 尝试作为日期比较
            Dim xDate, yDate As DateTime
            If DateTime.TryParse(xValue, xDate) AndAlso DateTime.TryParse(yValue, yDate) Then
                result = DateTime.Compare(xDate, yDate)
            Else
                ' 尝试作为数字比较
                Dim xNum, yNum As Double
                If Double.TryParse(xValue.TrimEnd("%"c), xNum) AndAlso Double.TryParse(yValue.TrimEnd("%"c), yNum) Then
                    result = xNum.CompareTo(yNum)
                Else
                    ' 默认作为字符串比较
                    result = String.Compare(xValue, yValue)
                End If
            End If

            ' 根据排序方向返回结果
            Return If(_sortOrder = SortOrder.Ascending, result, -result)
        Catch ex As Exception
            Return 0
        End Try
    End Function
End Class