Namespace OutlookMyList.Models
    Public Class TaskInfo
        Public Property TaskEntryID As String
        Public Property MailEntryID As String
        Public Property Subject As String
        Public Property DueDate As DateTime?
        Public Property Status As String
        Public Property PercentComplete As Integer
        Public Property LinkedMailSubject As String
        Public Property RelatedMailSubject As String
        ''' <summary>
        ''' StoreID，用于提升 GetItemFromID 性能
        ''' </summary>
        Public Property StoreID As String
    End Class
End Namespace