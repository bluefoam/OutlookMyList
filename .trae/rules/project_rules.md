1. 在VB.NET中应该使用System.Exception而不是Exception。
2. 属性属于计算属性（Computed Properties） 
    - TaskSubject、TaskCompletedDate、TaskDueDate等都是Outlook中的计算属性
    - 计算属性无法通过 Columns.Add 方法添加到Table对象中 2
    - 如果尝试添加这些属性，会收到 IDS_ERR_BLOCKED_PROPERTY 错误
    官方推荐的解决方案：要获取这些计算属性的值，必须：    1. 使用邮件的 EntryID 调用 GetItemFromID 获取完整的邮件对象；    2. 通过邮件对象直接访问这些属性值
3. 不要使用“dotnet build“