1. 在VB.NET中应该使用System.Exception而不是Exception。
2. 属性属于计算属性（Computed Properties） 
    - TaskSubject、TaskCompletedDate、TaskDueDate等都是Outlook中的计算属性
    - 计算属性无法通过 Columns.Add 方法添加到Table对象中 2
    - 如果尝试添加这些属性，会收到 IDS_ERR_BLOCKED_PROPERTY 错误
    官方推荐的解决方案：要获取这些计算属性的值，必须：    1. 使用邮件的 EntryID 调用 GetItemFromID 获取完整的邮件对象；    2. 通过邮件对象直接访问这些属性值
3. 不要使用“dotnet build“， 用“.\build.bat“编译工程
4. 核心搜索目录是GetAllMailFolders， 不要修改， 请使用这个来访问核心目录
5. 搜索时，必须使用命令：    rg "pattern" "filename"

6. 代码修改优化策略（避免修改失败）：
   - 第一步：使用rg获取精确行号
     ```
     rg -n "pattern" "filename"
     ```
   - 第二步：确认上下文
     ```
     rg -A 5 -B 5 -n "pattern" "filename"
     ```
   - 第三步：使用行号范围精确修改（避免字符串匹配失败）
   - 批量验证：修改后立即执行
     ```
     .\build.bat
     rg -c "修改后的模式"
     ```

7. 主题应用一致性检查：
   - 所有调用LoadMailContentDeferred的地方必须检查是否立即应用主题
   - 标准模式：在调用LoadMailContentDeferred后立即添加
     ```vb
     Dim currentThemeColors As (backgroundColor As Color, foregroundColor As Color) = GetCurrentThemeColors()
     UpdateWebBrowserTheme(currentThemeColors.backgroundColor, currentThemeColors.foregroundColor)
     ```