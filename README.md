# OA-MyList - Outlook 相关邮件插件

一个功能强大的 Microsoft Outlook 插件，用于显示和管理与当前选中邮件相关的邮件线程，提供高效的邮件会话浏览和任务管理功能。

## 🚀 主要功能

### 📧 相关邮件显示
- **智能邮件线程检测**：自动识别和显示与当前选中邮件相关的所有邮件
- **多视图支持**：支持 Explorer 窗口和 Inspector 窗口的邮件浏览
- **实时更新**：当选择不同邮件时，相关邮件列表自动更新
- **高性能显示**：采用虚拟化技术，支持大量邮件的高效显示

### 🎨 用户界面
- **自定义任务窗格**：在 Outlook 右侧显示"相关邮件v1.1"面板
- **主题适配**：自动适配 Outlook 的明暗主题
- **图标化显示**：使用直观的图标表示不同类型的邮件项目
  - 📧 普通邮件
  - 📅 日历/会议
  - 📋 任务
  - 👤 联系人
  - 📎 附件标识
  - 🚩 标记状态

### ⚡ 性能优化
- **MessageClass 映射缓存**：提高邮件类型判断效率
- **列表容量预分配**：减少动态扩容开销
- **ListView 虚拟化**：优化大量邮件会话的显示性能
- **分页功能**：支持分页浏览，提升大数据集的用户体验

### 📋 任务管理
- **任务监控**：集成任务监控功能
- **任务关联**：支持邮件与任务的关联管理

## 🛠️ 技术架构

### 核心组件

#### `ThisAddIn.vb`
- 插件主入口点
- 管理 Explorer 和 Inspector 窗口事件
- 处理主题变化和窗口状态

#### `MailThreadPane.vb`
- 主要用户界面组件
- 实现邮件列表显示和虚拟化
- 处理邮件选择和高亮显示

#### `OutlookRibbon.vb`
- Ribbon 界面集成
- 提供插件控制按钮
- 分页功能开关

#### `TaskMonitor.vb`
- 任务监控和管理
- 处理任务相关事件

### 工具类

#### `Utils/OutlookUtils.vb`
- Outlook 对象安全访问
- 性能优化的 GetItemFromID 方法
- 通用工具函数

#### `Utils/MailUtils.vb`
- 邮件内容处理工具

#### `Handlers/MailHandler.vb`
- 邮件处理逻辑
- EntryID 转换和管理
- 邮件高亮显示

#### `Models/TaskInfo.vb`
- 任务信息数据模型
- 支持邮件-任务关联

## 🔧 系统要求

- **操作系统**：Windows 10/11
- **Microsoft Outlook**：2016 或更高版本
- **.NET Framework**：4.7.2
- **Visual Studio Tools for Office Runtime**：必需

## 📦 安装和部署

### 开发环境
1. 安装 Visual Studio 2022 Community 或更高版本
2. 安装 Office Developer Tools
3. 克隆或下载项目源代码
4. 打开 `OutlookMyList.sln` 解决方案文件

### 构建项目
```bash
# 使用提供的批处理文件
build.bat

# 或使用 MSBuild 命令
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "OutlookMyList.sln" /p:Configuration="Debug" /p:Platform="Any CPU"
```

### 部署
1. 构建 Release 版本
2. 使用 ClickOnce 部署或手动安装
3. 确保目标机器安装了必要的运行时组件

## 🎯 使用方法

1. **启动插件**：安装后，Outlook 启动时会自动加载插件
2. **显示面板**：在 Outlook Ribbon 中点击插件按钮显示/隐藏相关邮件面板
3. **浏览相关邮件**：选择任意邮件，右侧面板会自动显示相关邮件列表
4. **分页浏览**：对于大量邮件，可启用分页功能提升性能
5. **主题适配**：插件会自动适配 Outlook 的当前主题

## 🔍 主要特性详解

### 虚拟化显示
- 采用 ListView 虚拟模式，支持数万封邮件的流畅显示
- 智能分页机制，根据数据量自动启用/禁用虚拟化
- 内存使用优化，避免大数据集导致的性能问题

### 缓存优化
- MessageClass 类型映射缓存，减少重复计算
- EntryID 比较缓存，提升邮件匹配效率
- 智能容量预分配，减少列表扩容开销

### 事件处理
- 防重复调用机制，避免频繁更新
- 异步处理，保持 UI 响应性
- 智能抑制机制，在列表构建时避免不必要的更新

## 🐛 故障排除

### 常见问题
1. **插件未加载**：检查 VSTO Runtime 是否正确安装
2. **性能问题**：启用分页功能或检查邮件数量
3. **显示异常**：重启 Outlook 或重新安装插件

### 调试信息
- 插件使用 Debug.WriteLine 输出调试信息
- 可通过 Visual Studio 输出窗口查看详细日志

## 📝 版本历史

### v1.1 (当前版本)
- 添加虚拟化显示支持
- 实现性能优化缓存
- 增强主题适配功能
- 修复编译警告
- 优化内存使用

## 🤝 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目。

## 📄 许可证

Copyright © 2025. 保留所有权利。

## 📞 支持

如有问题或建议，请通过以下方式联系：
- 创建 GitHub Issue
- 发送邮件至项目维护者

---

**注意**：此插件仅供内部使用，请确保在部署前进行充分测试。