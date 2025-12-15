# ArchiveBox - 离线ZIP资料聚合与导出桌面程序

一个基于Tauri开发的跨平台桌面应用程序，用于聚合和导出ZIP文件中的资料，支持Word文档合并和Excel数据导出。

## 功能特性

### 📁 ZIP文件处理
- 批量选择和处理ZIP文件
- 自动解压和文件扫描
- 支持嵌套ZIP文件处理
- 智能文件类型识别

### 📄 Word文档处理
- 自动提取ZIP中的Word文档(.docx)
- 智能合并多个Word文档
- 保持原始格式和样式
- 支持图片和表格合并
- 生成统一的合并文档

### 📊 Excel数据导出
- 提取文档中的结构化数据
- 自动生成Excel报表
- 支持多种数据格式
- 可自定义导出字段

### 🖥️ 用户界面
- 现代化的桌面界面
- 拖拽式文件操作
- 实时处理进度显示
- 跨平台兼容(Windows/macOS/Linux)

## 技术架构

- **前端**: HTML + CSS + JavaScript
- **后端**: Rust + Tauri
- **文档处理**: docx-rs, rust_xlsxwriter
- **压缩处理**: zip crate
- **PDF处理**: lopdf

## 安装使用

### 下载安装包

从 [Releases](../../releases) 页面下载对应平台的安装包：

- **Windows**: `ArchiveBox_x.x.x_x64_setup.exe`
- **macOS**: `ArchiveBox_x.x.x_aarch64.dmg` 或 `ArchiveBox_x.x.x_x64.dmg`
- **Linux**: `ArchiveBox_x.x.x_amd64.deb` 或 `ArchiveBox_x.x.x_amd64.AppImage`

### 使用方法

1. **启动应用程序**
   - 双击安装后的应用图标

2. **选择ZIP文件**
   - 点击"选择文件"按钮
   - 或直接拖拽ZIP文件到应用窗口

3. **处理和导出**
   - 应用会自动扫描ZIP文件内容
   - 选择需要的处理选项
   - 点击"开始处理"
   - 选择导出位置和格式

4. **查看结果**
   - 处理完成后会显示结果统计
   - 可以直接打开导出的文件

## 开发环境

### 环境要求

- **Rust**: 1.92.0 或更高版本
- **Node.js**: 18.x 或更高版本
- **Tauri CLI**: 2.x

### 本地开发

1. **克隆项目**
   ```bash
   git clone <repository-url>
   cd ArchiveBox
   ```

2. **安装依赖**
   ```bash
   # 安装Tauri CLI
   cargo install tauri-cli --version "^2.0.0"
   ```

3. **开发模式运行**
   ```bash
   cargo tauri dev
   ```

4. **构建生产版本**
   ```bash
   cargo tauri build
   ```

### 项目结构

```
ArchiveBox/
├── frontend/                 # 前端代码
│   ├── index.html           # 主页面
│   ├── main.js              # 主要逻辑
│   └── style.css            # 样式文件
├── src-tauri/               # Tauri后端
│   ├── src/                 # Rust源代码
│   ├── Cargo.toml           # Rust依赖配置
│   └── tauri.conf.json      # Tauri配置
├── .github/workflows/       # CI/CD配置
└── README.md               # 项目说明
```

## 构建部署

### 自动构建

项目配置了GitHub Actions自动构建，支持多平台：

- 推送代码到main分支会触发构建
- 创建tag会自动发布release
- 支持手动触发构建

### 手动构建

**macOS/Linux:**
```bash
cargo tauri build
```

**Windows:**
```bash
cargo tauri build
```

构建产物位于 `src-tauri/target/release/bundle/` 目录下。

## 依赖库

### Rust依赖
- `tauri`: 桌面应用框架
- `docx-rs`: Word文档处理
- `rust_xlsxwriter`: Excel文件生成
- `zip`: ZIP文件处理
- `lopdf`: PDF文件处理
- `serde`: 序列化/反序列化
- `anyhow`: 错误处理

### 系统要求

**Windows:**
- Windows 10 或更高版本
- WebView2 运行时

**macOS:**
- macOS 10.15 或更高版本

**Linux:**
- GTK 3.0+
- WebKitGTK 4.0+

## 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建 Pull Request

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 更新日志

### v0.1.0 (开发中)
- ✅ 基础ZIP文件处理功能
- ✅ Word文档合并功能
- ✅ Excel数据导出功能
- ✅ 跨平台桌面界面
- ✅ 自动构建和发布流程

## 问题反馈

如果遇到问题或有功能建议，请：

1. 查看 [Issues](../../issues) 页面
2. 创建新的Issue描述问题
3. 提供详细的错误信息和复现步骤

## 联系方式

- 项目主页: [GitHub Repository](../../)
- 问题报告: [Issues](../../issues)
- 功能请求: [Discussions](../../discussions)

---

**注意**: 本项目仍在开发中，功能可能会有变化。建议在生产环境使用前进行充分测试。