# 开发文档

## 开发环境搭建

### 1. 安装必要工具

**Rust 环境:**
```bash
# 安装 Rust
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh

# 更新到最新版本
rustup update

# 验证安装
rustc --version
cargo --version
```

**Tauri CLI:**
```bash
cargo install tauri-cli --version "^2.0.0"
```

### 2. 平台特定依赖

**Windows:**
- 安装 Visual Studio Build Tools
- 安装 WebView2 Runtime

**macOS:**
- 安装 Xcode Command Line Tools
```bash
xcode-select --install
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install libwebkit2gtk-4.0-dev \
    build-essential \
    curl \
    wget \
    file \
    libssl-dev \
    libgtk-3-dev \
    libayatana-appindicator3-dev \
    librsvg2-dev
```

## 项目架构

### 目录结构详解

```
ArchiveBox/
├── frontend/                    # 前端资源
│   ├── index.html              # 主HTML文件
│   ├── main.js                 # 前端JavaScript逻辑
│   ├── style.css               # 样式文件
│   └── vendor/                 # 第三方库
├── src-tauri/                  # Tauri后端
│   ├── src/                    # Rust源代码
│   │   ├── main.rs            # 主入口
│   │   ├── commands.rs        # Tauri命令
│   │   ├── file_processor.rs  # 文件处理逻辑
│   │   └── utils.rs           # 工具函数
│   ├── Cargo.toml             # Rust依赖配置
│   ├── tauri.conf.json        # Tauri配置
│   └── build.rs               # 构建脚本
├── .github/workflows/          # CI/CD配置
├── docs/                       # 文档
└── tests/                      # 测试文件
```

### 核心模块

**1. 文件处理模块 (file_processor.rs)**
- ZIP文件解压和扫描
- Word文档解析和合并
- Excel数据提取和生成
- PDF文件处理

**2. 前后端通信 (commands.rs)**
- Tauri命令定义
- 前端调用接口
- 错误处理和返回

**3. 用户界面 (frontend/)**
- 文件选择和拖拽
- 进度显示
- 结果展示

## 开发流程

### 1. 启动开发服务器

```bash
# 开发模式，支持热重载
cargo tauri dev
```

### 2. 代码规范

**Rust代码:**
- 使用 `cargo fmt` 格式化代码
- 使用 `cargo clippy` 检查代码质量
- 遵循Rust官方编码规范

**JavaScript代码:**
- 使用ES6+语法
- 保持代码简洁和可读性
- 添加必要的注释

### 3. 测试

**单元测试:**
```bash
# 运行Rust测试
cargo test

# 运行特定测试
cargo test test_zip_processing
```

**集成测试:**
```bash
# 构建并测试完整应用
cargo tauri build --debug
```

## 构建和发布

### 1. 本地构建

**开发构建:**
```bash
cargo tauri build --debug
```

**生产构建:**
```bash
cargo tauri build
```

### 2. 跨平台构建

**使用GitHub Actions:**
- 推送代码到main分支
- 创建tag触发release构建
- 从Actions页面下载构建产物

**本地交叉编译:**
```bash
# 添加目标平台
rustup target add x86_64-pc-windows-msvc

# 构建Windows版本（在macOS/Linux上）
cargo tauri build --target x86_64-pc-windows-msvc
```

### 3. 发布流程

1. 更新版本号 (`src-tauri/Cargo.toml` 和 `src-tauri/tauri.conf.json`)
2. 更新CHANGELOG.md
3. 创建git tag
4. 推送到GitHub触发自动构建
5. 从GitHub Releases下载并测试
6. 发布release

## 调试技巧

### 1. 日志调试

**Rust后端:**
```rust
use log::{info, warn, error, debug};

fn process_file() {
    info!("开始处理文件");
    debug!("详细调试信息");
}
```

**前端:**
```javascript
console.log("前端调试信息");
console.error("错误信息");
```

### 2. 开发者工具

在开发模式下，可以使用浏览器开发者工具：
- 右键 -> 检查元素
- 或按F12打开开发者工具

### 3. 性能分析

**Rust性能分析:**
```bash
# 使用cargo flamegraph
cargo install flamegraph
cargo flamegraph --bin app
```

## 常见问题

### 1. 构建失败

**问题**: `failed to run custom build command`
**解决**: 检查系统依赖是否完整安装

**问题**: `linker 'cc' not found`
**解决**: 安装C编译器 (gcc/clang)

### 2. 运行时错误

**问题**: WebView加载失败
**解决**: 检查WebView2运行时是否安装

**问题**: 文件权限错误
**解决**: 检查应用是否有文件访问权限

### 3. 性能问题

**问题**: 大文件处理缓慢
**解决**: 
- 使用流式处理
- 添加进度反馈
- 考虑多线程处理

## 贡献指南

### 1. 代码提交

1. Fork项目
2. 创建功能分支
3. 编写代码和测试
4. 提交PR

### 2. 提交信息格式

```
type(scope): description

[optional body]

[optional footer]
```

**类型:**
- feat: 新功能
- fix: 修复bug
- docs: 文档更新
- style: 代码格式
- refactor: 重构
- test: 测试相关
- chore: 构建/工具相关

### 3. 代码审查

- 确保所有测试通过
- 代码符合项目规范
- 添加必要的文档和注释
- 性能和安全考虑

## 发布计划

### v0.1.0 (当前开发版本)
- [x] 基础ZIP处理功能
- [x] Word文档合并
- [x] Excel导出功能
- [x] 基础UI界面
- [ ] 完善错误处理
- [ ] 添加更多测试

### v0.2.0 (计划中)
- [ ] 支持更多文档格式
- [ ] 批量处理优化
- [ ] 用户配置保存
- [ ] 插件系统

### v1.0.0 (正式版本)
- [ ] 完整功能测试
- [ ] 性能优化
- [ ] 用户文档完善
- [ ] 多语言支持