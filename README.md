# SeaTable Excel 同步工具

SeaTable Excel 同步工具是一个智能化的数据同步解决方案，专门设计用于将 SeaTable 中的数据同步到 Excel 文件。该工具支持多配置文件管理、智能菜单系统和灵活的字段映射，为企业数据管理提供便捷高效的解决方案。

## 🚀 核心特性

### 📁 多配置文件管理
- **自动发现**: 程序自动扫描目录下所有 JSON 配置文件
- **独立配置**: 每个配置文件支持独立的 SeaTable API Token
- **智能菜单**: 配置文件支持中文描述，菜单显示更直观

### 🔧 灵活的配置系统
- **环境变量配置**: 敏感信息通过 .env 文件管理
- **字段映射**: 支持 SeaTable 字段到 Excel 列的灵活映射
- **关联匹配**: 基于关联字段智能匹配和更新数据

### 💾 智能数据处理
- **增量更新**: 基于关联字段匹配，只更新变化的数据
- **自动备份**: 同步后自动生成带日期的备份文件
- **数据验证**: 同步前后数据完整性验证

### 🖥️ 用户友好界面
- **交互式菜单**: 分层菜单设计，操作简单直观
- **实时反馈**: 详细的同步进度和结果显示
- **错误处理**: 完善的错误提示和异常处理

## 📦 安装方式

### 方式一：直接运行（推荐）

1. **克隆仓库**
```bash
git clone https://github.com/your-repo/sea-sheet-sync.git
cd sea-sheet-sync
```

2. **安装依赖**
```bash
pip install -r requirements.txt
```

3. **配置环境变量**
```bash
cp .env.example .env
# 编辑 .env 文件，填入你的配置
```

4. **运行程序**
```bash
python main-name-pro.py
```

### 方式二：使用预编译版本

从 [Releases](https://github.com/your-repo/sea-sheet-sync/releases) 页面下载对应平台的可执行文件：

- **Windows**: `seatable-sync-windows.exe`
- **Linux**: `seatable-sync-linux`
- **macOS**: `seatable-sync-macos`

下载后直接运行即可。

## ⚙️ 配置说明

### 环境变量配置 (.env)

```env
# SeaTable 服务器地址
SEATABLE_SERVER_URL=https://cloud.seatable.cn

# 默认 API Token
DEFAULT_SEATABLE_API_TOKEN=your_default_token_here

# 特定配置文件的 API Token（可选）
MEMO_ANA2025_SEATABLE_API_TOKEN=your_2025_token_here
MEMO_ANA2024_SEATABLE_API_TOKEN=your_2024_token_here
```

### JSON 配置文件结构

```json
{
  "menu_description": "2025年目标数据同步",
  "date_version": "20250331",
  "tables": [
    {
      "table_name": "SeaTable表名",
      "excel_directory": "/path/to/excel/files",
      "excel_file_name": "Excel文件名（不含扩展名）",
      "sheet_name": "工作表名称",
      "relation_field": "关联字段名",
      "relation_field_mappings": {
        "关联字段名": "Excel中对应的列名"
      },
      "field_mappings": {
        "SeaTable字段名": "Excel列名"
      }
    }
  ]
}
```

### 配置文件示例

**memo-ana2025.json**:
```json
{
  "menu_description": "2025年目标数据同步",
  "date_version": "20250331",
  "tables": [
    {
      "table_name": "A1-新签合同对比2025",
      "excel_directory": "/Users/user/Documents/Excel",
      "excel_file_name": "A1-2025年新签合同对比表",
      "sheet_name": "新签合同对比表",
      "relation_field": "销售组",
      "relation_field_mappings": {
        "销售组": "销售组"
      },
      "field_mappings": {
        "截止目前签单额": "截止目前签单额",
        "自有软件签单额": "自有软件签单额",
        "年度合理目标": "年度合理目标"
      }
    }
  ]
}
```

## 🎯 使用方法

### 1. 启动程序
```bash
python main-name-pro.py
```

### 2. 选择配置文件
程序会自动显示所有可用的配置文件：
```
=== SeaTable Excel 同步工具 ===
可用的配置文件:
1. 2025年目标数据同步 (memo-ana2025) (包含 11 个表格)
2. 2024年目标数据同步 (memo-ana2024) (包含 16 个表格)
0. 退出
```

### 3. 选择同步表格
选择配置文件后，程序直接显示该配置下的所有表格：
```
配置文件: 2025年目标数据同步 (memo-ana2025)
可用的表格:
1. A1-新签合同对比2025
2. A2-费用产出对比2025
3. 同步所有表格
0. 返回配置文件选择
```

### 4. 执行同步
选择表格后，程序自动执行同步操作并显示详细进度。

## 🔧 开发和构建

### 本地开发
```bash
# 安装开发依赖
pip install -r requirements.txt

# 运行程序
python main-name-pro.py
```

### 构建可执行文件

**本地构建**:
```bash
python build_standalone.py
```

**Windows CI 构建**:
```bash
python build_windows_ci.py
```

### 自动构建
项目配置了 GitHub Actions，在以下情况会自动构建：
- 推送到 main/master 分支
- 创建 Pull Request
- 发布 Release

构建产物会自动上传到 Artifacts 和 Release。

## 📁 项目结构

```
sea-sheet-sync/
├── main-name-pro.py          # 主程序文件
├── requirements.txt          # Python 依赖
├── .env.example             # 环境变量示例
├── memo-ana2025.json        # 2025年配置文件
├── memo-ana2024.json        # 2024年配置文件
├── build_standalone.py      # 独立构建脚本
├── build_windows_ci.py      # Windows CI 构建脚本
├── .github/workflows/build.yml  # GitHub Actions 配置
└── README.md               # 项目说明
```

## 🛡️ 安全性

- ✅ 敏感信息通过环境变量管理，不会硬编码
- ✅ 支持不同配置文件使用不同的 API Token
- ✅ .env 文件默认被 .gitignore 忽略
- ✅ API Token 具有最小权限原则

## 🐛 常见问题

### Q: 同步后数据数量不一致？
A: 请检查 SeaTable 表中关联字段是否有重复值，确保关联字段的唯一性。

### Q: 找不到 Excel 文件？
A: 请检查配置文件中的 `excel_directory` 和 `excel_file_name` 路径是否正确。

### Q: API Token 权限错误？
A: 请确认 API Token 对相应表格有读取权限，可以在 SeaTable 中重新生成 Token。

### Q: 程序无法找到配置文件？
A: 确保 JSON 配置文件与程序在同一目录下，且文件格式正确。

## 📝 更新日志

### v2.0.0 (当前版本)
- ✨ 新增多配置文件自动发现功能
- ✨ 新增配置文件菜单描述支持
- ✨ 新增独立 Token 配置支持
- 🔧 简化菜单操作流程
- 🔧 优化用户界面和交互体验
- 🐛 修复构建脚本兼容性问题

## 🤝 贡献

欢迎提交 Issues 和 Pull Requests 来改进这个项目。

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户。