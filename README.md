# WordNumberFormatter — Word 数字千分符格式化加载项

一个 Word Office Web Add-in（任务窗格 + 功能命令），为文档中符合条件的数字添加英文千分符。

## 功能特性

- **菜单栏"千分符"SplitButton**（"开始"选项卡）：
  - 主按钮：一键为正文和表格中的所有数字添加千分符
  - 下拉"文档替换"：仅处理正文部分（自动排除表格）
  - 下拉"表格替换"：仅处理表格部分（支持嵌套表格）
- **任务窗格**：提供"文档替换 / 表格替换"两个按钮，带处理状态提示
- **匹配规则**：
  - 整数部分 ≥ 4 位（如 `1234.56` 会被处理）
  - 小数部分恰好 2 位
  - 数字前后不能是数字或小数点（避免误拆 `12.341234.56` 一类的数字串）
  - 示例：`1234567.89` → `1,234,567.89`
- **幂等安全**：已格式化的数字（含逗号）不会再被匹配，重复点击不会二次处理

## 技术栈

- Office JavaScript API（Word，hosted office.js）
- JavaScript（Babel 转译）
- webpack 5 + HtmlWebpackPlugin（多入口打包：taskpane / commands / polyfill）
- Yeoman generator-office 脚手架

## 快速开始

### 环境要求

- Node.js（LTS 版本）
- Word 桌面版（Microsoft 365）

### 步骤

1. 安装依赖：

   ```bash
   npm install
   ```

2. 启动开发服务器（自动编译、启动本地 HTTPS 服务并旁加载到 Word）：

   ```bash
   npm start
   ```

   首次运行会提示安装开发证书，请选择"是"。

3. Word 启动后，在"开始"选项卡点击"千分符"按钮，或打开右侧任务窗格操作。

4. 调试结束后在终端按 `Ctrl+C` 停止；生产构建：

   ```bash
   npm run build
   ```

## 处理流程

1. **搜索**：使用 Word 通配符 `[!0-9.][0-9]{4,}\.[0-9][0-9][!0-9.]` 在正文/单元格中定位候选数字
2. **过滤**：正文处理时通过 `parentTableOrNullObject` 排除表格内范围
3. **格式化**：JS 正则提取数字，整数部分按 3 位一组插入逗号
4. **替换**：`range.insertText` 批量提交，保留原前后边界字符

表格处理按 `body.tables → table.rows → row.cells` 遍历单元格，并通过 `table.tables` 递归处理嵌套表格；逐层提交替换，避免内外层匹配范围重叠导致重复替换。

## 项目结构

```
WordNumberFormatter/
├── manifest.xml              # 加载项清单（SplitButton 按钮组定义）
├── webpack.config.js         # 多入口打包配置
├── package.json
├── assets/                   # 菜单图标（16/32/64/80/128）
└── src/
    ├── commands/             # Ribbon 按钮回调（commands.html/js）
    ├── shared/
    │   └── number-formatter.js   # 核心逻辑：搜索、过滤、千分符替换
    └── taskpane/             # 任务窗格 UI（taskpane.html/js/css）
```

## 安全说明

- 项目不含任何密钥、凭据或敏感配置
- 开发阶段所有资源走 `https://localhost:3000`（webpack devServer 自签名证书）
- 生产部署前，请将 `webpack.config.js` 中的 `urlProd` 改为实际部署地址，`npm run build` 会自动替换 manifest 中的 URL

## 已知限制

- 数字位于文档最开头/最末尾（无任何相邻字符）时可能不被匹配——这是 Word 通配符搜索的固有限制
- 开发模式下加载项依赖 devServer，需通过 `npm start` 启动使用；永久安装可将 `manifest.xml` 通过 Word"文件 → 选项 → 加载项"手动旁加载

## 变更日志

### v1.0.0（2026-09-24）

- 初始版本：菜单栏 SplitButton + 任务窗格，正文/表格（含嵌套）数字千分符格式化
- 修正开发文档 6 处缺陷：模板字符串漏 `${}`（2 处）、Word 通配符字面小数点转义、manifest `<ExtensionPoint>` 错误嵌套进 `<FunctionFile>`、共享模块 webpack 打包方式（ES module + `Office.actions.associate`）、表格单元格遍历方式（`table.rows → row.cells`）