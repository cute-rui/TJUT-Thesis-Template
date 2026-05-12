# 天津理工大学论文模板

基于 Node.js + `docx` 库，通过代码生成符合**天津理工大学格式要求**的论文 Word 文档。

> **核心思路：** 用代码代替手动排版，实现论文格式的统一维护和反复生成。

## 项目结构

| 文件 | 说明 |
| --- | --- |
| **`generate_thesis.js`** | ⭐ **正式论文模板** — 直接在此文件中填充**题名、摘要、目录、正文、图表、公式、参考文献、致谢**等全部内容 |
| `generate_thesis.example.js` | 功能演示，仅供学习函数调用方式（格式不完整，勿用于正式论文） |
| `图/` | 论文插图资源目录 |

## 快速开始

```powershell
pnpm install
pnpm start           # 生成正式论文
pnpm run example     # 运行示例（可选）
```

生成的 `.docx` 文件已被 `.gitignore` 排除，不会进入版本控制。

## 写作流程

1. 阅读 `generate_thesis.js` 的结构，了解各章节对应位置
2. **直接在 `generate_thesis.js` 中填充题名、摘要、目录、正文、图表、公式、参考文献、致谢等内容**
3. 执行 `pnpm start` 生成 Word 文档
4. **保持原有格式函数、分页逻辑和标题层级不变**，仅替换正文内容

> [!WARNING]
> 不要使用 Python 拆包 `.docx` 或手动修改其内部 XML，以免结构损坏。所有内容应通过模板函数生成。

## 关于图表标题字体

> [!NOTE]
> 天津理工大学本科毕业设计说明书（毕业论文）撰写规范（2025）并未明确说明**图和表的标题在包含英文时是否应使用楷体**。模板默认使用 `KAITI`（中文楷体 + 英文 Times New Roman）。
>
> 如果希望英文部分**也显示为楷体**，可在 `figure()` 和 `tableCaption()` 函数内将 `font` 从 `KAITI` 更改为 `TITLE_FONT_PURE_KAITI`（纯楷体）：
>
> ```js
> // 修改前（默认：英文 Times New Roman + 中文楷体）
> new TextRun({ text: title, font: KAITI, size: WU_HAO })
>
> // 修改后（全部使用楷体）
> new TextRun({ text: title, font: TITLE_FONT_PURE_KAITI, size: WU_HAO })
> ```

## 自定义字体名称

> [!TIP]
> 如果需要使用其他变体的宋体、楷体等字体（如"华文宋体"、"华文楷体"、"仿宋"等），可直接修改 `generate_thesis.js` **开头的 `const` 定义**：
>
> ```js
> // 默认定义
> const SONGTI = { ascii: TNR, hAnsi: TNR, eastAsia: "宋体", cs: TNR };
> const HEITI  = { ascii: TNR, hAnsi: TNR, eastAsia: "黑体", cs: TNR };
> const KAITI  = { ascii: TNR, hAnsi: TNR, eastAsia: "楷体", cs: TNR };
>
> // 示例：改用 GB2312 / 小标宋字体
> const SONGTI = { ascii: TNR, hAnsi: TNR, eastAsia: "方正小标宋_GBK", cs: TNR };
> const KAITI  = { ascii: TNR, hAnsi: TNR, eastAsia: "楷体_GB2312", cs: TNR };
> ```
>
> 修改后，模板中所有引用这些常量的地方会自动应用新字体，无需逐处更改。

## 工具链

| 工具 | 用途 | 链接 |
| --- | --- | --- |
| **pnpm** | 包管理器（必需） | [pnpm.io](https://pnpm.io/installation) |
| Pandoc | 文档格式转换（可选） | [GitHub](https://github.com/jgm/pandoc/releases) |
| ripgrep | 项目内容检索（可选） | [GitHub](https://github.com/BurntSushi/ripgrep/releases) |

## 建议安装的 MCP 与 Skills

配合 AI 编码助手（如 Claude）使用时，推荐安装以下 MCP 服务器和 Skills 以提升论文写作效率：

### Skills

| 名称 | 用途 | 链接 |
| --- | --- | --- |
| **docx** | Anthropic 官方 docx 技能，支持 Word 文档的创建、编辑与分析 | [GitHub](https://github.com/anthropics/skills) |
| **cnki-skills** | 知网相关技能，辅助中文文献检索与引用 | [GitHub](https://github.com/cookjohn/cnki-skills) |

### MCP 服务器

| 名称 | 用途 | 链接 |
| --- | --- | --- |
| **chrome-devtools-mcp** | 通过 Chrome DevTools 协议调试和操控浏览器 | [GitHub](https://github.com/ChromeDevTools/chrome-devtools-mcp) |
| **paper-search-mcp** | 学术论文搜索（支持 Semantic Scholar、arXiv 等） | [GitHub](https://github.com/openags/paper-search-mcp) |

> [!WARNING]
> `paper-search-mcp` 的 PyPI 源已过时，**必须**使用 git 方式安装：
> ```
> uvx --from git+https://github.com/openags/paper-search-mcp.git paper-search-mcp
> ```

## 阅读批注（Comments）

当导师在 Word 文档中添加修改批注后，可通过 Pandoc 将 `.docx` 转为 Markdown，批注会以特定格式保留在文本中，再用 ripgrep 搜索：

```powershell
# 1. 转换为 Markdown（保留批注）
pandoc 论文初稿.docx -t markdown --wrap=none -o 论文初稿.md

# 2. 搜索批注内容
rg "comment-start|comment-end|\[.*\]\{\.comment-start\}" 论文初稿.md
```

> [!WARNING]
> 尽可能**不要通过 Python 解包 `.docx`** 来查看文件内容或批注，以免损坏文档内部 XML 结构。应始终使用 Pandoc 进行格式转换。
