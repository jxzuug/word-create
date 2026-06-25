# Word Create — Claude Code Word 文档生成 Skill 集合

在 Claude Code 中直接生成各类格式规范的 Word 文档（.docx），无需 Java、无需 Apache POI。

## 当前已实现的 Skill

### 📋 施工组织设计 / 服务方案

**Skill 名称**：`generate-construction-plan`

生成施工组织设计、施工方案、服务方案、投标技术文件、应急响应预案、售后服务方案等 Word 文档。

#### 使用方式

在 Claude Code 中直接说（不需要斜杠）：

```
我要投一个维修项目的标，帮我写技术方案
项目名称：XXX
服务内容：XXX
章节结构：
第一章、施工管理目标
第二章、施工组织部署
...
需要输出：约30页
```

或使用斜杠命令：

```
/施工组织设计
```

#### 功能特性

- ✅ **真实 .docx 格式**：基于 OOXML 标准，可用 Microsoft Word / WPS 直接打开
- ✅ **用户自定义章节**：章节结构和标题完全由用户指定
- ✅ **格式丰富**：支持封面、目录、段落（首行缩进）、多级标题、表格、列表、图片
- ✅ **图文结合**：支持嵌入图片（base64 编码），可用表格模拟流程图/组织架构图
- ✅ **字数/页数控制**：支持指定目标字数或页数，自动估算完成度
- ✅ **重复检测**：生成时自动扫描重复段落并告警
- ✅ **批量化处理**：支持超大型文档（100页以上）分批增量生成
- ✅ **格式规范**：A4 纸张、宋体/黑体、页眉页脚页码、1.5 倍行距

## 技术原理

```
用户输入项目信息 + 章节结构
        │
        ▼
Claude 生成各章节专业内容
        │
        ▼
构建结构化 JSON 数据
        │
        ▼
Node.js 脚本 (generate-docx.js)
  └─ docx 包 → 生成 .docx 文件
        │
        ▼
格式规范的 Word 文档 ✅
```

- 使用 Node.js + [`docx`](https://docx.js.org/) npm 包生成 .docx 文件
- 唯一依赖：`docx ^9.7.1`
- 无需 Python、Java、Apache POI、pandoc 等外部工具

## 安装

```bash
# 克隆项目
git clone https://github.com/jxzuug/word-create.git
cd word-create

# 安装依赖
npm install
```

安装后即可在 Claude Code 中使用。

## 项目结构

```
word-create/
├── .claude/
│   └── skills/
│       └── generate-construction-plan/
│           ├── SKILL.md                  # Skill 定义（交互流程 + 内容规范）
│           └── scripts/
│               └── generate-docx.js      # Word 文档生成脚本
├── .gitignore
├── package.json
├── package-lock.json
└── README.md
```

## 扩展新的 Word 文档类型

本项目的目标是成为多种 Word 文档生成 Skill 的集合。要添加新的文档类型：

1. 在 `.claude/skills/` 下创建新的 Skill 目录
2. 编写 `SKILL.md`（name + description + 生成规则）
3. 复用 `scripts/generate-docx.js` 作为文档生成引擎

## 依赖

- [docx](https://docx.js.org/) — 在 Node.js 中创建 Word 文档的库
- Node.js >= 18
