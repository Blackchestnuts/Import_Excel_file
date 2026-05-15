# Excel 导入校验工具

基于 Next.js + TypeScript 的 Excel 导入数据校验 Web 应用，实现了两阶段校验架构（文件级阻塞校验 + 行级遍历校验），支持多种校验规则。

## 功能特性

### 两阶段校验架构

1. **文件级阻塞校验**：校验不通过则整单拒绝
   - 文件格式校验（仅支持 .xls / .xlsx）
   - 模板匹配校验（表头列数与列名必须与模板一致）
   - 行数限制校验

2. **行级遍历校验**：逐行逐列校验，收集所有错误
   - **必填校验**：必填字段为空时跳过后续校验
   - **数据类型校验**：STRING / NUMBER / DATE / ENUM
   - **长度校验**：最小长度、最大长度
   - **正则校验**：手机号、身份证号等格式
   - **枚举校验**：值必须在可选列表内
   - **文件内重复校验**：基于唯一键组的 HashMap 去重

### 预设模板

- **员工信息导入**：姓名、年龄、手机号、入职日期、部门
- **学生信息导入**：学号、姓名、性别、班级、手机号

### 其他功能

- 模板下载（自动生成含示例数据和规则提示的 Excel 模板）
- 拖拽上传文件
- 校验结果可视化（通过率、错误类型分布、错误明细、有效数据预览）
- 错误按行号分组，可折叠查看详情

## 技术栈

- **框架**：Next.js 16 (App Router)
- **语言**：TypeScript
- **样式**：Tailwind CSS 4 + shadcn/ui
- **Excel 解析**：SheetJS (xlsx)
- **设计模式**：策略模式（每种校验规则独立实现）

## 项目结构

```
src/
├── app/
│   ├── api/
│   │   ├── validate/route.ts    # 校验 API
│   │   └── template/route.ts    # 模板下载 API
│   ├── layout.tsx
│   ├── page.tsx                 # 主页面
│   └── globals.css
├── lib/
│   ├── validator/
│   │   ├── types.ts             # 类型定义 & 预设模板配置
│   │   └── engine.ts            # 校验引擎核心逻辑
│   └── utils.ts
└── components/
    └── ui/                      # shadcn/ui 组件
```

## 快速开始

### 安装依赖

```bash
npm install
# 或
bun install
```

### 启动开发服务器

```bash
npm run dev
```

访问 [http://localhost:3000](http://localhost:3000) 查看应用。

### 构建生产版本

```bash
npm run build
npm start
```

## 校验规则配置

在 `src/lib/validator/types.ts` 中的 `PRESET_TEMPLATES` 数组中添加或修改模板配置：

```typescript
{
  templateCode: "YOUR_TEMPLATE_CODE",
  templateName: "模板名称",
  expectedHeaders: ["列1", "列2", "列3"],
  columnRules: {
    "列1": {
      columnName: "列1",
      required: true,
      dataType: "STRING",
      maxLength: 50,
    },
    // ...
  },
  uniqueKeyGroups: [["列1"]], // 唯一性约束
  maxRows: 10000,
}
```

## 校验错误类型

| 错误类型 | 说明 |
|---------|------|
| FILE_FORMAT | 文件格式不正确 |
| TEMPLATE_MISMATCH | 模板不匹配 |
| REQUIRED_EMPTY | 必填字段为空 |
| TYPE_MISMATCH | 数据类型不匹配 |
| LENGTH_EXCEEDED | 长度超限 |
| REGEX_MISMATCH | 格式不正确（正则不匹配）|
| ENUM_INVALID | 枚举值非法 |
| DUPLICATE_IN_FILE | 文件内重复 |

## License

MIT
