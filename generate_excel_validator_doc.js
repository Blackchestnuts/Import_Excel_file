const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, NumberFormat,
  AlignmentType, HeadingLevel, WidthType, BorderStyle, ShadingType,
  PageBreak, TableOfContents, LevelFormat,
} = require("docx");
const fs = require("fs");

// ── Palette ──
const P = {
  primary: "0A1628",
  body: "1A2B40",
  secondary: "6878A0",
  accent: "5B8DB8",
  surface: "F4F8FC",
};

const c = (hex) => hex;

// ── Helper functions ──
function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({
    heading: level,
    spacing: {
      before: level === HeadingLevel.HEADING_1 ? 360 : level === HeadingLevel.HEADING_2 ? 240 : 200,
      after: 120,
      line: 312,
    },
    children: [
      new TextRun({
        text,
        bold: true,
        color: c(P.primary),
        font: { ascii: "Calibri", eastAsia: "SimHei" },
        size: level === HeadingLevel.HEADING_1 ? 32 : level === HeadingLevel.HEADING_2 ? 28 : 24,
      }),
    ],
  });
}

function bodyPara(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { line: 312, after: 80 },
    ...opts,
    children: [
      new TextRun({
        text,
        size: 24,
        color: c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
      }),
    ],
  });
}

function codeBlock(code) {
  const lines = code.split("\n");
  return lines.map(
    (line) =>
      new Paragraph({
        spacing: { line: 240, after: 0 },
        indent: { left: 480 },
        children: [
          new TextRun({
            text: line || " ",
            size: 18,
            font: { ascii: "Consolas", eastAsia: "Consolas" },
            color: "1A1A1A",
          }),
        ],
      })
  );
}

function bulletItem(text, level = 0) {
  return new Paragraph({
    bullet: { level },
    spacing: { line: 312, after: 40 },
    children: [
      new TextRun({
        text,
        size: 24,
        color: c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
      }),
    ],
  });
}

function tableCell(text, opts = {}) {
  const isHeader = opts.header || false;
  return new TableCell({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text,
            bold: isHeader,
            size: isHeader ? 21 : 20,
            color: isHeader ? "FFFFFF" : c(P.body),
            font: { ascii: "Calibri", eastAsia: isHeader ? "SimHei" : "Microsoft YaHei" },
          }),
        ],
        spacing: { line: 280 },
      }),
    ],
    shading: isHeader
      ? { type: ShadingType.CLEAR, fill: P.accent }
      : opts.shaded
      ? { type: ShadingType.CLEAR, fill: P.surface }
      : undefined,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
  });
}

function makeTable(headers, rows, colWidths) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: P.accent },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: P.accent },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "D0D0D0" },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: headers.map((h, i) =>
          tableCell(h, { header: true, width: colWidths?.[i] })
        ),
      }),
      ...rows.map((row, ri) =>
        new TableRow({
          cantSplit: true,
          children: row.map((cell, ci) =>
            tableCell(cell, { shaded: ri % 2 === 0, width: colWidths?.[ci] })
          ),
        })
      ),
    ],
  });
}

// ── Cover ──
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const allNoBorders = {
  top: NB, bottom: NB, left: NB, right: NB,
  insideHorizontal: NB, insideVertical: NB,
};

function buildCover() {
  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: allNoBorders,
      rows: [
        new TableRow({
          height: { value: 16838, rule: "exact" },
          children: [
            new TableCell({
              borders: allNoBorders,
              width: { size: 100, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  spacing: { before: 4800, line: 828, lineRule: "atLeast" },
                  alignment: AlignmentType.LEFT,
                  indent: { left: 1200 },
                  children: [
                    new TextRun({
                      text: "Excel\u5BFC\u5165\u6821\u9A8C",
                      bold: true,
                      size: 72,
                      color: c(P.accent),
                      font: { ascii: "Calibri", eastAsia: "SimHei" },
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 200, line: 600, lineRule: "atLeast" },
                  alignment: AlignmentType.LEFT,
                  indent: { left: 1200 },
                  children: [
                    new TextRun({
                      text: "Java \u4EE3\u7801\u8BBE\u8BA1\u65B9\u6848",
                      size: 40,
                      color: c(P.secondary),
                      font: { ascii: "Calibri", eastAsia: "SimHei" },
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 600 },
                  alignment: AlignmentType.LEFT,
                  indent: { left: 1200 },
                  border: {
                    bottom: { style: BorderStyle.SINGLE, size: 6, color: P.accent, space: 12 },
                  },
                  children: [],
                }),
                new Paragraph({
                  spacing: { before: 400, line: 312 },
                  alignment: AlignmentType.LEFT,
                  indent: { left: 1200 },
                  children: [
                    new TextRun({
                      text: "\u57FA\u4E8E EasyExcel + \u7B56\u7565\u6A21\u5F0F\u7684\u6587\u4EF6\u5BFC\u5165\u6821\u9A8C\u6846\u67B6",
                      size: 24,
                      color: c(P.secondary),
                      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { before: 100 },
                  alignment: AlignmentType.LEFT,
                  indent: { left: 1200 },
                  children: [
                    new TextRun({
                      text: "2026\u5E745\u6708",
                      size: 22,
                      color: c(P.secondary),
                      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

// ── TOC Section ──
function buildTocSection() {
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [
        new TextRun({
          text: "\u76EE\u5F55",
          bold: true,
          size: 32,
          color: c(P.primary),
          font: { ascii: "Calibri", eastAsia: "SimHei" },
        }),
      ],
    }),
    new TableOfContents("\u76EE\u5F55", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "\u63D0\u793A\uFF1A\u8BF7\u53F3\u952E\u76EE\u5F55\u5E76\u9009\u62E9\u201C\u66F4\u65B0\u57DF\u201D\u4EE5\u5237\u65B0\u9875\u7801",
          italics: true,
          size: 18,
          color: "999999",
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
        }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new PageBreak()],
    }),
  ];
}

// ── Body content ──
function buildBody() {
  const children = [];

  // ═══ Chapter 1: Architecture Overview ═══
  children.push(heading("\u4E00\u3001\u67B6\u6784\u8BBE\u8BA1\u6982\u8FF0"));

  children.push(bodyPara(
    "\u672C\u65B9\u6848\u9488\u5BF9 Excel \u6587\u4EF6\u5BFC\u5165\u573A\u666F\uFF0C\u8BBE\u8BA1\u4E86\u4E00\u5957\u5206\u5C42\u3001\u53EF\u6269\u5C55\u7684\u6821\u9A8C\u6846\u67B6\u3002\u6838\u5FC3\u539F\u5219\u662F\uFF1A\u524D\u7F6E\u6821\u9A8C\uFF08\u6587\u4EF6\u683C\u5F0F\u3001\u6A21\u677F\u5339\u914D\uFF09\u5931\u8D25\u5219\u5FEB\u901F\u4E2D\u65AD\uFF1B\u4E1A\u52A1\u6821\u9A8C\uFF08\u6570\u636E\u5B8C\u6574\u6027\u3001\u7C7B\u578B\u3001\u957F\u5EA6\u3001\u91CD\u590D\uFF09\u9700\u6536\u96C6\u6240\u6709\u9519\u8BEF\u5E76\u8FD4\u56DE\uFF0C\u907F\u514D\u7528\u6237\u9677\u5165\u201C\u6539\u4E00\u4E2A\u9519\u518D\u5BFC\u4E00\u6B21\u201D\u7684\u6B7B\u5FAA\u73AF\u3002\u6574\u4F53\u6821\u9A8C\u5206\u4E3A\u4E24\u4E2A\u9636\u6BB5\uFF1A\u6587\u4EF6\u7EA7\u963B\u65AD\u6821\u9A8C\u548C\u884C\u7EA7\u904D\u5386\u6821\u9A8C\uFF0C\u4E24\u4E2A\u9636\u6BB5\u5404\u53F8\u5176\u804C\uFF0C\u4E92\u4E0D\u4FB5\u5165\u3002"
  ));

  children.push(heading("1.1 \u8BBE\u8BA1\u539F\u5219", HeadingLevel.HEADING_2));

  children.push(makeTable(
    ["\u539F\u5219", "\u8BF4\u660E"],
    [
      ["\u5FEB\u901F\u5931\u8D25\uFF08Fail-Fast\uFF09", "\u6587\u4EF6\u683C\u5F0F\u3001\u6A21\u677F\u5339\u914D\u7B49\u6574\u4F53\u6027\u9519\u8BEF\u76F4\u63A5\u4E2D\u65AD\uFF0C\u4E0D\u8FDB\u5165\u884C\u7EA7\u6821\u9A8C\uFF0C\u907F\u514D\u65E0\u8C13\u8BA1\u7B97"],
      ["\u9519\u8BEF\u5168\u91CF\u6536\u96C6", "\u884C\u7EA7\u6821\u9A8C\u9636\u6BB5\u4E0D\u4E2D\u65AD\uFF0C\u6536\u96C6\u6240\u6709\u884C\u7684\u6240\u6709\u9519\u8BEF\u4FE1\u606F\uFF0C\u4E00\u6B21\u6027\u8FD4\u56DE\u7ED9\u524D\u7AEF"],
      ["\u7B56\u7565\u6A21\u5F0F\u89E3\u8026", "\u6BCF\u79CD\u6821\u9A8C\u89C4\u5219\u5C01\u88C5\u4E3A\u72EC\u7ACB\u7684 Validator\uFF0C\u65B0\u589E\u89C4\u5219\u53EA\u9700\u65B0\u589E\u5B9E\u73B0\u7C7B"],
      ["\u914D\u7F6E\u5316\u9A71\u52A8", "\u6821\u9A8C\u89C4\u5219\u901A\u8FC7\u914D\u7F6E\u5BF9\u8C61\u5B9A\u4E49\uFF0C\u800C\u975E\u786C\u7F16\u7801\u5728\u4EE3\u7801\u4E2D\uFF0C\u652F\u6301\u52A8\u6001\u6269\u5C55"],
      ["SAX \u6D41\u5F0F\u8BFB\u53D6", "\u4F7F\u7528 EasyExcel \u7684\u76D1\u542C\u5668\u6A21\u5F0F\u9010\u884C\u89E3\u6790\uFF0C\u9632\u6B62\u5927\u6587\u4EF6 OOM"],
    ],
    [30, 70]
  ));

  children.push(heading("1.2 \u6821\u9A8C\u6D41\u7A0B", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6574\u4F53\u6821\u9A8C\u6D41\u7A0B\u5206\u4E3A\u4E24\u4E2A\u9636\u6BB5\u3002\u7B2C\u4E00\u9636\u6BB5\u662F\u6587\u4EF6\u7EA7\u963B\u65AD\u6821\u9A8C\uFF1A\u63A5\u6536\u5230\u4E0A\u4F20\u6587\u4EF6\u540E\uFF0C\u9996\u5148\u68C0\u67E5\u6587\u4EF6\u683C\u5F0F\u662F\u5426\u4E3A\u5408\u6CD5\u7684 Excel \u683C\u5F0F\uFF08.xls \u6216 .xlsx\uFF09\uFF0C\u7136\u540E\u68C0\u67E5\u8868\u5934\u662F\u5426\u4E0E\u9884\u8BBE\u6A21\u677F\u5339\u914D\u3002\u4EFB\u4F55\u4E00\u9879\u5931\u8D25\u5747\u6807\u8BB0\u4E3A\u201C\u5168\u90E8\u5BFC\u5165\u5931\u8D25\u201D\u5E76\u7ACB\u5373\u8FD4\u56DE\uFF0C\u4E0D\u518D\u6267\u884C\u540E\u7EED\u903B\u8F91\u3002\u7B2C\u4E8C\u9636\u6BB5\u662F\u884C\u7EA7\u904D\u5386\u6821\u9A8C\uFF1A\u901A\u8FC7 EasyExcel \u7684 ReadListener \u9010\u884C\u8BFB\u53D6\u6570\u636E\uFF0C\u5BF9\u6BCF\u4E00\u884C\u6267\u884C\u5FC5\u586B\u6821\u9A8C\u3001\u7C7B\u578B\u6821\u9A8C\u3001\u957F\u5EA6\u6821\u9A8C\u3001\u6587\u4EF6\u5185\u53BB\u91CD\u548C\u6570\u636E\u5E93\u53BB\u91CD\uFF0C\u6240\u6709\u9519\u8BEF\u4FE1\u606F\u7EDF\u4E00\u6536\u96C6\u3002\u6700\u7EC8\u5C06\u6821\u9A8C\u7ED3\u679C\u5305\u88C5\u4E3A ImportResult \u8FD4\u56DE\u7ED9\u8C03\u7528\u65B9\u3002"
  ));

  // ═══ Chapter 2: Core Model Design ═══
  children.push(heading("\u4E8C\u3001\u6838\u5FC3\u6A21\u578B\u8BBE\u8BA1"));

  children.push(heading("2.1 \u6821\u9A8C\u914D\u7F6E\u6A21\u578B", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6821\u9A8C\u914D\u7F6E\u6A21\u578B\u662F\u6574\u4E2A\u6846\u67B6\u7684\u57FA\u7840\uFF0C\u5B83\u5C06\u6821\u9A8C\u89C4\u5219\u4E0E\u4EE3\u7801\u903B\u8F91\u89E3\u8026\u3002\u6BCF\u4E2A\u5BFC\u5165\u6A21\u677F\u5BF9\u5E94\u4E00\u5957\u914D\u7F6E\uFF0C\u5305\u542B\u8868\u5934\u5217\u5B9A\u4E49\u3001\u5404\u5217\u7684\u6821\u9A8C\u89C4\u5219\uFF08\u5FC5\u586B\u3001\u7C7B\u578B\u3001\u6700\u5927\u957F\u5EA6\u3001\u6B63\u5219\u8868\u8FBE\u5F0F\u7B49\uFF09\u4EE5\u53CA\u552F\u4E00\u6027\u6821\u9A8C\u7684\u7EC4\u5408\u952E\u3002\u8FD9\u4E9B\u914D\u7F6E\u53EF\u4EE5\u5B58\u50A8\u5728\u6570\u636E\u5E93\u4E2D\uFF0C\u4E5F\u53EF\u4EE5\u901A\u8FC7\u914D\u7F6E\u6587\u4EF6\u7BA1\u7406\uFF0C\u652F\u6301\u8FD0\u884C\u65F6\u52A8\u6001\u8C03\u6574\u6821\u9A8C\u89C4\u5219\u800C\u65E0\u9700\u91CD\u65B0\u90E8\u7F72\u3002"
  ));

  children.push(...codeBlock(`/**
 * \u5217\u6821\u9A8C\u89C4\u5219\u914D\u7F6E
 */
@Data
@Builder
public class ColumnRule {
    /** \u5217\u540D\uFF08\u5BF9\u5E94\u8868\u5934\uFF09 */
    private String columnName;
    /** \u662F\u5426\u5FC5\u586B */
    private boolean required;
    /** \u6570\u636E\u7C7B\u578B\uFF1ASTRING / NUMBER / DATE / ENUM */
    private DataType dataType;
    /** \u6700\u5927\u957F\u5EA6 */
    private Integer maxLength;
    /** \u6700\u5C0F\u957F\u5EA6 */
    private Integer minLength;
    /** \u81EA\u5B9A\u4E49\u6B63\u5219\u6821\u9A8C */
    private String regexPattern;
    /** \u679A\u4E3E\u503C\u8303\u56F4\uFF08\u5F53 dataType=ENUM \u65F6\u751F\u6548\uFF09 */
    private List<String> enumValues;
    /** \u81EA\u5B9A\u4E49\u9519\u8BEF\u63D0\u793A\u6A21\u677F */
    private String errorMsg;
}

/**
 * \u5BFC\u5165\u6A21\u677F\u914D\u7F6E
 */
@Data
@Builder
public class ImportTemplateConfig {
    /** \u6A21\u677F\u7F16\u7801\uFF0C\u7528\u4E8E\u5339\u914D\u4E0D\u540C\u4E1A\u52A1\u573A\u666F */
    private String templateCode;
    /** \u9884\u671F\u7684\u8868\u5934\u5217\u987A\u5E8F */
    private List<String> expectedHeaders;
    /** \u6BCF\u5217\u7684\u6821\u9A8C\u89C4\u5219 */
    private Map<String, ColumnRule> columnRules;
    /** \u552F\u4E00\u6027\u6821\u9A8C\u7684\u5B57\u6BB5\u7EC4\u5408\uFF08\u652F\u6301\u8054\u5408\u552F\u4E00\uFF09 */
    private List<List<String>> uniqueKeyGroups;
    /** \u6700\u5927\u5141\u8BB8\u5BFC\u5165\u884C\u6570 */
    private int maxRows;
}

/**
 * \u6570\u636E\u7C7B\u578B\u679A\u4E3E
 */
public enum DataType {
    STRING, NUMBER, DATE, ENUM
}`));

  children.push(heading("2.2 \u6821\u9A8C\u7ED3\u679C\u6A21\u578B", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6821\u9A8C\u7ED3\u679C\u6A21\u578B\u8D1F\u8D23\u627F\u8F7D\u6821\u9A8C\u8FC7\u7A0B\u4E2D\u7684\u6240\u6709\u4FE1\u606F\u3002\u5B83\u533A\u5206\u4E86\u201C\u5168\u90E8\u5BFC\u5165\u5931\u8D25\u201D\u548C\u201C\u90E8\u5206\u5BFC\u5165\u5931\u8D25\u201D\u4E24\u79CD\u573A\u666F\uFF0C\u5E76\u5355\u72EC\u6536\u96C6\u6BCF\u4E00\u884C\u7684\u9519\u8BEF\u4FE1\u606F\u548C\u6821\u9A8C\u901A\u8FC7\u7684\u6570\u636E\uFF0C\u4FBF\u4E8E\u524D\u7AEF\u5C55\u793A\u9519\u8BEF\u660E\u7EC6\u5E76\u5C06\u6709\u6548\u6570\u636E\u5199\u5165\u6570\u636E\u5E93\u3002"
  ));

  children.push(...codeBlock(`/**
 * \u5355\u884C\u9519\u8BEF\u4FE1\u606F
 */
@Data
@AllArgsConstructor
public class RowError {
    /** \u539F\u59CB Excel \u884C\u53F7\uFF08\u4ECE1\u5F00\u59CB\uFF0C\u52A0\u4E0A\u8868\u5934\u5360\u4E00\u884C\uFF0C\u6240\u4EE5\u5B9E\u9645\u884C\u53F7 = rowIndex + 2\uFF09 */
    private int rowIndex;
    /** \u5217\u540D */
    private String columnName;
    /** \u9519\u8BEF\u7C7B\u578B */
    private ValidationErrorType errorType;
    /** \u9519\u8BEF\u63CF\u8FF0 */
    private String errorMessage;
}

/**
 * \u9519\u8BEF\u7C7B\u578B\u679A\u4E3E
 */
public enum ValidationErrorType {
    FILE_FORMAT,       // \u6587\u4EF6\u683C\u5F0F\u9519\u8BEF
    TEMPLATE_MISMATCH, // \u6A21\u677F\u4E0D\u5339\u914D
    REQUIRED_EMPTY,    // \u5FC5\u586B\u4E3A\u7A7A
    TYPE_MISMATCH,     // \u7C7B\u578B\u4E0D\u5339\u914D
    LENGTH_EXCEEDED,   // \u957F\u5EA6\u8D85\u9650
    DUPLICATE_IN_FILE, // \u6587\u4EF6\u5185\u91CD\u590D
    DUPLICATE_IN_DB,   // \u4E0E\u6570\u636E\u5E93\u6570\u636E\u91CD\u590D
    REGEX_MISMATCH,    // \u6B63\u5219\u6821\u9A8C\u4E0D\u901A\u8FC7
    ENUM_INVALID,      // \u679A\u4E3E\u503C\u4E0D\u5408\u6CD5
}

/**
 * \u5BFC\u5165\u6821\u9A8C\u7ED3\u679C
 */
@Data
@Builder
public class ImportResult {
    /** \u662F\u5426\u5168\u90E8\u5BFC\u5165\u5931\u8D25\uFF08\u6587\u4EF6\u7EA7\u9519\u8BEF\uFF09 */
    private boolean totalFail;
    /** \u6821\u9A8C\u901A\u8FC7\u7684\u884C\u6570 */
    private int successCount;
    /** \u6821\u9A8C\u5931\u8D25\u7684\u884C\u6570 */
    private int failCount;
    /** \u6240\u6709\u9519\u8BEF\u4FE1\u606F */
    private List<RowError> errors;
    /** \u6821\u9A8C\u901A\u8FC7\u7684\u6570\u636E\uFF08\u7528\u4E8E\u540E\u7EED\u5199\u5165\u6570\u636E\u5E93\uFF09 */
    private List<Map<String, String>> validRows;
}`));

  // ═══ Chapter 3: Core Validation Framework ═══
  children.push(heading("\u4E09\u3001\u6838\u5FC3\u6821\u9A8C\u6846\u67B6"));

  children.push(heading("3.1 \u6821\u9A8C\u5668\u63A5\u53E3\u4E0E\u62BD\u8C61\u57FA\u7C7B", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u91C7\u7528\u7B56\u7565\u6A21\u5F0F\uFF0C\u5C06\u6BCF\u79CD\u6821\u9A8C\u89C4\u5219\u5C01\u88C5\u4E3A\u72EC\u7ACB\u7684 Validator\u3002\u6240\u6709\u6821\u9A8C\u5668\u5B9E\u73B0\u7EDF\u4E00\u7684\u63A5\u53E3\uFF0C\u6838\u5FC3\u65B9\u6CD5\u63A5\u6536\u5F53\u524D\u884C\u6570\u636E\u548C\u5217\u89C4\u5219\u914D\u7F6E\uFF0C\u8FD4\u56DE\u6821\u9A8C\u9519\u8BEF\u5217\u8868\u3002\u65B0\u589E\u6821\u9A8C\u89C4\u5219\u53EA\u9700\u65B0\u5EFA\u4E00\u4E2A\u5B9E\u73B0\u7C7B\uFF0C\u65E0\u9700\u4FEE\u6539\u5DF2\u6709\u4EE3\u7801\uFF0C\u7B26\u5408\u5F00\u95ED\u539F\u5219\u3002"
  ));

  children.push(...codeBlock(`/**
 * \u6821\u9A8C\u5668\u7EDF\u4E00\u63A5\u53E3
 */
public interface CellValidator {
    /**
     * \u6821\u9A8C\u5355\u4E2A\u5355\u5143\u683C
     * @param rowIndex    \u884C\u53F7\uFF08Excel \u5B9E\u9645\u884C\u53F7\uFF0C\u4ECE 1 \u5F00\u59CB\uFF09
     * @param columnName  \u5217\u540D
     * @param cellValue   \u5355\u5143\u683C\u503C\uFF08\u5DF2\u8F6C\u4E3A\u5B57\u7B26\u4E32\uFF09
     * @param rule        \u8BE5\u5217\u7684\u6821\u9A8C\u89C4\u5219
     * @return \u9519\u8BEF\u5217\u8868\uFF0C\u7A7A\u5217\u8868\u8868\u793A\u901A\u8FC7
     */
    List<RowError> validate(int rowIndex, String columnName,
                            String cellValue, ColumnRule rule);
}

/**
 * \u884C\u7EA7\u6821\u9A8C\u5668\u63A5\u53E3\uFF08\u7528\u4E8E\u8DE8\u5217\u6821\u9A8C\uFF0C\u5982\u552F\u4E00\u6027\u6821\u9A8C\uFF09
 */
public interface RowValidator {
    /**
     * \u6821\u9A8C\u5355\u884C\u6570\u636E
     * @param rowIndex \u884C\u53F7
     * @param rowData  \u5F53\u524D\u884C\u6570\u636E\uFF08\u5217\u540D\u2192\u503C\uFF09
     * @param config   \u6A21\u677F\u914D\u7F6E
     * @return \u9519\u8BEF\u5217\u8868
     */
    List<RowError> validate(int rowIndex, Map<String, String> rowData,
                            ImportTemplateConfig config);
}`));

  children.push(heading("3.2 \u6821\u9A8C\u5668\u6CE8\u518C\u4E2D\u5FC3", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6821\u9A8C\u5668\u6CE8\u518C\u4E2D\u5FC3\u8D1F\u8D23\u7BA1\u7406\u6240\u6709\u5355\u5143\u683C\u7EA7\u548C\u884C\u7EA7\u6821\u9A8C\u5668\u7684\u5B9E\u4F8B\u3002\u5B83\u901A\u8FC7 Spring \u7684\u4F9D\u8D56\u6CE8\u5165\u81EA\u52A8\u6536\u96C6\u6240\u6709\u6821\u9A8C\u5668 Bean\uFF0C\u5E76\u6309\u7167\u6267\u884C\u987A\u5E8F\u6392\u5E8F\u3002\u5F53\u9700\u8981\u65B0\u589E\u6821\u9A8C\u89C4\u5219\u65F6\uFF0C\u53EA\u9700\u521B\u5EFA\u4E00\u4E2A\u65B0\u7684 Validator \u7C7B\u5E76\u52A0\u4E0A @Component \u6CE8\u89E3\uFF0C\u6846\u67B6\u4F1A\u81EA\u52A8\u53D1\u73B0\u5E76\u7EB3\u5165\u6821\u9A8C\u6D41\u7A0B\u3002"
  ));

  children.push(...codeBlock(`@Component
public class ValidatorRegistry {

    private final List<CellValidator> cellValidators;
    private final List<RowValidator> rowValidators;

    /**
     * Spring \u81EA\u52A8\u6CE8\u5165\u6240\u6709\u6821\u9A8C\u5668\u5B9E\u4F8B\uFF0C\u6309 @Order \u6392\u5E8F
     */
    public ValidatorRegistry(
            List<CellValidator> cellValidators,
            List<RowValidator> rowValidators) {
        this.cellValidators = cellValidators.stream()
            .sorted(Comparator.comparingInt(v ->
                v.getClass().getAnnotation(Order.class) != null
                ? v.getClass().getAnnotation(Order.class).value()
                : Integer.MAX_VALUE))
            .collect(Collectors.toList());
        this.rowValidators = rowValidators.stream()
            .sorted(Comparator.comparingInt(v ->
                v.getClass().getAnnotation(Order.class) != null
                ? v.getClass().getAnnotation(Order.class).value()
                : Integer.MAX_VALUE))
            .collect(Collectors.toList());
    }

    public List<CellValidator> getCellValidators() {
        return Collections.unmodifiableList(cellValidators);
    }

    public List<RowValidator> getRowValidators() {
        return Collections.unmodifiableList(rowValidators);
    }
}`));

  // ═══ Chapter 4: File-level Validators ═══
  children.push(heading("\u56DB\u3001\u6587\u4EF6\u7EA7\u963B\u65AD\u6821\u9A8C"));

  children.push(heading("4.1 \u6587\u4EF6\u683C\u5F0F\u6821\u9A8C", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6587\u4EF6\u683C\u5F0F\u6821\u9A8C\u662F\u6700\u524D\u7F6E\u7684\u963B\u65AD\u6821\u9A8C\u3002\u5B83\u68C0\u67E5\u4E0A\u4F20\u6587\u4EF6\u7684\u6269\u5C55\u540D\u548C MIME \u7C7B\u578B\uFF0C\u786E\u4FDD\u6587\u4EF6\u662F\u5408\u6CD5\u7684 Excel \u683C\u5F0F\u3002\u5982\u679C\u6587\u4EF6\u683C\u5F0F\u4E0D\u7B26\u5408\u8981\u6C42\uFF0C\u76F4\u63A5\u6807\u8BB0\u4E3A\u5168\u90E8\u5BFC\u5165\u5931\u8D25\u5E76\u8FD4\u56DE\uFF0C\u4E0D\u518D\u8FDB\u884C\u4EFB\u4F55\u540E\u7EED\u5904\u7406\u3002\u8FD9\u91CC\u540C\u65F6\u68C0\u67E5\u6587\u4EF6\u6269\u5C55\u540D\u548C Content-Type\uFF0C\u53EF\u4EE5\u6709\u6548\u9632\u6B62\u7528\u6237\u5C06\u975E Excel \u6587\u4EF6\uFF08\u5982 CSV\u3001\u56FE\u7247\u7B49\uFF09\u6539\u6269\u5C55\u540D\u540E\u4E0A\u4F20\u3002"
  ));

  children.push(...codeBlock(`/**
 * \u6587\u4EF6\u683C\u5F0F\u6821\u9A8C\u5668\uFF08\u975E\u7B56\u7565\u6A21\u5F0F\uFF0C\u4F5C\u4E3A\u72EC\u7ACB\u7684\u524D\u7F6E\u68C0\u67E5\uFF09
 */
public class FileFormatValidator {

    private static final Set<String> ALLOWED_EXTENSIONS =
        Set.of(".xls", ".xlsx");
    private static final Set<String> ALLOWED_CONTENT_TYPES =
        Set.of(
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument"
                + ".spreadsheetml.sheet"
        );

    /**
     * \u6821\u9A8C\u6587\u4EF6\u683C\u5F0F
     * @return \u9519\u8BEF\u4FE1\u606F\uFF0Cnull \u8868\u793A\u901A\u8FC7
     */
    public static RowError validate(MultipartFile file) {
        String originalFilename = file.getOriginalFilename();
        if (originalFilename == null) {
            return new RowError(0, "",
                ValidationErrorType.FILE_FORMAT,
                "\u65E0\u6CD5\u8BC6\u522B\u6587\u4EF6\u540D");
        }

        String extension = originalFilename.substring(
            originalFilename.lastIndexOf(".")).toLowerCase();
        if (!ALLOWED_EXTENSIONS.contains(extension)) {
            return new RowError(0, "",
                ValidationErrorType.FILE_FORMAT,
                "\u6587\u4EF6\u683C\u5F0F\u4E0D\u6B63\u786E\uFF0C\u4EC5\u652F\u6301 .xls \u548C .xlsx \u683C\u5F0F");
        }

        String contentType = file.getContentType();
        if (contentType != null
                && !ALLOWED_CONTENT_TYPES.contains(contentType)) {
            return new RowError(0, "",
                ValidationErrorType.FILE_FORMAT,
                "\u6587\u4EF6 Content-Type \u4E0D\u5339\u914D\uFF0C\u8BF7\u786E\u8BA4\u4E0A\u4F20\u7684\u662F Excel \u6587\u4EF6");
        }

        return null; // \u6821\u9A8C\u901A\u8FC7
    }
}`));

  children.push(heading("4.2 \u6A21\u677F\u5339\u914D\u6821\u9A8C", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6A21\u677F\u5339\u914D\u6821\u9A8C\u786E\u4FDD\u7528\u6237\u4E0A\u4F20\u7684 Excel \u8868\u5934\u4E0E\u9884\u8BBE\u6A21\u677F\u5B8C\u5168\u4E00\u81F4\u3002\u5B83\u901A\u8FC7\u8BFB\u53D6 Excel \u7684\u7B2C\u4E00\u884C\uFF08\u8868\u5934\u884C\uFF09\uFF0C\u4E0E\u914D\u7F6E\u4E2D\u7684 expectedHeaders \u8FDB\u884C\u9010\u5217\u6BD4\u5BF9\u3002\u5982\u679C\u8868\u5934\u4E0D\u5339\u914D\uFF0C\u8BF4\u660E\u7528\u6237\u4F7F\u7528\u7684\u4E0D\u662F\u6700\u65B0\u6A21\u677F\uFF0C\u6B64\u65F6\u5E94\u63D0\u793A\u7528\u6237\u4E0B\u8F7D\u6700\u65B0\u6A21\u677F\u5E76\u91CD\u65B0\u586B\u5199\u3002\u6B64\u6821\u9A8C\u4E5F\u662F\u963B\u65AD\u6027\u7684\uFF0C\u5931\u8D25\u540E\u4E0D\u518D\u8FDB\u884C\u884C\u7EA7\u6821\u9A8C\u3002"
  ));

  children.push(...codeBlock(`/**
 * \u6A21\u677F\u5339\u914D\u6821\u9A8C\u5668
 */
public class TemplateMatchValidator {

    /**
     * \u68C0\u67E5\u8868\u5934\u662F\u5426\u4E0E\u6A21\u677F\u914D\u7F6E\u4E00\u81F4
     * @param actualHeaders  Excel \u5B9E\u9645\u8868\u5934
     * @param config         \u6A21\u677F\u914D\u7F6E
     * @return \u9519\u8BEF\u4FE1\u606F\uFF0Cnull \u8868\u793A\u901A\u8FC7
     */
    public static RowError validate(
            List<String> actualHeaders,
            ImportTemplateConfig config) {

        List<String> expected = config.getExpectedHeaders();

        // \u5217\u6570\u4E0D\u540C
        if (actualHeaders.size() != expected.size()) {
            return new RowError(0, "",
                ValidationErrorType.TEMPLATE_MISMATCH,
                String.format("\u5BFC\u5165\u6A21\u677F\u4E0D\u5339\u914D\uFF1A\u671F\u671B %d \u5217\uFF0C\u5B9E\u9645 %d \u5217\uFF0C\u8BF7\u4E0B\u8F7D\u6700\u65B0\u6A21\u677F",
                    expected.size(), actualHeaders.size()));
        }

        // \u9010\u5217\u6BD4\u5BF9
        List<String> mismatches = new ArrayList<>();
        for (int i = 0; i < expected.size(); i++) {
            if (!expected.get(i).equals(actualHeaders.get(i))) {
                mismatches.add(String.format("\u7B2C%d\u5217\uFF1A\u671F\u671B[%s]\uFF0C\u5B9E\u9645[%s]",
                    i + 1, expected.get(i), actualHeaders.get(i)));
            }
        }

        if (!mismatches.isEmpty()) {
            return new RowError(0, "",
                ValidationErrorType.TEMPLATE_MISMATCH,
                "\u5BFC\u5165\u6A21\u677F\u4E0D\u5339\u914D\uFF0C\u8BF7\u4E0B\u8F7D\u6700\u65B0\u6A21\u677F\u3002\u8BE6\u60C5\uFF1A"
                    + String.join("\uFF1B", mismatches));
        }

        return null; // \u6821\u9A8C\u901A\u8FC7
    }
}`));

  // ═══ Chapter 5: Row-level Validators ═══
  children.push(heading("\u4E94\u3001\u884C\u7EA7\u6570\u636E\u6821\u9A8C\u5B9E\u73B0"));

  children.push(heading("5.1 \u5FC5\u586B\u6821\u9A8C\u5668", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u5FC5\u586B\u6821\u9A8C\u662F\u884C\u7EA7\u6821\u9A8C\u4E2D\u6700\u57FA\u7840\u4E14\u6700\u5148\u6267\u884C\u7684\u6821\u9A8C\u3002\u5F53\u67D0\u4E2A\u5FC5\u586B\u5B57\u6BB5\u4E3A\u7A7A\u65F6\uFF0C\u5E94\u8BE5\u7ACB\u5373\u8BB0\u5F55\u9519\u8BEF\u5E76\u8DF3\u8FC7\u8BE5\u5B57\u6BB5\u7684\u540E\u7EED\u6821\u9A8C\uFF08\u7C7B\u578B\u3001\u957F\u5EA6\u7B49\uFF09\uFF0C\u56E0\u4E3A\u5BF9\u7A7A\u503C\u505A\u7C7B\u578B\u6216\u957F\u5EA6\u6821\u9A8C\u6CA1\u6709\u610F\u4E49\uFF0C\u53CD\u800C\u4F1A\u4EA7\u751F\u5197\u4F59\u4E14\u8BA9\u7528\u6237\u56F0\u60D1\u7684\u63D0\u793A\u4FE1\u606F\u3002\u8FD9\u662F\u6821\u9A8C\u5668\u6392\u5E8F\u4E2D\u5C06\u5FC5\u586B\u6821\u9A8C\u7F6E\u4E8E\u6700\u524D\u7684\u6838\u5FC3\u539F\u56E0\u3002"
  ));

  children.push(...codeBlock(`@Component
@Order(1)  // \u6700\u5148\u6267\u884C
public class RequiredValidator implements CellValidator {

    @Override
    public List<RowError> validate(int rowIndex, String columnName,
                                   String cellValue, ColumnRule rule) {
        if (!rule.isRequired()) {
            return Collections.emptyList();
        }

        if (cellValue == null || cellValue.trim().isEmpty()) {
            String msg = rule.getErrorMsg() != null
                ? rule.getErrorMsg()
                : String.format("\u7B2C%d\u884C[%s]\u4E0D\u80FD\u4E3A\u7A7A",
                    rowIndex, columnName);
            return List.of(new RowError(
                rowIndex, columnName,
                ValidationErrorType.REQUIRED_EMPTY, msg));
        }

        return Collections.emptyList();
    }
}`));

  children.push(heading("5.2 \u6570\u636E\u7C7B\u578B\u6821\u9A8C\u5668", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6570\u636E\u7C7B\u578B\u6821\u9A8C\u5668\u68C0\u67E5\u5355\u5143\u683C\u7684\u503C\u662F\u5426\u80FD\u591F\u88AB\u89E3\u6790\u4E3A\u914D\u7F6E\u4E2D\u6307\u5B9A\u7684\u6570\u636E\u7C7B\u578B\u3002\u5B83\u652F\u6301\u56DB\u79CD\u57FA\u672C\u7C7B\u578B\uFF1ASTRING\uFF08\u5B57\u7B26\u4E32\uFF0C\u9ED8\u8BA4\u59CB\u7EC8\u901A\u8FC7\uFF09\u3001NUMBER\uFF08\u6570\u5B57\uFF0C\u5C1D\u8BD5\u8F6C\u6362\u4E3A Double\uFF09\u3001DATE\uFF08\u65E5\u671F\uFF0C\u5C1D\u8BD5\u89E3\u6790\u4E3A LocalDate\uFF09\u548C ENUM\uFF08\u679A\u4E3E\uFF0C\u68C0\u67E5\u662F\u5426\u5728\u5141\u8BB8\u503C\u5217\u8868\u4E2D\uFF09\u3002\u8FD9\u91CC\u7279\u522B\u6CE8\u610F\u7684\u662F\uFF0CEasyExcel \u4F7F\u7528 String \u7C7B\u578B\u7EDF\u4E00\u8BFB\u53D6\u6240\u6709\u5355\u5143\u683C\uFF0C\u907F\u514D\u4E86 Excel \u5C06\u624B\u673A\u53F7\u8F6C\u4E3A\u79D1\u5B66\u8BA1\u6570\u6CD5\u6216\u5C06\u65E5\u671F\u8F6C\u4E3A\u6D6E\u70B9\u6570\u7684\u95EE\u9898\u3002"
  ));

  children.push(...codeBlock(`@Component
@Order(2)
public class DataTypeValidator implements CellValidator {

    private static final DateTimeFormatter[] DATE_FORMATS = {
        DateTimeFormatter.ofPattern("yyyy-MM-dd"),
        DateTimeFormatter.ofPattern("yyyy/MM/dd"),
        DateTimeFormatter.ofPattern("yyyyMMdd"),
    };

    @Override
    public List<RowError> validate(int rowIndex, String columnName,
                                   String cellValue, ColumnRule rule) {
        if (rule.getDataType() == null || cellValue == null
                || cellValue.trim().isEmpty()) {
            return Collections.emptyList();
        }

        String trimmed = cellValue.trim();
        boolean valid = switch (rule.getDataType()) {
            case STRING -> true;  // \u5B57\u7B26\u4E32\u7C7B\u578B\u59CB\u7EC8\u5408\u6CD5
            case NUMBER -> isNumber(trimmed);
            case DATE   -> isDate(trimmed);
            case ENUM   -> isValidEnum(trimmed, rule.getEnumValues());
        };

        if (!valid) {
            String msg = String.format(
                "\u7B2C%d\u884C[%s]\u6570\u636E\u7C7B\u578B\u9519\u8BEF\uFF0C\u5E94\u4E3A[%s]",
                rowIndex, columnName, rule.getDataType().name());
            return List.of(new RowError(
                rowIndex, columnName,
                ValidationErrorType.TYPE_MISMATCH, msg));
        }

        return Collections.emptyList();
    }

    private boolean isNumber(String value) {
        try {
            Double.parseDouble(value);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    private boolean isDate(String value) {
        for (DateTimeFormatter fmt : DATE_FORMATS) {
            try {
                LocalDate.parse(value, fmt);
                return true;
            } catch (DateTimeParseException ignored) {}
        }
        return false;
    }

    private boolean isValidEnum(String value, List<String> enumValues) {
        return enumValues != null && enumValues.contains(value);
    }
}`));

  children.push(heading("5.3 \u6570\u636E\u957F\u5EA6\u6821\u9A8C\u5668", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6570\u636E\u957F\u5EA6\u6821\u9A8C\u5668\u68C0\u67E5\u5355\u5143\u683C\u503C\u7684\u5B57\u7B26\u957F\u5EA6\u662F\u5426\u5728\u914D\u7F6E\u7684\u5141\u8BB8\u8303\u56F4\u5185\u3002\u5B83\u652F\u6301\u6700\u5C0F\u957F\u5EA6\u548C\u6700\u5927\u957F\u5EA6\u7684\u53CC\u5411\u7EA6\u675F\uFF0C\u4F8B\u5982\u624B\u673A\u53F7\u5FC5\u987B\u7B49\u4E8E11\u4F4D\uFF0C\u59D3\u540D\u4E0D\u80FD\u8D85\u8FC710\u4E2A\u5B57\u7B26\u7B49\u3002\u957F\u5EA6\u6821\u9A8C\u5E94\u5728\u7C7B\u578B\u6821\u9A8C\u901A\u8FC7\u4E4B\u540E\u6267\u884C\uFF0C\u56E0\u4E3A\u5982\u679C\u7C7B\u578B\u5DF2\u7ECF\u4E0D\u5339\u914D\uFF0C\u957F\u5EA6\u6821\u9A8C\u7684\u610F\u4E49\u4E5F\u4E0D\u5927\u3002"
  ));

  children.push(...codeBlock(`@Component
@Order(3)
public class LengthValidator implements CellValidator {

    @Override
    public List<RowError> validate(int rowIndex, String columnName,
                                   String cellValue, ColumnRule rule) {
        if (cellValue == null || cellValue.trim().isEmpty()) {
            return Collections.emptyList();
        }

        List<RowError> errors = new ArrayList<>();
        int len = cellValue.trim().length();

        if (rule.getMaxLength() != null && len > rule.getMaxLength()) {
            errors.add(new RowError(rowIndex, columnName,
                ValidationErrorType.LENGTH_EXCEEDED,
                String.format("\u7B2C%d\u884C[%s]\u957F\u5EA6\u4E0D\u80FD\u8D85\u8FC7%d\u4E2A\u5B57\u7B26\uFF0C\u5F53\u524D%d\u4E2A",
                    rowIndex, columnName, rule.getMaxLength(), len)));
        }

        if (rule.getMinLength() != null && len < rule.getMinLength()) {
            errors.add(new RowError(rowIndex, columnName,
                ValidationErrorType.LENGTH_EXCEEDED,
                String.format("\u7B2C%d\u884C[%s]\u957F\u5EA6\u4E0D\u80FD\u5C11\u4E8E%d\u4E2A\u5B57\u7B26\uFF0C\u5F53\u524D%d\u4E2A",
                    rowIndex, columnName, rule.getMinLength(), len)));
        }

        return errors;
    }
}`));

  children.push(heading("5.4 \u6B63\u5219\u6821\u9A8C\u5668", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6B63\u5219\u6821\u9A8C\u5668\u652F\u6301\u901A\u8FC7\u6B63\u5219\u8868\u8FBE\u5F0F\u5BF9\u5355\u5143\u683C\u503C\u8FDB\u884C\u7CBE\u7EC6\u5316\u7684\u683C\u5F0F\u6821\u9A8C\uFF0C\u4F8B\u5982\u624B\u673A\u53F7\u683C\u5F0F\u3001\u8EAB\u4EFD\u8BC1\u53F7\u7801\u3001\u90AE\u7BB1\u5730\u5740\u7B49\u3002\u5B83\u662F\u7C7B\u578B\u6821\u9A8C\u7684\u6709\u529B\u8865\u5145\uFF1A\u7C7B\u578B\u6821\u9A8C\u53EA\u80FD\u786E\u8BA4\u201C\u8FD9\u662F\u4E00\u4E2A\u6570\u5B57\u201D\uFF0C\u4F46\u6B63\u5219\u6821\u9A8C\u53EF\u4EE5\u786E\u8BA4\u201C\u8FD9\u662F\u4E00\u4E2A\u7B26\u5408\u89C4\u5219\u7684\u624B\u673A\u53F7\u201D\u3002\u914D\u7F6E\u4E2D\u7684 regexPattern \u5B57\u6BB5\u652F\u6301\u4E3A\u4E0D\u540C\u5217\u5B9A\u4E49\u4E0D\u540C\u7684\u6B63\u5219\u89C4\u5219\uFF0C\u6846\u67B6\u4F1A\u81EA\u52A8\u7F16\u8BD1\u5E76\u7F13\u5B58\u6B63\u5219\u8868\u8FBE\u5F0F\u5BF9\u8C61\u4EE5\u63D0\u5347\u6027\u80FD\u3002"
  ));

  children.push(...codeBlock(`@Component
@Order(4)
public class RegexValidator implements CellValidator {

    /** \u6B63\u5219\u8868\u8FBE\u5F0F\u7F13\u5B58\uFF0C\u907F\u514D\u91CD\u590D\u7F16\u8BD1 */
    private final ConcurrentHashMap<String, Pattern> patternCache
        = new ConcurrentHashMap<>();

    @Override
    public List<RowError> validate(int rowIndex, String columnName,
                                   String cellValue, ColumnRule rule) {
        if (rule.getRegexPattern() == null || cellValue == null
                || cellValue.trim().isEmpty()) {
            return Collections.emptyList();
        }

        Pattern pattern = patternCache.computeIfAbsent(
            rule.getRegexPattern(), Pattern::compile);

        if (!pattern.matcher(cellValue.trim()).matches()) {
            return List.of(new RowError(
                rowIndex, columnName,
                ValidationErrorType.REGEX_MISMATCH,
                String.format("\u7B2C%d\u884C[%s]\u683C\u5F0F\u4E0D\u6B63\u786E",
                    rowIndex, columnName)));
        }

        return Collections.emptyList();
    }
}`));

  children.push(heading("5.5 \u6587\u4EF6\u5185\u91CD\u590D\u6821\u9A8C\u5668", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6587\u4EF6\u5185\u91CD\u590D\u6821\u9A8C\u5668\u662F\u4E00\u4E2A\u884C\u7EA7\u6821\u9A8C\u5668\uFF0C\u5B83\u4E0D\u662F\u5355\u72EC\u68C0\u67E5\u67D0\u4E2A\u5355\u5143\u683C\uFF0C\u800C\u662F\u5728\u904D\u5386\u8FC7\u7A0B\u4E2D\u8DDF\u8E2A\u5386\u53F2\u503C\uFF0C\u68C0\u6D4B\u5F53\u524D\u884C\u7684\u67D0\u4E9B\u5B57\u6BB5\u7EC4\u5408\u662F\u5426\u4E0E\u5DF2\u51FA\u73B0\u7684\u884C\u91CD\u590D\u3002\u6838\u5FC3\u8BBE\u8BA1\u662F\u4F7F\u7528 HashMap \u6765\u8BB0\u5F55\u5DF2\u7ECF\u51FA\u73B0\u8FC7\u7684\u503C\u53CA\u5176\u884C\u53F7\uFF0C\u67E5\u627E\u65F6\u95F4\u590D\u6742\u5EA6\u4E3A O(1)\uFF0C\u6BD4\u6BCF\u6B21\u904D\u5386\u53BB\u91CD\u6548\u7387\u9AD8\u5F97\u591A\u3002\u914D\u7F6E\u4E2D\u7684 uniqueKeyGroups \u652F\u6301\u591A\u5B57\u6BB5\u8054\u5408\u552F\u4E00\u7EA6\u675F\uFF0C\u4F8B\u5982\u201C\u59D3\u540D+\u624B\u673A\u53F7\u201D\u7684\u7EC4\u5408\u552F\u4E00\u3002"
  ));

  children.push(...codeBlock(`@Component
@Order(1)
public class InFileDuplicateValidator implements RowValidator {

    /** \u6BCF\u4E2A\u552F\u4E00\u952E\u7EC4\u5408\u7684\u5386\u53F2\u503C\uFF1Akey\u2192\u9996\u6B21\u51FA\u73B0\u7684\u884C\u53F7 */
    private final Map<String, Map<String, Integer>> historyMap
        = new HashMap<>();

    @Override
    public List<RowError> validate(int rowIndex,
                                   Map<String, String> rowData,
                                   ImportTemplateConfig config) {
        List<RowError> errors = new ArrayList<>();

        for (List<String> keyGroup : config.getUniqueKeyGroups()) {
            // \u62FC\u63A5\u8054\u5408\u552F\u4E00\u952E\uFF1A\u5982 "\u5F20\u4E09|13800138000"
            String compositeKey = keyGroup.stream()
                .map(col -> rowData.getOrDefault(col, ""))
                .collect(Collectors.joining("|"));

            String groupKey = String.join(",", keyGroup);
            Map<String, Integer> groupHistory =
                historyMap.computeIfAbsent(groupKey, k -> new HashMap<>());

            if (groupHistory.containsKey(compositeKey)) {
                int firstRow = groupHistory.get(compositeKey);
                errors.add(new RowError(
                    rowIndex, String.join("+", keyGroup),
                    ValidationErrorType.DUPLICATE_IN_FILE,
                    String.format(
                        "\u7B2C%d\u884C\u4E0E\u7B2C%d\u884C\u7684[%s]\u91CD\u590D",
                        rowIndex, firstRow,
                        String.join("+", keyGroup))));
            } else {
                groupHistory.put(compositeKey, rowIndex);
            }
        }

        return errors;
    }

    /** \u6BCF\u6B21\u5BFC\u5165\u524D\u8C03\u7528\uFF0C\u6E05\u7A7A\u5386\u53F2\u8BB0\u5F55 */
    public void reset() {
        historyMap.clear();
    }
}`));

  children.push(heading("5.6 \u6570\u636E\u5E93\u53BB\u91CD\u6821\u9A8C\u5668", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u6570\u636E\u5E93\u53BB\u91CD\u6821\u9A8C\u662F\u4E0E\u5916\u90E8\u7CFB\u7EDF\u4EA4\u4E92\u7684\u6821\u9A8C\u5668\u3002\u5173\u952E\u8BBE\u8BA1\u539F\u5219\u662F\uFF1A\u7EDD\u4E0D\u80FD\u9010\u884C\u53BB\u67E5\u6570\u636E\u5E93\uFF0C\u5426\u5219\u4E00\u4E2A\u4E07\u884C\u7684 Excel \u4F1A\u4EA7\u751F\u4E07\u6B21\u6570\u636E\u5E93\u67E5\u8BE2\uFF0C\u6027\u80FD\u707E\u96BE\u6027\u53EF\u60F3\u800C\u77E5\u3002\u6B63\u786E\u7684\u505A\u6CD5\u662F\uFF1A\u5728\u6240\u6709\u884C\u7EA7\u6821\u9A8C\u5B8C\u6210\u540E\uFF0C\u5C06\u6821\u9A8C\u901A\u8FC7\u7684\u6570\u636E\u4E2D\u7684\u53BB\u91CD\u5B57\u6BB5\u6536\u96C6\u4E3A\u4E00\u4E2A List\uFF0C\u4E00\u6B21\u6027\u901A\u8FC7 IN \u67E5\u8BE2\u6279\u91CF\u83B7\u53D6\u6570\u636E\u5E93\u4E2D\u5DF2\u5B58\u5728\u7684\u8BB0\u5F55\uFF0C\u518D\u4E0E\u5F53\u524D\u6570\u636E\u6BD4\u5BF9\u3002\u8FD9\u6837\u65E0\u8BBA\u6587\u4EF6\u6709\u591A\u5C11\u884C\uFF0C\u6570\u636E\u5E93\u67E5\u8BE2\u6B21\u6570\u90FD\u53EA\u4E0E\u552F\u4E00\u952E\u7EC4\u7684\u6570\u91CF\u76F8\u5173\uFF0C\u800C\u975E\u884C\u6570\u3002"
  ));

  children.push(...codeBlock(`@Component
public class DbDuplicateValidator {

    /** \u6570\u636E\u5E93\u53BB\u91CD\u67E5\u8BE2\u63A5\u53E3\uFF08\u7531\u4E1A\u52A1\u5C42\u5B9E\u73B0\uFF09 */
    public interface DuplicateChecker {
        /**
         * \u6279\u91CF\u67E5\u8BE2\u6570\u636E\u5E93\u4E2D\u5DF2\u5B58\u5728\u7684\u503C
         * @param columnNames  \u53BB\u91CD\u5B57\u6BB5\u540D\u5217\u8868
         * @param values       \u5F85\u67E5\u8BE2\u7684\u503C\u5217\u8868
         * @return \u5DF2\u5B58\u5728\u4E8E\u6570\u636E\u5E93\u7684\u503C\u96C6\u5408
         */
        Set<String> findExistingValues(
            List<String> columnNames, List<String> values);
    }

    private final DuplicateChecker duplicateChecker;

    public DbDuplicateValidator(DuplicateChecker duplicateChecker) {
        this.duplicateChecker = duplicateChecker;
    }

    /**
     * \u5BF9\u6821\u9A8C\u901A\u8FC7\u7684\u6570\u636E\u8FDB\u884C\u6570\u636E\u5E93\u53BB\u91CD\u6821\u9A8C
     * @param validRows  \u5DF2\u901A\u8FC7\u5176\u4ED6\u6821\u9A8C\u7684\u884C\u6570\u636E
     * @param config     \u6A21\u677F\u914D\u7F6E
     * @return \u6570\u636E\u5E93\u91CD\u590D\u7684\u9519\u8BEF\u5217\u8868
     */
    public List<RowError> validate(
            List<Map<String, String>> validRows,
            ImportTemplateConfig config) {

        List<RowError> errors = new ArrayList<>();

        for (List<String> keyGroup : config.getUniqueKeyGroups()) {
            // 1. \u6536\u96C6\u6240\u6709\u884C\u7684\u8054\u5408\u952E\u503C
            List<String> compositeKeys = validRows.stream()
                .map(row -> keyGroup.stream()
                    .map(col -> row.getOrDefault(col, ""))
                    .collect(Collectors.joining("|")))
                .collect(Collectors.toList());

            // 2. \u4E00\u6B21\u6027\u67E5\u8BE2\u6570\u636E\u5E93
            Set<String> existingValues =
                duplicateChecker.findExistingValues(keyGroup, compositeKeys);

            // 3. \u6807\u8BB0\u91CD\u590D\u884C
            for (int i = 0; i < compositeKeys.size(); i++) {
                if (existingValues.contains(compositeKeys.get(i))) {
                    errors.add(new RowError(
                        i + 2, // Excel \u884C\u53F7
                        String.join("+", keyGroup),
                        ValidationErrorType.DUPLICATE_IN_DB,
                        String.format("\u7B2C%d\u884C\u4E0E\u7CFB\u7EDF\u5DF2\u6709\u6570\u636E\u91CD\u590D",
                            i + 2)));
                }
            }
        }

        return errors;
    }
}`));

  // ═══ Chapter 6: EasyExcel Integration ═══
  children.push(heading("\u516D\u3001EasyExcel \u96C6\u6210\u4E0E\u8C03\u5EA6"));

  children.push(heading("6.1 ReadListener \u5B9E\u73B0", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "EasyExcel \u7684\u6838\u5FC3\u662F ReadListener \u673A\u5236\uFF0C\u5B83\u91C7\u7528 SAX \u6A21\u5F0F\u9010\u884C\u89E3\u6790 Excel\uFF0C\u4E0D\u4F1A\u5C06\u6574\u4E2A\u6587\u4EF6\u52A0\u8F7D\u5230\u5185\u5B58\u4E2D\uFF0C\u4ECE\u800C\u907F\u514D\u4E86\u5927\u6587\u4EF6\u5BFC\u81F4\u7684 OOM \u95EE\u9898\u3002\u6211\u4EEC\u7684\u8C03\u5EA6\u5668\u5C06\u6574\u4E2A\u6821\u9A8C\u6D41\u7A0B\u5C01\u88C5\u5728 ReadListener \u5185\u90E8\uFF0C\u6BCF\u8BFB\u53D6\u4E00\u884C\u6570\u636E\u5C31\u7ACB\u5373\u89E6\u53D1\u5355\u5143\u683C\u7EA7\u548C\u884C\u7EA7\u6821\u9A8C\u3002doAfterAllAnalysed \u65B9\u6CD5\u4E2D\u6267\u884C\u6570\u636E\u5E93\u53BB\u91CD\u6821\u9A8C\uFF0C\u56E0\u4E3A\u5B83\u9700\u8981\u5728\u6240\u6709\u884C\u8BFB\u53D6\u5B8C\u6BD5\u540E\u624D\u80FD\u6279\u91CF\u67E5\u8BE2\u3002"
  ));

  children.push(...codeBlock(`/**
 * \u5BFC\u5165\u6821\u9A8C\u8C03\u5EA6\u5668\uFF08\u6838\u5FC3\u7EC4\u4EF6\uFF09
 */
public class ImportValidationListener
        implements ReadListener<Map<Integer, String>> {

    private final ImportTemplateConfig config;
    private final ValidatorRegistry registry;
    private final InFileDuplicateValidator duplicateValidator;
    private final ImportResult.ImportResultBuilder resultBuilder;
    private final List<Map<String, String>> validRows;
    private final List<RowError> allErrors;

    public ImportValidationListener(ImportTemplateConfig config,
                                    ValidatorRegistry registry) {
        this.config = config;
        this.registry = registry;
        this.duplicateValidator = new InFileDuplicateValidator();
        this.resultBuilder = ImportResult.builder();
        this.validRows = new ArrayList<>();
        this.allErrors = new ArrayList<>();
    }

    @Override
    public void invoke(Map<Integer, String> rowData,
                       AnalysisContext context) {
        int rowIndex = context.readRowHolder().getRowIndex() + 1;
        // +1 \u56E0\u4E3A Excel \u884C\u53F7\u4ECE 1 \u5F00\u59CB

        // \u5C06\u884C\u7D22\u5F15\u8F6C\u4E3A \u5217\u540D\u2192\u503C \u7684 Map
        Map<String, String> namedRow = new LinkedHashMap<>();
        List<String> headers = config.getExpectedHeaders();
        for (int i = 0; i < headers.size(); i++) {
            namedRow.put(headers.get(i),
                rowData.getOrDefault(i, ""));
        }

        // \u904D\u5386\u6BCF\u5217\u6267\u884C\u5355\u5143\u683C\u7EA7\u6821\u9A8C
        List<RowError> rowErrors = new ArrayList<>();
        for (Map.Entry<String, String> entry : namedRow.entrySet()) {
            String colName = entry.getKey();
            String cellValue = entry.getValue();
            ColumnRule rule = config.getColumnRules().get(colName);

            if (rule == null) continue;

            for (CellValidator validator : registry.getCellValidators()) {
                List<RowError> errors = validator.validate(
                    rowIndex, colName, cellValue, rule);
                rowErrors.addAll(errors);

                // \u5FC5\u586B\u6821\u9A8C\u5931\u8D25\u65F6\u8DF3\u8FC7\u540E\u7EED\u6821\u9A8C
                if (!errors.isEmpty()
                    && errors.get(0).getErrorType()
                        == ValidationErrorType.REQUIRED_EMPTY) {
                    break;
                }
            }
        }

        // \u6267\u884C\u884C\u7EA7\u6821\u9A8C\uFF08\u5982\u6587\u4EF6\u5185\u53BB\u91CD\uFF09
        for (RowValidator rv : registry.getRowValidators()) {
            rowErrors.addAll(rv.validate(rowIndex, namedRow, config));
        }

        if (rowErrors.isEmpty()) {
            validRows.add(namedRow);
        } else {
            allErrors.addAll(rowErrors);
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // \u6B64\u5904\u53EF\u63D2\u5165\u6570\u636E\u5E93\u53BB\u91CD\u6821\u9A8C
        // DbDuplicateValidator dbValidator = ...
        // List<RowError> dbErrors = dbValidator.validate(validRows, config);
        // allErrors.addAll(dbErrors);
        // \u4ECE validRows \u4E2D\u79FB\u9664\u6570\u636E\u5E93\u91CD\u590D\u884C

        resultBuilder
            .totalFail(false)
            .successCount(validRows.size())
            .failCount(allErrors.size())
            .errors(allErrors)
            .validRows(validRows);
    }

    public ImportResult getResult() {
        return resultBuilder.build();
    }
}`));

  children.push(heading("6.2 \u5BFC\u5165\u670D\u52A1\u4E3B\u5165\u53E3", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u5BFC\u5165\u670D\u52A1\u662F\u6574\u4E2A\u6846\u67B6\u7684\u4E3B\u5165\u53E3\uFF0C\u5B83\u7F16\u6392\u4E86\u6587\u4EF6\u7EA7\u963B\u65AD\u6821\u9A8C\u548C\u884C\u7EA7\u904D\u5386\u6821\u9A8C\u7684\u6267\u884C\u987A\u5E8F\u3002\u5728\u6587\u4EF6\u7EA7\u6821\u9A8C\u901A\u8FC7\u540E\uFF0C\u5B83\u4F7F\u7528 EasyExcel \u7684 read \u65B9\u6CD5\u542F\u52A8 SAX \u6A21\u5F0F\u8BFB\u53D6\uFF0C\u5E76\u4F20\u5165\u81EA\u5B9A\u4E49\u7684 ReadListener\u3002\u6240\u6709\u6821\u9A8C\u903B\u8F91\u5747\u5728 Listener \u4E2D\u5B8C\u6210\uFF0C\u670D\u52A1\u5C42\u53EA\u8D1F\u8D23\u7F16\u6392\u548C\u8FD4\u56DE\u7ED3\u679C\u3002\u8FD9\u79CD\u8BBE\u8BA1\u4F7F\u5F97\u6821\u9A8C\u903B\u8F91\u4E0E\u4E1A\u52A1\u903B\u8F91\u5B8C\u5168\u89E3\u8026\u3002"
  ));

  children.push(...codeBlock(`@Service
public class ExcelImportService {

    @Autowired
    private ValidatorRegistry validatorRegistry;

    @Autowired
    private ImportTemplateConfigService configService;

    /**
     * \u6267\u884C\u5BFC\u5165\u6821\u9A8C
     * @param file         \u4E0A\u4F20\u7684\u6587\u4EF6
     * @param templateCode \u6A21\u677F\u7F16\u7801
     * @return \u6821\u9A8C\u7ED3\u679C
     */
    public ImportResult validateImport(
            MultipartFile file, String templateCode) {

        // \u2500\u2500 \u9636\u6BB51\uFF1A\u6587\u4EF6\u683C\u5F0F\u6821\u9A8C\uFF08\u963B\u65AD\u6027\uFF09 \u2500\u2500
        RowError formatError = FileFormatValidator.validate(file);
        if (formatError != null) {
            return ImportResult.builder()
                .totalFail(true)
                .errors(List.of(formatError))
                .successCount(0)
                .failCount(0)
                .build();
        }

        // \u52A0\u8F7D\u6A21\u677F\u914D\u7F6E
        ImportTemplateConfig config =
            configService.getConfig(templateCode);

        // \u2500\u2500 \u9636\u6BB52\uFF1A\u6A21\u677F\u5339\u914D\u6821\u9A8C\uFF08\u963B\u65AD\u6027\uFF09 \u2500\u2500
        List<String> actualHeaders = readHeaders(file);
        RowError templateError =
            TemplateMatchValidator.validate(actualHeaders, config);
        if (templateError != null) {
            return ImportResult.builder()
                .totalFail(true)
                .errors(List.of(templateError))
                .successCount(0)
                .failCount(0)
                .build();
        }

        // \u2500\u2500 \u9636\u6BB53\uFF1A\u884C\u7EA7\u904D\u5386\u6821\u9A8C \u2500\u2500
        ImportValidationListener listener =
            new ImportValidationListener(config, validatorRegistry);

        EasyExcel.read(file.getInputStream())
            .sheet()
            .headRowNumber(1)
            .registerReadListener(listener)
            .doRead();

        return listener.getResult();
    }

    /**
     * \u8BFB\u53D6 Excel \u8868\u5934\uFF08\u4EC5\u8BFB\u7B2C\u4E00\u884C\uFF09
     */
    private List<String> readHeaders(MultipartFile file) {
        List<String> headers = new ArrayList<>();
        EasyExcel.read(file.getInputStream())
            .sheet()
            .headRowNumber(0)
            .registerReadListener(new ReadListener<Map<Integer, String>>() {
                @Override
                public void invoke(Map<Integer, String> data,
                                   AnalysisContext ctx) {
                    // \u4EC5\u8BFB\u7B2C\u4E00\u884C
                }
                @Override
                public void doAfterAllAnalysed(
                        AnalysisContext ctx) {}
            })
            .doRead();

        // \u66F4\u7B80\u6D01\u7684\u65B9\u5F0F\uFF1A\u4F7F\u7528 head() \u65B9\u6CD5
        try (ExcelReader reader = EasyExcel.read(
                file.getInputStream()).build()) {
            ReadSheet sheet = EasyExcel.readSheet(0)
                .headRowNumber(0).build();
            List<Map<Integer, String>> headData =
                reader.read(sheet);
            if (!headData.isEmpty()) {
                Map<Integer, String> headerRow = headData.get(0);
                int maxCol = headerRow.keySet().stream()
                    .max(Integer::compare).orElse(-1);
                for (int i = 0; i <= maxCol; i++) {
                    headers.add(headerRow.getOrDefault(i, ""));
                }
            }
        }
        return headers;
    }
}`));

  // ═══ Chapter 7: Usage Example ═══
  children.push(heading("\u4E03\u3001\u4F7F\u7528\u793A\u4F8B"));

  children.push(heading("7.1 \u6A21\u677F\u914D\u7F6E\u793A\u4F8B", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u4EE5\u4E0B\u662F\u4E00\u4E2A\u5458\u5DE5\u4FE1\u606F\u5BFC\u5165\u7684\u6A21\u677F\u914D\u7F6E\u793A\u4F8B\u3002\u5B83\u5B9A\u4E49\u4E86\u56DB\u4E2A\u5B57\u6BB5\u7684\u6821\u9A8C\u89C4\u5219\u548C\u624B\u673A\u53F7\u7684\u552F\u4E00\u6027\u7EA6\u675F\u3002\u5176\u4E2D\u624B\u673A\u53F7\u540C\u65F6\u5177\u5907\u4E86\u6B63\u5219\u6821\u9A8C\u548C\u957F\u5EA6\u6821\u9A8C\uFF0C\u59D3\u540D\u548C\u5E74\u9F84\u4E3A\u5FC5\u586B\u5B57\u6BB5\uFF0C\u5165\u804C\u65E5\u671F\u4E3A\u975E\u5FC5\u586B\u7684\u65E5\u671F\u7C7B\u578B\u5B57\u6BB5\u3002\u8FD9\u79CD\u914D\u7F6E\u5316\u7684\u65B9\u5F0F\u4F7F\u5F97\u65B0\u589E\u4E00\u4E2A\u5BFC\u5165\u573A\u666F\u53EA\u9700\u589E\u52A0\u4E00\u5957\u914D\u7F6E\uFF0C\u65E0\u9700\u7F16\u5199\u4EFB\u4F55\u4EE3\u7801\u3002"
  ));

  children.push(...codeBlock(`@Configuration
public class EmployeeImportConfig {

    @Bean("employeeImportConfig")
    public ImportTemplateConfig employeeConfig() {
        Map<String, ColumnRule> rules = new LinkedHashMap<>();
        rules.put("\u59D3\u540D", ColumnRule.builder()
            .columnName("\u59D3\u540D")
            .required(true)
            .dataType(DataType.STRING)
            .maxLength(10)
            .build());
        rules.put("\u5E74\u9F84", ColumnRule.builder()
            .columnName("\u5E74\u9F84")
            .required(true)
            .dataType(DataType.NUMBER)
            .maxLength(3)
            .build());
        rules.put("\u624B\u673A\u53F7", ColumnRule.builder()
            .columnName("\u624B\u673A\u53F7")
            .required(true)
            .dataType(DataType.STRING)
            .minLength(11)
            .maxLength(11)
            .regexPattern("^1[3-9]\\\\d{9}$")
            .build());
        rules.put("\u5165\u804C\u65E5\u671F", ColumnRule.builder()
            .columnName("\u5165\u804C\u65E5\u671F")
            .required(false)
            .dataType(DataType.DATE)
            .build());

        return ImportTemplateConfig.builder()
            .templateCode("EMPLOYEE_IMPORT")
            .expectedHeaders(
                List.of("\u59D3\u540D", "\u5E74\u9F84", "\u624B\u673A\u53F7", "\u5165\u804C\u65E5\u671F"))
            .columnRules(rules)
            .uniqueKeyGroups(List.of(List.of("\u624B\u673A\u53F7")))
            .maxRows(10000)
            .build();
    }
}`));

  children.push(heading("7.2 Controller \u8C03\u7528\u793A\u4F8B", HeadingLevel.HEADING_2));

  children.push(...codeBlock(`@RestController
@RequestMapping("/api/import")
public class ImportController {

    @Autowired
    private ExcelImportService importService;

    @PostMapping("/employee")
    public ResponseEntity<ImportResult> importEmployee(
            @RequestParam("file") MultipartFile file) {

        ImportResult result =
            importService.validateImport(file, "EMPLOYEE_IMPORT");

        if (result.isTotalFail()) {
            return ResponseEntity.badRequest().body(result);
        }

        if (!result.getErrors().isEmpty()) {
            // \u90E8\u5206\u5931\u8D25\uFF0C\u8FD4\u56DE\u9519\u8BEF\u660E\u7EC6
            return ResponseEntity.status(207).body(result);
        }

        // \u5168\u90E8\u901A\u8FC7\uFF0C\u6267\u884C\u5199\u5165\u6570\u636E\u5E93
        // employeeService.batchInsert(result.getValidRows());
        return ResponseEntity.ok(result);
    }
}`));

  // ═══ Chapter 8: Production Considerations ═══
  children.push(heading("\u516B\u3001\u751F\u4EA7\u73AF\u5883\u8FDB\u9636\u5EFA\u8BAE"));

  children.push(heading("8.1 OOM \u9632\u62A4", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u5F53 Excel \u6587\u4EF6\u8D85\u8FC7 10 \u4E07\u884C\u65F6\uFF0C\u5373\u4F7F\u4F7F\u7528 EasyExcel \u7684 SAX \u6A21\u5F0F\u9010\u884C\u89E3\u6790\uFF0C\u4ECD\u7136\u9700\u8981\u6CE8\u610F\u5185\u5B58\u7BA1\u7406\u3002\u5173\u952E\u7B56\u7565\u5305\u62EC\uFF1A\u8BBE\u7F6E\u6700\u5927\u884C\u6570\u9650\u5236\uFF08\u5728 ImportTemplateConfig \u7684 maxRows \u4E2D\u914D\u7F6E\uFF09\uFF0C\u5F53\u8BFB\u53D6\u884C\u6570\u8D85\u8FC7\u9650\u5236\u65F6\u7ACB\u5373\u4E2D\u65AD\u89E3\u6790\u5E76\u8FD4\u56DE\u9519\u8BEF\u63D0\u793A\uFF1B\u5BF9\u4E8E\u6821\u9A8C\u901A\u8FC7\u7684 validRows \u5217\u8868\uFF0C\u91C7\u7528\u5206\u6279\u5199\u5165\u6570\u636E\u5E93\u7684\u65B9\u5F0F\uFF0C\u6BCF\u6279\u5904\u7406 500\u20131000 \u884C\u540E\u6E05\u7A7A\u5217\u8868\uFF0C\u907F\u514D\u5185\u5B58\u4E2D\u79EF\u538B\u8FC7\u591A\u6570\u636E\u3002"
  ));

  children.push(...codeBlock(`// \u5728 ReadListener.invoke() \u4E2D\u589E\u52A0\u884C\u6570\u9650\u5236\u68C0\u67E5
@Override
public void invoke(Map<Integer, String> rowData,
                   AnalysisContext context) {
    int currentRow = context.readRowHolder().getRowIndex();
    if (currentRow > config.getMaxRows()) {
        throw new ExcelAnalysisStopException(
            String.format("\u6587\u4EF6\u884C\u6570\u8D85\u8FC7\u9650\u5236\uFF08\u6700\u5927%d\u884C\uFF09",
                config.getMaxRows()));
    }
    // ... \u6B63\u5E38\u6821\u9A8C\u903B\u8F91
}

// \u5206\u6279\u5199\u5165\u6570\u636E\u5E93
public void batchInsert(List<Map<String, String>> validRows) {
    int batchSize = 500;
    List<List<Map<String, String>>> partitions =
        Lists.partition(validRows, batchSize);
    for (List<Map<String, String>> batch : partitions) {
        mapper.batchInsert(batch);
    }
}`));

  children.push(heading("8.2 \u5F02\u6B65\u5BFC\u5165\u4E0E\u8FDB\u5EA6\u63A8\u9001", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u5BF9\u4E8E\u5927\u6587\u4EF6\u5BFC\u5165\u573A\u666F\uFF0C\u524D\u7AEF\u5E94\u91C7\u7528\u5F02\u6B65\u5BFC\u5165\u673A\u5236\u3002\u540E\u7AEF\u63A5\u6536\u5230\u6587\u4EF6\u540E\u7ACB\u5373\u8FD4\u56DE\u4E00\u4E2A\u4EFB\u52A1 ID\uFF0C\u7136\u540E\u901A\u8FC7\u7EBF\u7A0B\u6C60\u5F02\u6B65\u6267\u884C\u6821\u9A8C\u548C\u5199\u5165\u3002\u524D\u7AEF\u901A\u8FC7\u8F6E\u8BE2\u63A5\u53E3\u6216 WebSocket \u83B7\u53D6\u5B9E\u65F6\u8FDB\u5EA6\u3002\u5177\u4F53\u5B9E\u73B0\u65B9\u5F0F\u662F\uFF1A\u5728 ReadListener \u7684 invoke \u65B9\u6CD5\u4E2D\u5B9A\u671F\u66F4\u65B0\u8FDB\u5EA6\u4FE1\u606F\u5230\u7F13\u5B58\uFF08\u5982 Redis\uFF09\uFF0C\u524D\u7AEF\u901A\u8FC7\u4EFB\u52A1 ID \u67E5\u8BE2\u5F53\u524D\u5DF2\u6821\u9A8C\u884C\u6570\u548C\u603B\u884C\u6570\uFF0C\u8BA1\u7B97\u51FA\u767E\u5206\u6BD4\u8FDB\u5EA6\u3002"
  ));

  children.push(...codeBlock(`// \u5F02\u6B65\u5BFC\u5165\u670D\u52A1
@Service
public class AsyncImportService {

    @Autowired
    private ExcelImportService importService;
    @Autowired
    private RedisTemplate<String, String> redisTemplate;

    @Async("importExecutor")
    public void asyncImport(String taskId,
                           MultipartFile file,
                           String templateCode) {
        try {
            // \u6267\u884C\u6821\u9A8C
            ImportResult result =
                importService.validateImport(file, templateCode);

            // \u5199\u5165\u6570\u636E\u5E93
            if (!result.isTotalFail() && result.getErrors().isEmpty()) {
                // employeeService.batchInsert(result.getValidRows());
            }

            // \u5B58\u50A8\u7ED3\u679C
            redisTemplate.opsForValue().set(
                "import:result:" + taskId,
                JSON.toJSONString(result),
                30, TimeUnit.MINUTES);
        } catch (Exception e) {
            redisTemplate.opsForValue().set(
                "import:result:" + taskId,
                "{\"totalFail\":true,\"errors\":[{\"errorMessage\":\""
                    + e.getMessage() + "\"}]}",
                30, TimeUnit.MINUTES);
        }
    }
}`));

  children.push(heading("8.3 \u6846\u67B6\u7C7B\u56FE\u6982\u89C8", HeadingLevel.HEADING_2));

  children.push(bodyPara(
    "\u4EE5\u4E0B\u8868\u683C\u5C55\u793A\u4E86\u6574\u4E2A\u6821\u9A8C\u6846\u67B6\u7684\u7C7B\u7ED3\u6784\u53CA\u5176\u804C\u8D23\u5212\u5206\uFF0C\u5E2E\u52A9\u5F00\u53D1\u8005\u5FEB\u901F\u7406\u89E3\u5404\u7EC4\u4EF6\u4E4B\u95F4\u7684\u5173\u7CFB\u3002"
  ));

  children.push(makeTable(
    ["\u5C42\u6B21", "\u7C7B\u540D", "\u804C\u8D23"],
    [
      ["\u914D\u7F6E\u5C42", "ImportTemplateConfig", "\u5BFC\u5165\u6A21\u677F\u914D\u7F6E\uFF1A\u8868\u5934\u3001\u5217\u89C4\u5219\u3001\u552F\u4E00\u7EA6\u675F"],
      ["\u914D\u7F6E\u5C42", "ColumnRule", "\u5355\u5217\u6821\u9A8C\u89C4\u5219\uFF1A\u5FC5\u586B\u3001\u7C7B\u578B\u3001\u957F\u5EA6\u3001\u6B63\u5219"],
      ["\u6A21\u578B\u5C42", "ImportResult", "\u6821\u9A8C\u7ED3\u679C\uFF1A\u6210\u529F\u884C\u6570\u3001\u5931\u8D25\u884C\u6570\u3001\u9519\u8BEF\u660E\u7EC6"],
      ["\u6A21\u578B\u5C42", "RowError", "\u5355\u884C\u9519\u8BEF\u4FE1\u606F\uFF1A\u884C\u53F7\u3001\u5217\u540D\u3001\u9519\u8BEF\u7C7B\u578B\u3001\u63CF\u8FF0"],
      ["\u6821\u9A8C\u5C42", "CellValidator", "\u5355\u5143\u683C\u7EA7\u6821\u9A8C\u63A5\u53E3"],
      ["\u6821\u9A8C\u5C42", "RowValidator", "\u884C\u7EA7\u6821\u9A8C\u63A5\u53E3\uFF08\u8DE8\u5217\u6821\u9A8C\uFF09"],
      ["\u6821\u9A8C\u5C42", "RequiredValidator", "\u5FC5\u586B\u6821\u9A8C\u5B9E\u73B0"],
      ["\u6821\u9A8C\u5C42", "DataTypeValidator", "\u6570\u636E\u7C7B\u578B\u6821\u9A8C\u5B9E\u73B0"],
      ["\u6821\u9A8C\u5C42", "LengthValidator", "\u957F\u5EA6\u6821\u9A8C\u5B9E\u73B0"],
      ["\u6821\u9A8C\u5C42", "RegexValidator", "\u6B63\u5219\u6821\u9A8C\u5B9E\u73B0"],
      ["\u6821\u9A8C\u5C42", "InFileDuplicateValidator", "\u6587\u4EF6\u5185\u91CD\u590D\u6821\u9A8C\u5B9E\u73B0"],
      ["\u6821\u9A8C\u5C42", "DbDuplicateValidator", "\u6570\u636E\u5E93\u53BB\u91CD\u6821\u9A8C\u5B9E\u73B0"],
      ["\u6821\u9A8C\u5C42", "FileFormatValidator", "\u6587\u4EF6\u683C\u5F0F\u963B\u65AD\u6821\u9A8C"],
      ["\u6821\u9A8C\u5C42", "TemplateMatchValidator", "\u6A21\u677F\u5339\u914D\u963B\u65AD\u6821\u9A8C"],
      ["\u8C03\u5EA6\u5C42", "ValidatorRegistry", "\u6821\u9A8C\u5668\u6CE8\u518C\u4E2D\u5FC3\uFF0C\u7BA1\u7406\u6240\u6709 Validator \u5B9E\u4F8B"],
      ["\u8C03\u5EA6\u5C42", "ImportValidationListener", "EasyExcel ReadListener\uFF0C\u8C03\u5EA6\u884C\u7EA7\u6821\u9A8C"],
      ["\u670D\u52A1\u5C42", "ExcelImportService", "\u5BFC\u5165\u4E3B\u5165\u53E3\uFF0C\u7F16\u6392\u6587\u4EF6\u7EA7+\u884C\u7EA7\u6821\u9A8C"],
    ],
    [15, 30, 55]
  ));

  return children;
}

// ── Assemble Document ──
const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
          size: 24,
          color: c(P.body),
        },
        paragraph: {
          spacing: { line: 312 },
        },
      },
      heading1: {
        run: {
          font: { ascii: "Calibri", eastAsia: "SimHei" },
          size: 32,
          bold: true,
          color: c(P.primary),
        },
        paragraph: { spacing: { before: 360, after: 160, line: 312 } },
      },
      heading2: {
        run: {
          font: { ascii: "Calibri", eastAsia: "SimHei" },
          size: 28,
          bold: true,
          color: c(P.primary),
        },
        paragraph: { spacing: { before: 240, after: 120, line: 312 } },
      },
      heading3: {
        run: {
          font: { ascii: "Calibri", eastAsia: "SimHei" },
          size: 24,
          bold: true,
          color: c(P.primary),
        },
        paragraph: { spacing: { before: 200, after: 100, line: 312 } },
      },
    },
  },
  numbering: {
    config: [
      {
        reference: "bullet-list",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          },
        ],
      },
    ],
  },
  sections: [
    // Section 1: Cover
    {
      properties: {
        page: {
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
        },
      },
      children: buildCover(),
    },
    // Section 2: TOC
    {
      properties: {
        page: {
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.UPPER_ROMAN },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Excel\u5BFC\u5165\u6821\u9A8C \u2014 Java \u4EE3\u7801\u8BBE\u8BA1\u65B9\u6848",
                  size: 18,
                  color: P.secondary,
                }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ children: [PageNumber.CURRENT], size: 18, color: P.secondary }),
              ],
            }),
          ],
        }),
      },
      children: buildTocSection(),
    },
    // Section 3: Body
    {
      properties: {
        page: {
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Excel\u5BFC\u5165\u6821\u9A8C \u2014 Java \u4EE3\u7801\u8BBE\u8BA1\u65B9\u6848",
                  size: 18,
                  color: P.secondary,
                }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ children: [PageNumber.CURRENT], size: 18, color: P.secondary }),
              ],
            }),
          ],
        }),
      },
      children: buildBody(),
    },
  ],
});

// ── Export ──
const OUTPUT = "/home/z/my-project/download/Excel\u5BFC\u5165\u6821\u9A8C_Java\u4EE3\u7801\u8BBE\u8BA1\u65B9\u6848.docx";

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(OUTPUT, buffer);
  console.log("Document generated:", OUTPUT);
});
