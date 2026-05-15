import {
  ImportTemplateConfig,
  ColumnRule,
  RowError,
  ImportResult,
  ValidationErrorType,
} from "./types";

// ═══ 文件格式校验 ═══

const ALLOWED_EXTENSIONS = [".xls", ".xlsx"];

export function validateFileFormat(fileName: string): RowError | null {
  if (!fileName) {
    return {
      rowIndex: 0,
      columnName: "",
      errorType: "FILE_FORMAT",
      errorMessage: "无法识别文件名",
    };
  }
  const ext = fileName.substring(fileName.lastIndexOf(".")).toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    return {
      rowIndex: 0,
      columnName: "",
      errorType: "FILE_FORMAT",
      errorMessage: "文件格式不正确，仅支持 .xls 和 .xlsx 格式",
    };
  }
  return null;
}

// ═══ 模板匹配校验 ═══

export function validateTemplateMatch(
  actualHeaders: string[],
  config: ImportTemplateConfig
): RowError | null {
  const expected = config.expectedHeaders;

  if (actualHeaders.length !== expected.length) {
    return {
      rowIndex: 0,
      columnName: "",
      errorType: "TEMPLATE_MISMATCH",
      errorMessage: `导入模板不匹配：期望 ${expected.length} 列，实际 ${actualHeaders.length} 列，请下载最新模板`,
    };
  }

  const mismatches: string[] = [];
  for (let i = 0; i < expected.length; i++) {
    if (expected[i] !== actualHeaders[i]) {
      mismatches.push(
        `第${i + 1}列：期望[${expected[i]}]，实际[${actualHeaders[i]}]`
      );
    }
  }

  if (mismatches.length > 0) {
    return {
      rowIndex: 0,
      columnName: "",
      errorType: "TEMPLATE_MISMATCH",
      errorMessage: `导入模板不匹配，请下载最新模板。详情：${mismatches.join("；")}`,
    };
  }

  return null;
}

// ═══ 单元格级校验 ═══

function isNumber(value: string): boolean {
  return !isNaN(Number(value)) && value.trim() !== "";
}

function isDate(value: string): boolean {
  const patterns = [
    /^\d{4}-\d{2}-\d{2}$/,
    /^\d{4}\/\d{2}\/\d{2}$/,
    /^\d{8}$/,
  ];
  for (const p of patterns) {
    if (p.test(value)) {
      const d = new Date(value.replace(/\//g, "-"));
      return !isNaN(d.getTime());
    }
  }
  return false;
}

function isValidEnum(value: string, enumValues?: string[]): boolean {
  return enumValues != null && enumValues.includes(value);
}

function validateCell(
  rowIndex: number,
  columnName: string,
  cellValue: string,
  rule: ColumnRule
): RowError[] {
  const errors: RowError[] = [];
  const trimmed = cellValue.trim();

  // 必填校验
  if (rule.required && trimmed === "") {
    errors.push({
      rowIndex,
      columnName,
      errorType: "REQUIRED_EMPTY",
      errorMessage: rule.errorMsg || `第${rowIndex}行[${columnName}]不能为空`,
    });
    return errors; // 必填为空时跳过后续校验
  }

  if (trimmed === "") return errors;

  // 类型校验
  let typeValid = true;
  switch (rule.dataType) {
    case "NUMBER":
      typeValid = isNumber(trimmed);
      break;
    case "DATE":
      typeValid = isDate(trimmed);
      break;
    case "ENUM":
      typeValid = isValidEnum(trimmed, rule.enumValues);
      break;
    case "STRING":
    default:
      typeValid = true;
  }
  if (!typeValid) {
    let hint = `应为[${rule.dataType}]`;
    if (rule.dataType === "ENUM" && rule.enumValues) {
      hint = `应为：${rule.enumValues.join("、")}`;
    }
    errors.push({
      rowIndex,
      columnName,
      errorType: rule.dataType === "ENUM" ? "ENUM_INVALID" : "TYPE_MISMATCH",
      errorMessage: `第${rowIndex}行[${columnName}]数据类型错误，${hint}`,
    });
  }

  // 长度校验
  if (rule.maxLength != null && trimmed.length > rule.maxLength) {
    errors.push({
      rowIndex,
      columnName,
      errorType: "LENGTH_EXCEEDED",
      errorMessage: `第${rowIndex}行[${columnName}]长度不能超过${rule.maxLength}个字符，当前${trimmed.length}个`,
    });
  }
  if (rule.minLength != null && trimmed.length < rule.minLength) {
    errors.push({
      rowIndex,
      columnName,
      errorType: "LENGTH_EXCEEDED",
      errorMessage: `第${rowIndex}行[${columnName}]长度不能少于${rule.minLength}个字符，当前${trimmed.length}个`,
    });
  }

  // 正则校验
  if (rule.regexPattern) {
    try {
      const regex = new RegExp(rule.regexPattern);
      if (!regex.test(trimmed)) {
        errors.push({
          rowIndex,
          columnName,
          errorType: "REGEX_MISMATCH",
          errorMessage: `第${rowIndex}行[${columnName}]格式不正确`,
        });
      }
    } catch {
      // 正则表达式无效，跳过
    }
  }

  return errors;
}

// ═══ 文件内重复校验 ═══

function validateDuplicates(
  rows: Record<string, string>[],
  config: ImportTemplateConfig
): RowError[] {
  const errors: RowError[] = [];

  for (const keyGroup of config.uniqueKeyGroups) {
    const history = new Map<string, number>();

    for (let i = 0; i < rows.length; i++) {
      const compositeKey = keyGroup
        .map((col) => rows[i][col] || "")
        .join("|");

      if (history.has(compositeKey)) {
        const firstRow = history.get(compositeKey)!;
        errors.push({
          rowIndex: i + 2,
          columnName: keyGroup.join("+"),
          errorType: "DUPLICATE_IN_FILE",
          errorMessage: `第${i + 2}行与第${firstRow}行的[${keyGroup.join("+")}]重复`,
        });
      } else {
        history.set(compositeKey, i + 2);
      }
    }
  }

  return errors;
}

// ═══ 主校验入口 ═══

export function validateImportData(
  headers: string[],
  rows: Record<string, string>[],
  config: ImportTemplateConfig
): ImportResult {
  // 行数限制
  if (rows.length > config.maxRows) {
    return {
      totalFail: true,
      successCount: 0,
      failCount: 0,
      errors: [
        {
          rowIndex: 0,
          columnName: "",
          errorType: "FILE_FORMAT" as ValidationErrorType,
          errorMessage: `文件行数超过限制（最大${config.maxRows}行），当前${rows.length}行`,
        },
      ],
      validRows: [],
      totalRows: rows.length,
    };
  }

  const allErrors: RowError[] = [];
  const validRows: Record<string, string>[] = [];

  // 逐行逐列校验
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowIndex = i + 2; // Excel 行号（1-based + 表头1行）
    const rowErrors: RowError[] = [];

    for (const header of config.expectedHeaders) {
      const rule = config.columnRules[header];
      if (!rule) continue;
      const cellValue = row[header] ?? "";
      const cellErrors = validateCell(rowIndex, header, cellValue, rule);
      rowErrors.push(...cellErrors);
    }

    if (rowErrors.length === 0) {
      validRows.push(row);
    } else {
      allErrors.push(...rowErrors);
    }
  }

  // 文件内重复校验（对所有行，包括有错误的行也检查重复）
  const dupErrors = validateDuplicates(rows, config);
  allErrors.push(...dupErrors);

  // 从 validRows 中移除重复行
  const dupRowIndices = new Set(dupErrors.map((e) => e.rowIndex));
  const filteredValidRows = validRows.filter(
    (_, idx) => !dupRowIndices.has(idx + 2)
  );

  return {
    totalFail: false,
    successCount: filteredValidRows.length,
    failCount: rows.length - filteredValidRows.length,
    errors: allErrors,
    validRows: filteredValidRows,
    totalRows: rows.length,
  };
}
