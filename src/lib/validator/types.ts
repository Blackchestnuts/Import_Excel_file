// ═══ 校验配置模型 ═══

export type DataType = "STRING" | "NUMBER" | "DATE" | "ENUM";

export interface ColumnRule {
  columnName: string;
  required: boolean;
  dataType: DataType;
  maxLength?: number;
  minLength?: number;
  regexPattern?: string;
  enumValues?: string[];
  errorMsg?: string;
}

export interface ImportTemplateConfig {
  templateCode: string;
  templateName: string;
  expectedHeaders: string[];
  columnRules: Record<string, ColumnRule>;
  uniqueKeyGroups: string[][];
  maxRows: number;
}

// ═══ 校验结果模型 ═══

export type ValidationErrorType =
  | "FILE_FORMAT"
  | "TEMPLATE_MISMATCH"
  | "REQUIRED_EMPTY"
  | "TYPE_MISMATCH"
  | "LENGTH_EXCEEDED"
  | "DUPLICATE_IN_FILE"
  | "DUPLICATE_IN_DB"
  | "REGEX_MISMATCH"
  | "ENUM_INVALID";

export interface RowError {
  rowIndex: number;
  columnName: string;
  errorType: ValidationErrorType;
  errorMessage: string;
}

export interface ImportResult {
  totalFail: boolean;
  successCount: number;
  failCount: number;
  errors: RowError[];
  validRows: Record<string, string>[];
  totalRows: number;
}

// ═══ 错误类型中文映射 ═══

export const ERROR_TYPE_LABELS: Record<ValidationErrorType, string> = {
  FILE_FORMAT: "文件格式错误",
  TEMPLATE_MISMATCH: "模板不匹配",
  REQUIRED_EMPTY: "必填为空",
  TYPE_MISMATCH: "类型不匹配",
  LENGTH_EXCEEDED: "长度超限",
  DUPLICATE_IN_FILE: "文件内重复",
  DUPLICATE_IN_DB: "数据库重复",
  REGEX_MISMATCH: "格式不正确",
  ENUM_INVALID: "枚举值非法",
};

export const ERROR_TYPE_COLORS: Record<ValidationErrorType, string> = {
  FILE_FORMAT: "bg-red-100 text-red-800",
  TEMPLATE_MISMATCH: "bg-red-100 text-red-800",
  REQUIRED_EMPTY: "bg-orange-100 text-orange-800",
  TYPE_MISMATCH: "bg-yellow-100 text-yellow-800",
  LENGTH_EXCEEDED: "bg-amber-100 text-amber-800",
  DUPLICATE_IN_FILE: "bg-purple-100 text-purple-800",
  DUPLICATE_IN_DB: "bg-pink-100 text-pink-800",
  REGEX_MISMATCH: "bg-cyan-100 text-cyan-800",
  ENUM_INVALID: "bg-teal-100 text-teal-800",
};

// ═══ 预设模板配置 ═══

export const PRESET_TEMPLATES: ImportTemplateConfig[] = [
  {
    templateCode: "EMPLOYEE_IMPORT",
    templateName: "员工信息导入",
    expectedHeaders: ["姓名", "年龄", "手机号", "入职日期", "部门"],
    columnRules: {
      "姓名": {
        columnName: "姓名",
        required: true,
        dataType: "STRING",
        maxLength: 10,
      },
      "年龄": {
        columnName: "年龄",
        required: true,
        dataType: "NUMBER",
        maxLength: 3,
      },
      "手机号": {
        columnName: "手机号",
        required: true,
        dataType: "STRING",
        minLength: 11,
        maxLength: 11,
        regexPattern: "^1[3-9]\\d{9}$",
      },
      "入职日期": {
        columnName: "入职日期",
        required: false,
        dataType: "DATE",
      },
      "部门": {
        columnName: "部门",
        required: true,
        dataType: "ENUM",
        enumValues: ["技术部", "市场部", "财务部", "人事部", "运营部"],
      },
    },
    uniqueKeyGroups: [["手机号"]],
    maxRows: 10000,
  },
  {
    templateCode: "STUDENT_IMPORT",
    templateName: "学生信息导入",
    expectedHeaders: ["学号", "姓名", "性别", "班级", "手机号"],
    columnRules: {
      "学号": {
        columnName: "学号",
        required: true,
        dataType: "STRING",
        minLength: 8,
        maxLength: 12,
      },
      "姓名": {
        columnName: "姓名",
        required: true,
        dataType: "STRING",
        maxLength: 20,
      },
      "性别": {
        columnName: "性别",
        required: true,
        dataType: "ENUM",
        enumValues: ["男", "女"],
      },
      "班级": {
        columnName: "班级",
        required: true,
        dataType: "STRING",
        maxLength: 30,
      },
      "手机号": {
        columnName: "手机号",
        required: true,
        dataType: "STRING",
        minLength: 11,
        maxLength: 11,
        regexPattern: "^1[3-9]\\d{9}$",
      },
    },
    uniqueKeyGroups: [["学号"], ["手机号"]],
    maxRows: 5000,
  },
];
