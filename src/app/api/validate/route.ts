import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import { PRESET_TEMPLATES } from "@/lib/validator/types";
import {
  validateFileFormat,
  validateTemplateMatch,
  validateImportData,
} from "@/lib/validator/engine";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const templateCode = formData.get("templateCode") as string | null;

    if (!file) {
      return NextResponse.json(
        { totalFail: true, errors: [{ errorMessage: "未上传文件" }] },
        { status: 400 }
      );
    }

    if (!templateCode) {
      return NextResponse.json(
        { totalFail: true, errors: [{ errorMessage: "未选择模板" }] },
        { status: 400 }
      );
    }

    // 1. 文件格式校验
    const formatError = validateFileFormat(file.name);
    if (formatError) {
      return NextResponse.json({
        totalFail: true,
        successCount: 0,
        failCount: 0,
        errors: [formatError],
        validRows: [],
        totalRows: 0,
      });
    }

    // 2. 查找模板配置
    const config = PRESET_TEMPLATES.find((t) => t.templateCode === templateCode);
    if (!config) {
      return NextResponse.json(
        {
          totalFail: true,
          successCount: 0,
          failCount: 0,
          errors: [
            {
              rowIndex: 0,
              columnName: "",
              errorType: "TEMPLATE_MISMATCH",
              errorMessage: `未找到模板配置：${templateCode}`,
            },
          ],
          validRows: [],
          totalRows: 0,
        },
        { status: 400 }
      );
    }

    // 3. 解析 Excel
    const buffer = Buffer.from(await file.arrayBuffer());
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      return NextResponse.json({
        totalFail: true,
        successCount: 0,
        failCount: 0,
        errors: [
          {
            rowIndex: 0,
            columnName: "",
            errorType: "FILE_FORMAT",
            errorMessage: "无法读取工作表",
          },
        ],
        validRows: [],
        totalRows: 0,
      });
    }

    // 读取为二维数组，统一用字符串
    const rawData: string[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
    });

    if (rawData.length < 2) {
      return NextResponse.json({
        totalFail: true,
        successCount: 0,
        failCount: 0,
        errors: [
          {
            rowIndex: 0,
            columnName: "",
            errorType: "FILE_FORMAT",
            errorMessage: "文件为空或仅有表头，无数据行",
          },
        ],
        validRows: [],
        totalRows: 0,
      });
    }

    // 4. 模板匹配校验
    const headers = rawData[0].map((h: string) => String(h).trim());
    const templateError = validateTemplateMatch(headers, config);
    if (templateError) {
      return NextResponse.json({
        totalFail: true,
        successCount: 0,
        failCount: 0,
        errors: [templateError],
        validRows: [],
        totalRows: 0,
      });
    }

    // 5. 转换为对象数组
    const dataRows: Record<string, string>[] = [];
    for (let i = 1; i < rawData.length; i++) {
      const row: Record<string, string> = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = String(rawData[i]?.[j] ?? "").trim();
      }
      dataRows.push(row);
    }

    // 6. 执行校验
    const result = validateImportData(headers, dataRows, config);

    return NextResponse.json(result);
  } catch (error) {
    console.error("Validation error:", error);
    return NextResponse.json(
      {
        totalFail: true,
        successCount: 0,
        failCount: 0,
        errors: [
          {
            rowIndex: 0,
            columnName: "",
            errorType: "FILE_FORMAT",
            errorMessage: `文件解析失败：${error instanceof Error ? error.message : "未知错误"}`,
          },
        ],
        validRows: [],
        totalRows: 0,
      },
      { status: 500 }
    );
  }
}
