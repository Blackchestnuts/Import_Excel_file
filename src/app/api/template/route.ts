import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import { PRESET_TEMPLATES } from "@/lib/validator/types";

export async function GET(request: NextRequest) {
  const { searchParams } = new URL(request.url);
  const templateCode = searchParams.get("templateCode");

  if (!templateCode) {
    return NextResponse.json({ error: "缺少 templateCode 参数" }, { status: 400 });
  }

  const config = PRESET_TEMPLATES.find((t) => t.templateCode === templateCode);
  if (!config) {
    return NextResponse.json({ error: "模板不存在" }, { status: 404 });
  }

  // 生成模板 Excel
  const wb = XLSX.utils.book_new();

  // 表头行
  const headerRow = config.expectedHeaders;

  // 示例数据行（帮助用户理解格式）
  const sampleRow = config.expectedHeaders.map((header) => {
    const rule = config.columnRules[header];
    if (rule?.enumValues?.length) return rule.enumValues[0];
    switch (rule?.dataType) {
      case "NUMBER": return "25";
      case "DATE": return "2024-01-15";
      default: return "示例数据";
    }
  });

  // 规则提示行
  const ruleRow = config.expectedHeaders.map((header) => {
    const rule = config.columnRules[header];
    const hints: string[] = [];
    if (rule?.required) hints.push("必填");
    if (rule?.dataType === "NUMBER") hints.push("数字");
    if (rule?.dataType === "DATE") hints.push("日期(yyyy-MM-dd)");
    if (rule?.enumValues?.length) hints.push(`可选: ${rule.enumValues.join("/")}`);
    if (rule?.maxLength) hints.push(`最长${rule.maxLength}字符`);
    if (rule?.minLength) hints.push(`最短${rule.minLength}字符`);
    if (rule?.regexPattern) hints.push("需符合特定格式");
    return hints.length > 0 ? `[${hints.join(", ")}]` : "";
  });

  const wsData = [headerRow, sampleRow, ruleRow];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // 设置列宽
  ws["!cols"] = config.expectedHeaders.map(() => ({ wch: 20 }));

  XLSX.utils.book_append_sheet(wb, ws, "导入数据");

  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  return new NextResponse(buffer, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${encodeURIComponent(config.templateName + "_模板.xlsx")}"`,
    },
  });
}
