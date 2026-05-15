"use client";

import React, { useState, useCallback, useRef } from "react";
import {
  PRESET_TEMPLATES,
  ERROR_TYPE_LABELS,
  ERROR_TYPE_COLORS,
  type ImportResult,
  type RowError,
  type ValidationErrorType,
} from "@/lib/validator/types";
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  XCircle,
  AlertTriangle,
  Download,
  RotateCcw,
  ChevronDown,
  ChevronRight,
  FileX2,
  Shield,
  Table2,
  BarChart3,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Progress } from "@/components/ui/progress";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Separator } from "@/components/ui/separator";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useToast } from "@/hooks/use-toast";

export default function Home() {
  const [selectedTemplate, setSelectedTemplate] = useState<string>("");
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isValidating, setIsValidating] = useState(false);
  const [result, setResult] = useState<ImportResult | null>(null);
  const [expandedErrors, setExpandedErrors] = useState<Set<number>>(new Set());
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const currentTemplate = PRESET_TEMPLATES.find(
    (t) => t.templateCode === selectedTemplate
  );

  // ── 文件拖放处理 ──
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const droppedFile = e.dataTransfer.files[0];
      if (
        droppedFile &&
        (droppedFile.name.endsWith(".xlsx") ||
          droppedFile.name.endsWith(".xls"))
      ) {
        setFile(droppedFile);
        setResult(null);
      } else {
        toast({
          title: "文件格式错误",
          description: "请上传 .xlsx 或 .xls 格式的 Excel 文件",
          variant: "destructive",
        });
      }
    },
    [toast]
  );

  const handleFileSelect = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const selected = e.target.files?.[0];
      if (selected) {
        setFile(selected);
        setResult(null);
      }
    },
    []
  );

  // ── 下载模板 ──
  const handleDownloadTemplate = useCallback(async () => {
    if (!selectedTemplate) {
      toast({
        title: "请先选择模板",
        description: "下载模板前需要选择一个导入模板类型",
        variant: "destructive",
      });
      return;
    }
    try {
      const res = await fetch(
        `/api/template?templateCode=${selectedTemplate}`
      );
      if (!res.ok) throw new Error("下载失败");
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${currentTemplate?.templateName || "导入"}_模板.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
      toast({ title: "模板下载成功" });
    } catch {
      toast({
        title: "下载失败",
        description: "模板下载出错，请重试",
        variant: "destructive",
      });
    }
  }, [selectedTemplate, currentTemplate, toast]);

  // ── 开始校验 ──
  const handleValidate = useCallback(async () => {
    if (!file || !selectedTemplate) {
      toast({
        title: "信息不完整",
        description: "请选择模板并上传文件",
        variant: "destructive",
      });
      return;
    }

    setIsValidating(true);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("templateCode", selectedTemplate);

      const res = await fetch("/api/validate", {
        method: "POST",
        body: formData,
      });

      const data: ImportResult = await res.json();
      setResult(data);

      if (data.totalFail) {
        toast({
          title: "全部导入失败",
          description: "文件级校验未通过，请检查文件格式和模板",
          variant: "destructive",
        });
      } else if (data.errors.length > 0) {
        toast({
          title: "校验完成，存在错误",
          description: `成功 ${data.successCount} 行，失败 ${data.failCount} 行`,
          variant: "default",
        });
      } else {
        toast({
          title: "校验全部通过",
          description: `共 ${data.successCount} 行数据全部有效`,
        });
      }
    } catch {
      toast({
        title: "校验失败",
        description: "网络或服务异常，请重试",
        variant: "destructive",
      });
    } finally {
      setIsValidating(false);
    }
  }, [file, selectedTemplate, toast]);

  // ── 重置 ──
  const handleReset = useCallback(() => {
    setFile(null);
    setResult(null);
    setSelectedTemplate("");
    setExpandedErrors(new Set());
    if (fileInputRef.current) fileInputRef.current.value = "";
  }, []);

  // ── 按行号分组错误 ──
  const groupedErrors = React.useMemo(() => {
    if (!result?.errors) return {};
    const groups: Record<number, RowError[]> = {};
    for (const err of result.errors) {
      if (!groups[err.rowIndex]) groups[err.rowIndex] = [];
      groups[err.rowIndex].push(err);
    }
    return groups;
  }, [result]);

  const toggleErrorGroup = useCallback((rowIndex: number) => {
    setExpandedErrors((prev) => {
      const next = new Set(prev);
      if (next.has(rowIndex)) next.delete(rowIndex);
      else next.add(rowIndex);
      return next;
    });
  }, []);

  // ── 错误类型统计 ──
  const errorTypeStats = React.useMemo(() => {
    if (!result?.errors) return {};
    const stats: Record<ValidationErrorType, number> = {} as Record<
      ValidationErrorType,
      number
    >;
    for (const err of result.errors) {
      stats[err.errorType] = (stats[err.errorType] || 0) + 1;
    }
    return stats;
  }, [result]);

  const successRate =
    result && result.totalRows > 0
      ? Math.round((result.successCount / result.totalRows) * 100)
      : 0;

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-slate-100">
      {/* Header */}
      <header className="border-b bg-white/80 backdrop-blur-sm sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center">
              <Shield className="w-5 h-5 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-900">
                Excel 导入校验
              </h1>
              <p className="text-sm text-slate-500">
                上传 Excel 文件，自动校验数据质量
              </p>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Step 1: Template Selection */}
        <div className="mb-8">
          <div className="flex items-center gap-2 mb-4">
            <span className="w-7 h-7 rounded-full bg-emerald-600 text-white text-sm font-bold flex items-center justify-center">
              1
            </span>
            <h2 className="text-lg font-semibold text-slate-900">
              选择导入模板
            </h2>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {PRESET_TEMPLATES.map((tpl) => (
              <Card
                key={tpl.templateCode}
                className={`cursor-pointer transition-all duration-200 hover:shadow-md ${
                  selectedTemplate === tpl.templateCode
                    ? "ring-2 ring-emerald-500 shadow-md"
                    : "hover:border-slate-300"
                }`}
                onClick={() => {
                  setSelectedTemplate(tpl.templateCode);
                  setResult(null);
                }}
              >
                <CardHeader className="pb-2">
                  <div className="flex items-center justify-between">
                    <CardTitle className="text-base flex items-center gap-2">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                      {tpl.templateName}
                    </CardTitle>
                    {selectedTemplate === tpl.templateCode && (
                      <Badge className="bg-emerald-100 text-emerald-800 hover:bg-emerald-100">
                        已选择
                      </Badge>
                    )}
                  </div>
                </CardHeader>
                <CardContent>
                  <div className="flex flex-wrap gap-1.5 mb-3">
                    {tpl.expectedHeaders.map((h) => {
                      const rule = tpl.columnRules[h];
                      return (
                        <Badge
                          key={h}
                          variant="outline"
                          className="text-xs font-normal"
                        >
                          {h}
                          {rule?.required && (
                            <span className="text-red-500 ml-0.5">*</span>
                          )}
                        </Badge>
                      );
                    })}
                  </div>
                  <div className="text-xs text-slate-500 space-y-1">
                    <div>
                      唯一性约束：
                      {tpl.uniqueKeyGroups
                        .map((g) => g.join("+"))
                        .join("、") || "无"}
                    </div>
                    <div>最大行数：{tpl.maxRows.toLocaleString()}</div>
                  </div>
                </CardContent>
              </Card>
            ))}
          </div>
          {selectedTemplate && (
            <div className="mt-3 flex justify-end">
              <Button
                variant="outline"
                size="sm"
                onClick={handleDownloadTemplate}
                className="gap-1.5"
              >
                <Download className="w-3.5 h-3.5" />
                下载该模板
              </Button>
            </div>
          )}
        </div>

        {/* Step 2: File Upload */}
        <div className="mb-8">
          <div className="flex items-center gap-2 mb-4">
            <span className="w-7 h-7 rounded-full bg-emerald-600 text-white text-sm font-bold flex items-center justify-center">
              2
            </span>
            <h2 className="text-lg font-semibold text-slate-900">
              上传 Excel 文件
            </h2>
          </div>
          <div
            className={`border-2 border-dashed rounded-xl p-8 text-center transition-all duration-200 ${
              isDragging
                ? "border-emerald-500 bg-emerald-50"
                : file
                ? "border-emerald-300 bg-emerald-50/50"
                : "border-slate-300 bg-slate-50 hover:border-slate-400"
            }`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            {file ? (
              <div className="flex items-center justify-center gap-3">
                <FileSpreadsheet className="w-10 h-10 text-emerald-600" />
                <div className="text-left">
                  <p className="font-medium text-slate-900">{file.name}</p>
                  <p className="text-sm text-slate-500">
                    {(file.size / 1024).toFixed(1)} KB
                  </p>
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => {
                    setFile(null);
                    setResult(null);
                    if (fileInputRef.current) fileInputRef.current.value = "";
                  }}
                  className="ml-4 text-slate-400 hover:text-red-500"
                >
                  <XCircle className="w-4 h-4" />
                </Button>
              </div>
            ) : (
              <div>
                <Upload className="w-12 h-12 mx-auto text-slate-400 mb-3" />
                <p className="text-slate-600 font-medium">
                  拖拽文件至此处，或点击选择文件
                </p>
                <p className="text-sm text-slate-400 mt-1">
                  支持 .xlsx、.xls 格式
                </p>
              </div>
            )}
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={handleFileSelect}
              id="file-upload"
            />
            {!file && (
              <Button
                variant="outline"
                className="mt-4"
                onClick={() => fileInputRef.current?.click()}
              >
                选择文件
              </Button>
            )}
          </div>
        </div>

        {/* Action Buttons */}
        <div className="flex items-center gap-3 mb-8">
          <Button
            onClick={handleValidate}
            disabled={!file || !selectedTemplate || isValidating}
            className="bg-emerald-600 hover:bg-emerald-700 gap-2 px-8"
            size="lg"
          >
            {isValidating ? (
              <>
                <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                校验中...
              </>
            ) : (
              <>
                <Shield className="w-4 h-4" />
                开始校验
              </>
            )}
          </Button>
          <Button
            variant="outline"
            onClick={handleReset}
            className="gap-2"
            size="lg"
          >
            <RotateCcw className="w-4 h-4" />
            重置
          </Button>
        </div>

        {/* Step 3: Results */}
        {result && (
          <div>
            <div className="flex items-center gap-2 mb-4">
              <span className="w-7 h-7 rounded-full bg-emerald-600 text-white text-sm font-bold flex items-center justify-center">
                3
              </span>
              <h2 className="text-lg font-semibold text-slate-900">
                校验结果
              </h2>
            </div>

            {/* Overview Cards */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
              <Card
                className={
                  result.totalFail
                    ? "border-red-200 bg-red-50"
                    : result.errors.length === 0
                    ? "border-emerald-200 bg-emerald-50"
                    : "border-amber-200 bg-amber-50"
                }
              >
                <CardContent className="pt-5">
                  <div className="flex items-center gap-3">
                    {result.totalFail ? (
                      <FileX2 className="w-8 h-8 text-red-600" />
                    ) : result.errors.length === 0 ? (
                      <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                    ) : (
                      <AlertTriangle className="w-8 h-8 text-amber-600" />
                    )}
                    <div>
                      <p className="text-sm font-medium text-slate-500">
                        校验状态
                      </p>
                      <p className="text-lg font-bold text-slate-900">
                        {result.totalFail
                          ? "全部失败"
                          : result.errors.length === 0
                          ? "全部通过"
                          : "部分失败"}
                      </p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardContent className="pt-5">
                  <div className="flex items-center gap-3">
                    <BarChart3 className="w-8 h-8 text-slate-400" />
                    <div>
                      <p className="text-sm font-medium text-slate-500">
                        总行数
                      </p>
                      <p className="text-lg font-bold text-slate-900">
                        {result.totalRows}
                      </p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-emerald-200">
                <CardContent className="pt-5">
                  <div className="flex items-center gap-3">
                    <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                    <div>
                      <p className="text-sm font-medium text-slate-500">
                        通过行数
                      </p>
                      <p className="text-lg font-bold text-emerald-700">
                        {result.successCount}
                      </p>
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-red-200">
                <CardContent className="pt-5">
                  <div className="flex items-center gap-3">
                    <XCircle className="w-8 h-8 text-red-500" />
                    <div>
                      <p className="text-sm font-medium text-slate-500">
                        失败行数
                      </p>
                      <p className="text-lg font-bold text-red-700">
                        {result.failCount}
                      </p>
                    </div>
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* Success Rate Progress */}
            {!result.totalFail && (
              <Card className="mb-6">
                <CardContent className="pt-5">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-sm font-medium text-slate-600">
                      校验通过率
                    </span>
                    <span className="text-sm font-bold text-slate-900">
                      {successRate}%
                    </span>
                  </div>
                  <Progress
                    value={successRate}
                    className="h-3"
                  />
                </CardContent>
              </Card>
            )}

            {/* Error Type Stats */}
            {Object.keys(errorTypeStats).length > 0 && (
              <Card className="mb-6">
                <CardHeader className="pb-3">
                  <CardTitle className="text-base flex items-center gap-2">
                    <AlertTriangle className="w-4 h-4 text-amber-500" />
                    错误类型分布
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="flex flex-wrap gap-2">
                    {(
                      Object.entries(errorTypeStats) as [
                        ValidationErrorType,
                        number
                      ][]
                    ).map(([type, count]) => (
                      <Badge
                        key={type}
                        className={`${ERROR_TYPE_COLORS[type]} text-sm px-3 py-1`}
                      >
                        {ERROR_TYPE_LABELS[type]}：{count}
                      </Badge>
                    ))}
                  </div>
                </CardContent>
              </Card>
            )}

            {/* Detail Tabs */}
            {!result.totalFail && (
              <Tabs defaultValue="errors" className="space-y-4">
                <TabsList>
                  <TabsTrigger value="errors" className="gap-1.5">
                    <XCircle className="w-3.5 h-3.5" />
                    错误明细
                    {result.errors.length > 0 && (
                      <Badge
                        variant="secondary"
                        className="ml-1 text-xs px-1.5"
                      >
                        {result.errors.length}
                      </Badge>
                    )}
                  </TabsTrigger>
                  <TabsTrigger value="valid" className="gap-1.5">
                    <Table2 className="w-3.5 h-3.5" />
                    有效数据
                    {result.validRows.length > 0 && (
                      <Badge
                        variant="secondary"
                        className="ml-1 text-xs px-1.5"
                      >
                        {result.validRows.length}
                      </Badge>
                    )}
                  </TabsTrigger>
                </TabsList>

                {/* Errors Tab */}
                <TabsContent value="errors">
                  {result.errors.length === 0 ? (
                    <Card>
                      <CardContent className="py-12 text-center">
                        <CheckCircle2 className="w-16 h-16 mx-auto text-emerald-500 mb-4" />
                        <p className="text-lg font-semibold text-emerald-700">
                          全部校验通过！
                        </p>
                        <p className="text-sm text-slate-500 mt-1">
                          共 {result.successCount} 行数据均有效，可安全导入
                        </p>
                      </CardContent>
                    </Card>
                  ) : (
                    <Card>
                      <CardContent className="p-0">
                        <ScrollArea className="max-h-[500px]">
                          {Object.entries(groupedErrors)
                            .sort(([a], [b]) => Number(a) - Number(b))
                            .map(([rowIndex, errors]) => (
                              <div key={rowIndex}>
                                <div
                                  className="flex items-center gap-2 px-4 py-3 cursor-pointer hover:bg-slate-50 transition-colors"
                                  onClick={() =>
                                    toggleErrorGroup(Number(rowIndex))
                                  }
                                >
                                  {expandedErrors.has(Number(rowIndex)) ? (
                                    <ChevronDown className="w-4 h-4 text-slate-400" />
                                  ) : (
                                    <ChevronRight className="w-4 h-4 text-slate-400" />
                                  )}
                                  <Badge
                                    variant="outline"
                                    className="text-xs font-mono"
                                  >
                                    行 {rowIndex}
                                  </Badge>
                                  <span className="text-sm text-slate-600">
                                    {errors.length} 个错误
                                  </span>
                                  <div className="flex gap-1 ml-auto">
                                    {[
                                      ...new Set(
                                        errors.map((e) => e.errorType)
                                      ),
                                    ].map((type) => (
                                      <Badge
                                        key={type}
                                        className={`${ERROR_TYPE_COLORS[type]} text-[10px] px-1.5 py-0`}
                                      >
                                        {ERROR_TYPE_LABELS[type]}
                                      </Badge>
                                    ))}
                                  </div>
                                </div>
                                {expandedErrors.has(Number(rowIndex)) && (
                                  <div className="px-8 pb-3 space-y-1.5">
                                    {errors.map((err, idx) => (
                                      <div
                                        key={idx}
                                        className="flex items-start gap-2 text-sm"
                                      >
                                        <XCircle className="w-3.5 h-3.5 text-red-400 mt-0.5 shrink-0" />
                                        <span className="text-slate-700">
                                          [{err.columnName}]
                                        </span>
                                        <span className="text-slate-500">
                                          {err.errorMessage}
                                        </span>
                                      </div>
                                    ))}
                                  </div>
                                )}
                                <Separator />
                              </div>
                            ))}
                        </ScrollArea>
                      </CardContent>
                    </Card>
                  )}
                </TabsContent>

                {/* Valid Data Tab */}
                <TabsContent value="valid">
                  {result.validRows.length === 0 ? (
                    <Card>
                      <CardContent className="py-12 text-center">
                        <XCircle className="w-16 h-16 mx-auto text-slate-300 mb-4" />
                        <p className="text-lg font-semibold text-slate-500">
                          无有效数据
                        </p>
                        <p className="text-sm text-slate-400 mt-1">
                          所有行均未通过校验
                        </p>
                      </CardContent>
                    </Card>
                  ) : (
                    <Card>
                      <CardContent className="p-0">
                        <ScrollArea className="max-h-[500px]">
                          <Table>
                            <TableHeader>
                              <TableRow>
                                <TableHead className="w-16 text-center">
                                  #
                                </TableHead>
                                {currentTemplate?.expectedHeaders.map((h) => (
                                  <TableHead key={h}>{h}</TableHead>
                                ))}
                              </TableRow>
                            </TableHeader>
                            <TableBody>
                              {result.validRows.map((row, idx) => (
                                <TableRow key={idx}>
                                  <TableCell className="text-center text-slate-400 text-xs font-mono">
                                    {idx + 2}
                                  </TableCell>
                                  {currentTemplate?.expectedHeaders.map(
                                    (h) => (
                                      <TableCell key={h} className="text-sm">
                                        {row[h] || "-"}
                                      </TableCell>
                                    )
                                  )}
                                </TableRow>
                              ))}
                            </TableBody>
                          </Table>
                        </ScrollArea>
                      </CardContent>
                    </Card>
                  )}
                </TabsContent>
              </Tabs>
            )}

            {/* Total Fail Display */}
            {result.totalFail && (
              <Card className="border-red-200 bg-red-50">
                <CardContent className="py-8 text-center">
                  <FileX2 className="w-16 h-16 mx-auto text-red-400 mb-4" />
                  <p className="text-lg font-semibold text-red-700">
                    全部导入失败
                  </p>
                  <p className="text-sm text-red-600 mt-2">
                    {result.errors[0]?.errorMessage}
                  </p>
                  <div className="mt-4">
                    <Button
                      variant="outline"
                      onClick={handleDownloadTemplate}
                      className="gap-2"
                    >
                      <Download className="w-4 h-4" />
                      下载最新模板
                    </Button>
                  </div>
                </CardContent>
              </Card>
            )}
          </div>
        )}

        {/* Template Rules Reference */}
        {currentTemplate && !result && (
          <Card className="mt-8">
            <CardHeader>
              <CardTitle className="text-base flex items-center gap-2">
                <Shield className="w-4 h-4 text-emerald-600" />
                {currentTemplate.templateName} — 校验规则一览
              </CardTitle>
              <CardDescription>
                以下为该模板的所有校验规则，请确保数据符合要求
              </CardDescription>
            </CardHeader>
            <CardContent>
              <ScrollArea className="max-h-[400px]">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>字段名</TableHead>
                      <TableHead>必填</TableHead>
                      <TableHead>类型</TableHead>
                      <TableHead>长度</TableHead>
                      <TableHead>格式</TableHead>
                      <TableHead>可选值</TableHead>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {currentTemplate.expectedHeaders.map((header) => {
                      const rule = currentTemplate.columnRules[header];
                      return (
                        <TableRow key={header}>
                          <TableCell className="font-medium">
                            {header}
                            {rule?.required && (
                              <span className="text-red-500 ml-0.5">*</span>
                            )}
                          </TableCell>
                          <TableCell>
                            {rule?.required ? (
                              <Badge className="bg-red-100 text-red-700 text-xs">
                                必填
                              </Badge>
                            ) : (
                              <Badge
                                variant="outline"
                                className="text-xs text-slate-400"
                              >
                                选填
                              </Badge>
                            )}
                          </TableCell>
                          <TableCell>
                            <Badge variant="outline" className="text-xs">
                              {rule?.dataType || "-"}
                            </Badge>
                          </TableCell>
                          <TableCell className="text-sm text-slate-600">
                            {rule?.minLength && rule?.maxLength
                              ? `${rule.minLength}-${rule.maxLength}`
                              : rule?.maxLength
                              ? `≤${rule.maxLength}`
                              : rule?.minLength
                              ? `≥${rule.minLength}`
                              : "-"}
                          </TableCell>
                          <TableCell className="text-xs text-slate-500 font-mono">
                            {rule?.regexPattern || "-"}
                          </TableCell>
                          <TableCell className="text-xs text-slate-500">
                            {rule?.enumValues?.join("、") || "-"}
                          </TableCell>
                        </TableRow>
                      );
                    })}
                  </TableBody>
                </Table>
              </ScrollArea>
              {currentTemplate.uniqueKeyGroups.length > 0 && (
                <div className="mt-4 pt-4 border-t">
                  <p className="text-sm text-slate-500">
                    <span className="font-medium">唯一性约束：</span>
                    {currentTemplate.uniqueKeyGroups
                      .map((g) => g.join(" + "))
                      .join("、")}
                  </p>
                </div>
              )}
            </CardContent>
          </Card>
        )}
      </main>

      {/* Footer */}
      <footer className="border-t bg-white mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 text-center text-sm text-slate-400">
          Excel 导入校验工具 — 基于 Java + EasyExcel 设计方案的前端实现
        </div>
      </footer>
    </div>
  );
}
