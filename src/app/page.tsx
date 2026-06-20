"use client";

import React, { useState, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  AreaChart,
  Area,
} from "recharts";

// ─── Types ────────────────────────────────────────────────────────────────
type ThemeName = "cinematic" | "modern" | "dark" | "professional";
type ChartType = "bar" | "line" | "pie";

interface DataRow {
  [key: string]: string | number;
}

// ─── Theme Definitions ───────────────────────────────────────────────────
const THEME_PALETTES: Record<ThemeName, string[]> = {
  cinematic: ["#e8a040", "#d4622a", "#f0c060", "#c2956b", "#8b5e3c", "#f5d89a"],
  modern: ["#06b6d4", "#3b82f6", "#8b5cf6", "#ec4899", "#14b8a6", "#f59e0b"],
  dark: ["#a78bfa", "#c084fc", "#e879f9", "#f472b6", "#818cf8", "#67e8f9"],
  professional: ["#3b82f6", "#1d4ed8", "#60a5fa", "#2563eb", "#93c5fd", "#1e40af"],
};

const THEME_INFO: Record<ThemeName, { label: string; icon: string; desc: string }> = {
  cinematic: { label: "Cinematic", icon: "🎬", desc: "Warm & dramatic tones" },
  modern: { label: "Modern", icon: "✨", desc: "Vibrant & fresh palette" },
  dark: { label: "Dark Mode", icon: "🌙", desc: "Sleek & minimal design" },
  professional: { label: "Professional", icon: "💼", desc: "Corporate & clean style" },
};

// ─── Utility Functions ──────────────────────────────────────────────────
function detectNumericColumns(data: DataRow[]): string[] {
  if (data.length === 0) return [];
  const keys = Object.keys(data[0]);
  return keys.filter((key) => {
    const sample = data.slice(0, 10);
    return sample.every((row) => {
      const val = row[key];
      return typeof val === "number" || (!isNaN(Number(val)) && val !== "");
    });
  });
}

function detectLabelColumn(data: DataRow[], numericCols: string[]): string {
  const keys = Object.keys(data[0] || {});
  const nonNumeric = keys.filter((k) => !numericCols.includes(k));
  return nonNumeric[0] || keys[0] || "label";
}

function generateInsights(data: DataRow[], numericCols: string[], labelCol: string): string[] {
  if (data.length === 0 || numericCols.length === 0) return [];

  const insights: string[] = [];

  insights.push(
    `📊 Dataset contains ${data.length} records with ${Object.keys(data[0]).length} columns. ` +
    `Found ${numericCols.length} numeric metric(s) suitable for visualization.`
  );

  for (const col of numericCols.slice(0, 3)) {
    const values = data.map((r) => Number(r[col]) || 0);
    const total = values.reduce((s, v) => s + v, 0);
    const avg = total / values.length;
    const max = Math.max(...values);
    const min = Math.min(...values);
    const maxIndex = values.indexOf(max);
    const minIndex = values.indexOf(min);

    insights.push(
      `📈 "${col}" ranges from ${min.toLocaleString()} to ${max.toLocaleString()} ` +
      `(avg: ${avg.toFixed(1)}). Highest at "${data[maxIndex]?.[labelCol] || `row ${maxIndex + 1}`}", ` +
      `lowest at "${data[minIndex]?.[labelCol] || `row ${minIndex + 1}`}".`
    );

    if (values.length >= 3) {
      const firstHalf = values.slice(0, Math.floor(values.length / 2));
      const secondHalf = values.slice(Math.floor(values.length / 2));
      const avgFirst = firstHalf.reduce((s, v) => s + v, 0) / firstHalf.length;
      const avgSecond = secondHalf.reduce((s, v) => s + v, 0) / secondHalf.length;
      const change = ((avgSecond - avgFirst) / avgFirst) * 100;

      if (Math.abs(change) > 5) {
        insights.push(
          `${change > 0 ? "🔺" : "🔻"} "${col}" shows a ${Math.abs(change).toFixed(1)}% ` +
          `${change > 0 ? "increase" : "decrease"} trend from the first half to the second half of data.`
        );
      }
    }
  }

  if (numericCols.length >= 2) {
    const col1Values = data.map((r) => Number(r[numericCols[0]]) || 0);
    const col2Values = data.map((r) => Number(r[numericCols[1]]) || 0);
    const total1 = col1Values.reduce((s, v) => s + v, 0);
    const total2 = col2Values.reduce((s, v) => s + v, 0);
    insights.push(
      `💡 Total "${numericCols[0]}": ${total1.toLocaleString()} vs Total "${numericCols[1]}": ${total2.toLocaleString()}. ` +
      `Ratio is ${(total1 / total2).toFixed(2)}.`
    );
  }

  insights.push(
    `🎯 Recommendation: Use Bar Chart for comparing categories, Line Chart for tracking trends over time, ` +
    `and Pie Chart for showing proportional distribution of a single metric.`
  );

  return insights;
}

function parseTextReport(text: string): DataRow[] {
  const lines = text.trim().split("\n").filter((l) => l.trim());
  if (lines.length < 2) return [];

  const data: DataRow[] = [];
  const numberPattern = /[\d,]+\.?\d*/g;
  const wordPattern = /[a-zA-Z\u00C0-\u024F\u0400-\u04FF]+[\w\s]*/g;

  for (const line of lines) {
    const numbers = line.match(numberPattern)?.map((n) => parseFloat(n.replace(/,/g, ""))) || [];
    const words = line.match(wordPattern)?.map((w) => w.trim()).filter((w) => w.length > 1) || [];

    if (numbers.length > 0 && words.length > 0) {
      const row: DataRow = { Label: words[0] };
      numbers.forEach((num, i) => {
        row[`Metric ${i + 1}`] = num;
      });
      data.push(row);
    }
  }

  return data;
}

// ─── Main Component ─────────────────────────────────────────────────────
export default function Home() {
  const [data, setData] = useState<DataRow[]>([]);
  const [fileName, setFileName] = useState<string>("");
  const [theme, setTheme] = useState<ThemeName>("modern");
  const [activeChart, setActiveChart] = useState<ChartType>("bar");
  const [isDragOver, setIsDragOver] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedMetric, setSelectedMetric] = useState<string>("");
  const dashboardRef = useRef<HTMLDivElement>(null);

  const numericCols = useMemo(() => detectNumericColumns(data), [data]);
  const labelCol = useMemo(() => detectLabelColumn(data, numericCols), [data, numericCols]);
  const insights = useMemo(
    () => generateInsights(data, numericCols, labelCol),
    [data, numericCols, labelCol]
  );

  const currentMetric = selectedMetric || numericCols[0] || "";
  const colors = THEME_PALETTES[theme];

  // ─── File Processing ────────────────────────────────────────────────
  const processFile = useCallback(async (file: File) => {
    setIsProcessing(true);
    setFileName(file.name);

    try {
      const ext = file.name.split(".").pop()?.toLowerCase();

      if (ext === "csv") {
        const text = await file.text();
        Papa.parse<DataRow>(text, {
          header: true,
          dynamicTyping: true,
          skipEmptyLines: true,
          complete: (results) => {
            setData(results.data);
            setSelectedMetric("");
            setIsProcessing(false);
          },
        });
      } else if (ext === "xlsx" || ext === "xls") {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json<DataRow>(sheet);
        setData(jsonData);
        setSelectedMetric("");
        setIsProcessing(false);
      } else if (ext === "txt") {
        const text = await file.text();
        const parsed = parseTextReport(text);
        if (parsed.length > 0) {
          setData(parsed);
          setSelectedMetric("");
        }
        setIsProcessing(false);
      } else {
        alert("Unsupported file format. Please upload CSV, Excel (.xlsx/.xls), or TXT files.");
        setIsProcessing(false);
      }
    } catch {
      alert("Error processing file. Please check the format and try again.");
      setIsProcessing(false);
    }
  }, []);

  // ─── Drag & Drop Handlers ──────────────────────────────────────────
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) processFile(file);
    },
    [processFile]
  );

  // ─── Export Functions ─────────────────────────────────────────────
  const exportAsImage = useCallback(async (format: "png" | "jpeg") => {
    if (!dashboardRef.current) return;

    const html2canvas = (await import("html2canvas-pro")).default;
    const canvas = await html2canvas(dashboardRef.current, {
      backgroundColor: "#0b0f1a",
      scale: 2,
      useCORS: true,
    });

    const link = document.createElement("a");
    link.download = `infographic-dashboard.${format}`;
    link.href = canvas.toDataURL(`image/${format}`, 0.95);
    link.click();
  }, []);

  const exportAsPDF = useCallback(async () => {
    if (!dashboardRef.current) return;

    const html2canvas = (await import("html2canvas-pro")).default;
    const { jsPDF } = await import("jspdf");

    const canvas = await html2canvas(dashboardRef.current, {
      backgroundColor: "#0b0f1a",
      scale: 2,
      useCORS: true,
    });

    const imgData = canvas.toDataURL("image/png");
    const pdf = new jsPDF("landscape", "mm", "a4");
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;

    pdf.addImage(imgData, "PNG", 0, 0, pdfWidth, pdfHeight);
    pdf.save("infographic-dashboard.pdf");
  }, []);

  // ─── Pie Data ─────────────────────────────────────────────────────
  const pieData = useMemo(() => {
    if (!currentMetric || data.length === 0) return [];
    return data.slice(0, 8).map((row) => ({
      name: String(row[labelCol] || ""),
      value: Number(row[currentMetric]) || 0,
    }));
  }, [data, currentMetric, labelCol]);

  // ─── Chart Data (limited for readability) ─────────────────────────
  const chartData = useMemo(() => {
    return data.slice(0, 20);
  }, [data]);

  // ─── Custom Tooltip ───────────────────────────────────────────────
  const CustomTooltip = ({ active, payload, label }: { active?: boolean; payload?: Array<{ name: string; value: number; color: string }>; label?: string }) => {
    if (!active || !payload?.length) return null;
    return (
      <div className="glass-card p-3 !rounded-lg text-sm">
        <p className="font-semibold mb-1" style={{ color: "var(--foreground)" }}>{label}</p>
        {payload.map((entry, i) => (
          <p key={i} style={{ color: entry.color }} className="text-xs">
            {entry.name}: {Number(entry.value).toLocaleString()}
          </p>
        ))}
      </div>
    );
  };

  // ─── Render ─────────────────────────────────────────────────────
  return (
    <div className={`bg-animated min-h-screen theme-${theme}`}>
      {/* ─── Header ─────────────────────────────────────────── */}
      <header className="px-6 py-5 flex items-center justify-between border-b border-[var(--card-border)] backdrop-blur-md bg-[var(--background)]/60 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center text-lg font-bold shadow-lg shadow-indigo-500/20">
            AI
          </div>
          <div>
            <h1 className="text-lg font-bold tracking-tight">
              <span className="gradient-text">Infographic & Dashboard</span>{" "}
              <span className="text-[var(--foreground)]">Generator</span>
            </h1>
            <p className="text-xs text-[var(--text-muted)]">
              AI-powered data visualization
            </p>
          </div>
        </div>

        {/* Theme Selector */}
        <div className="flex items-center gap-2">
          {(Object.keys(THEME_INFO) as ThemeName[]).map((t) => (
            <button
              key={t}
              id={`theme-${t}`}
              onClick={() => setTheme(t)}
              className={`
                px-3 py-2 rounded-xl text-xs font-medium transition-all duration-300
                ${
                  theme === t
                    ? "tab-active"
                    : "bg-[var(--glass-bg)] border border-[var(--glass-border)] hover:border-[var(--accent-primary)] text-[var(--text-muted)] hover:text-[var(--foreground)]"
                }
              `}
              title={THEME_INFO[t].desc}
            >
              <span className="mr-1">{THEME_INFO[t].icon}</span>
              <span className="hidden sm:inline">{THEME_INFO[t].label}</span>
            </button>
          ))}
        </div>
      </header>

      <main className="max-w-[1440px] mx-auto px-4 sm:px-6 py-8 space-y-8">
        {/* ─── Upload Section ────────────────────────────────── */}
        <section className="fade-in-up" id="upload-section">
          <div
            className={`drop-zone p-8 sm:p-12 text-center cursor-pointer ${
              isDragOver ? "drag-over" : ""
            }`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => document.getElementById("file-input")?.click()}
          >
            <input
              type="file"
              id="file-input"
              className="hidden"
              accept=".csv,.xlsx,.xls,.txt"
              onChange={handleFileInput}
            />

            {isProcessing ? (
              <div className="flex flex-col items-center gap-4">
                <div className="w-16 h-16 border-4 border-[var(--accent-primary)] border-t-transparent rounded-full spin-slow" />
                <p className="text-[var(--accent-primary)] font-semibold text-lg">
                  Processing your data...
                </p>
              </div>
            ) : (
              <>
                <div className="mx-auto w-20 h-20 rounded-2xl bg-gradient-to-br from-[var(--accent-primary)]/10 to-[var(--accent-secondary)]/10 border border-[var(--accent-primary)]/20 flex items-center justify-center mb-6">
                  <svg className="w-10 h-10 text-[var(--accent-primary)]" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                </div>
                <h2 className="text-xl font-bold text-[var(--foreground)] mb-2">
                  {fileName
                    ? `Loaded: ${fileName}`
                    : "Drag & Drop Your Data File"}
                </h2>
                <p className="text-[var(--text-muted)] text-sm mb-4">
                  Supports <span className="text-[var(--accent-primary)] font-medium">Excel (.xlsx)</span>,{" "}
                  <span className="text-[var(--accent-secondary)] font-medium">CSV</span>, and{" "}
                  <span className="text-[var(--accent-tertiary)] font-medium">Text Report (.txt)</span>
                </p>
                <div className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl bg-[var(--accent-primary)]/10 border border-[var(--accent-primary)]/20 text-[var(--accent-primary)] text-sm font-medium hover:bg-[var(--accent-primary)]/20 transition-all">
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
                  </svg>
                  Browse Files
                </div>
              </>
            )}
          </div>

          {/* Data Preview */}
          {data.length > 0 && (
            <div className="mt-4 glass-card p-4 fade-in-up">
              <div className="flex items-center justify-between mb-3">
                <h3 className="text-sm font-semibold text-[var(--foreground)] flex items-center gap-2">
                  <span className="status-badge success">
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 inline-block" />
                    Data Ready
                  </span>
                  <span className="text-[var(--text-muted)]">
                    {data.length} rows · {Object.keys(data[0]).length} columns
                  </span>
                </h3>
                <button
                  onClick={() => {
                    setData([]);
                    setFileName("");
                    setSelectedMetric("");
                  }}
                  className="text-xs text-red-400 hover:text-red-300 transition-colors"
                >
                  Clear Data
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead>
                    <tr className="border-b border-[var(--card-border)]">
                      {Object.keys(data[0]).map((key) => (
                        <th
                          key={key}
                          className="text-left py-2 px-3 text-[var(--text-muted)] font-medium"
                        >
                          {key}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {data.slice(0, 5).map((row, i) => (
                      <tr
                        key={i}
                        className="border-b border-[var(--card-border)]/50 hover:bg-white/[0.02] transition-colors"
                      >
                        {Object.values(row).map((val, j) => (
                          <td
                            key={j}
                            className="py-2 px-3 text-[var(--foreground)]"
                          >
                            {String(val)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {data.length > 5 && (
                  <p className="text-[var(--text-muted)] text-xs mt-2 text-center">
                    ...and {data.length - 5} more rows
                  </p>
                )}
              </div>
            </div>
          )}
        </section>

        {/* ─── Dashboard Section ─────────────────────────────── */}
        {data.length > 0 && numericCols.length > 0 && (
          <>
            <div ref={dashboardRef} className="space-y-6" id="dashboard-section">
              {/* Controls */}
              <section className="fade-in-up fade-in-up-delay-1">
                <div className="glass-card p-5">
                  <div className="flex flex-wrap items-center justify-between gap-4">
                    {/* Chart Type Selector */}
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-[var(--text-muted)] font-medium uppercase tracking-wider mr-2">
                        Chart
                      </span>
                      {(["bar", "line", "pie"] as ChartType[]).map((type) => (
                        <button
                          key={type}
                          id={`chart-${type}`}
                          onClick={() => setActiveChart(type)}
                          className={`
                            px-4 py-2 rounded-lg text-sm font-medium transition-all duration-300
                            ${
                              activeChart === type
                                ? "tab-active"
                                : "bg-[var(--glass-bg)] border border-[var(--glass-border)] text-[var(--text-muted)] hover:text-[var(--foreground)] hover:border-[var(--accent-primary)]/50"
                            }
                          `}
                        >
                          {type === "bar" && "📊 "}
                          {type === "line" && "📈 "}
                          {type === "pie" && "🥧 "}
                          {type.charAt(0).toUpperCase() + type.slice(1)}
                        </button>
                      ))}
                    </div>

                    {/* Metric Selector */}
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-[var(--text-muted)] font-medium uppercase tracking-wider mr-2">
                        Metric
                      </span>
                      <select
                        id="metric-selector"
                        value={currentMetric}
                        onChange={(e) => setSelectedMetric(e.target.value)}
                        className="bg-[var(--glass-bg)] border border-[var(--glass-border)] rounded-lg px-4 py-2 text-sm text-[var(--foreground)] focus:outline-none focus:border-[var(--accent-primary)] transition-colors cursor-pointer"
                      >
                        {numericCols.map((col) => (
                          <option
                            key={col}
                            value={col}
                            className="bg-[var(--background)] text-[var(--foreground)]"
                          >
                            {col}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
              </section>

              {/* Charts */}
              <section className="fade-in-up fade-in-up-delay-2">
                <div className="glass-card p-6 pulse-glow">
                  <h2 className="text-base font-bold text-[var(--foreground)] mb-1 flex items-center gap-2">
                    <span className="w-1.5 h-6 rounded-full bg-gradient-to-b from-[var(--accent-primary)] to-[var(--accent-secondary)]" />
                    AI Visual Designer
                  </h2>
                  <p className="text-xs text-[var(--text-muted)] mb-6">
                    Automatically selected the best visualization for your data
                  </p>

                  <div className="chart-container bg-[var(--glass-bg)] p-4 rounded-xl border border-[var(--glass-border)]">
                    {activeChart === "bar" && (
                      <ResponsiveContainer width="100%" height={400}>
                        <BarChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                          <XAxis
                            dataKey={labelCol}
                            tick={{ fill: "var(--text-muted)", fontSize: 11 }}
                            tickLine={false}
                            axisLine={{ stroke: "rgba(255,255,255,0.1)" }}
                            angle={-30}
                            textAnchor="end"
                            height={60}
                          />
                          <YAxis
                            tick={{ fill: "var(--text-muted)", fontSize: 11 }}
                            tickLine={false}
                            axisLine={{ stroke: "rgba(255,255,255,0.1)" }}
                          />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: "12px", paddingTop: "16px" }} />
                          {numericCols.slice(0, 4).map((col, i) => (
                            <Bar
                              key={col}
                              dataKey={col}
                              fill={colors[i % colors.length]}
                              radius={[6, 6, 0, 0]}
                              fillOpacity={0.85}
                            />
                          ))}
                        </BarChart>
                      </ResponsiveContainer>
                    )}

                    {activeChart === "line" && (
                      <ResponsiveContainer width="100%" height={400}>
                        <AreaChart data={chartData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                          <defs>
                            {numericCols.slice(0, 4).map((col, i) => (
                              <linearGradient key={col} id={`gradient-${col}`} x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor={colors[i % colors.length]} stopOpacity={0.3} />
                                <stop offset="95%" stopColor={colors[i % colors.length]} stopOpacity={0} />
                              </linearGradient>
                            ))}
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.06)" />
                          <XAxis
                            dataKey={labelCol}
                            tick={{ fill: "var(--text-muted)", fontSize: 11 }}
                            tickLine={false}
                            axisLine={{ stroke: "rgba(255,255,255,0.1)" }}
                            angle={-30}
                            textAnchor="end"
                            height={60}
                          />
                          <YAxis
                            tick={{ fill: "var(--text-muted)", fontSize: 11 }}
                            tickLine={false}
                            axisLine={{ stroke: "rgba(255,255,255,0.1)" }}
                          />
                          <Tooltip content={<CustomTooltip />} />
                          <Legend wrapperStyle={{ fontSize: "12px", paddingTop: "16px" }} />
                          {numericCols.slice(0, 4).map((col, i) => (
                            <React.Fragment key={col}>
                              <Area
                                type="monotone"
                                dataKey={col}
                                stroke={colors[i % colors.length]}
                                fill={`url(#gradient-${col})`}
                                strokeWidth={2.5}
                              />
                            </React.Fragment>
                          ))}
                        </AreaChart>
                      </ResponsiveContainer>
                    )}

                    {activeChart === "pie" && (
                      <ResponsiveContainer width="100%" height={400}>
                        <PieChart>
                          <Pie
                            data={pieData}
                            cx="50%"
                            cy="50%"
                            innerRadius={80}
                            outerRadius={150}
                            paddingAngle={3}
                            dataKey="value"
                            label={({ name, percent }) =>
                              `${name} (${(percent * 100).toFixed(0)}%)`
                            }
                            labelLine={{ stroke: "rgba(255,255,255,0.2)" }}
                          >
                            {pieData.map((_, i) => (
                              <Cell
                                key={i}
                                fill={colors[i % colors.length]}
                                fillOpacity={0.85}
                                stroke="rgba(0,0,0,0.3)"
                                strokeWidth={1}
                              />
                            ))}
                          </Pie>
                          <Tooltip />
                          <Legend wrapperStyle={{ fontSize: "12px", paddingTop: "16px" }} />
                        </PieChart>
                      </ResponsiveContainer>
                    )}
                  </div>
                </div>
              </section>

              {/* ─── AI Insights ──────────────────────────────────── */}
              <section className="fade-in-up fade-in-up-delay-3" id="insights-section">
                <div className="glass-card p-6">
                  <h2 className="text-base font-bold text-[var(--foreground)] mb-1 flex items-center gap-2">
                    <span className="w-1.5 h-6 rounded-full bg-gradient-to-b from-[var(--accent-tertiary)] to-[var(--accent-primary)]" />
                    AI Caption & Summary
                  </h2>
                  <p className="text-xs text-[var(--text-muted)] mb-5">
                    Auto-generated insights for your presentation materials
                  </p>

                  <div className="grid gap-3 sm:grid-cols-2">
                    {insights.map((insight, i) => (
                      <div key={i} className="insight-card">
                        <p className="text-sm text-[var(--foreground)] leading-relaxed">
                          {insight}
                        </p>
                      </div>
                    ))}
                  </div>
                </div>
              </section>
            </div>

            {/* ─── Export Section ─────────────────────────────── */}
            <section className="fade-in-up fade-in-up-delay-4" id="export-section">
              <div className="glass-card p-6">
                <h2 className="text-base font-bold text-[var(--foreground)] mb-1 flex items-center gap-2">
                  <span className="w-1.5 h-6 rounded-full bg-gradient-to-b from-emerald-400 to-cyan-500" />
                  Export & Download
                </h2>
                <p className="text-xs text-[var(--text-muted)] mb-5">
                  One-click export in multiple formats
                </p>

                <div className="flex flex-wrap gap-3">
                  <button
                    id="export-png"
                    onClick={() => exportAsImage("png")}
                    className="export-btn"
                  >
                    <svg className="w-5 h-5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                    </svg>
                    Download PNG
                  </button>

                  <button
                    id="export-jpg"
                    onClick={() => exportAsImage("jpeg")}
                    className="export-btn"
                  >
                    <svg className="w-5 h-5 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z" />
                    </svg>
                    Download JPG
                  </button>

                  <button
                    id="export-pdf"
                    onClick={exportAsPDF}
                    className="export-btn"
                  >
                    <svg className="w-5 h-5 text-red-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
                    </svg>
                    Download PDF
                  </button>
                </div>
              </div>
            </section>
          </>
        )}

        {/* ─── Empty State ───────────────────────────────────── */}
        {data.length === 0 && !isProcessing && (
          <section className="text-center py-16 fade-in-up fade-in-up-delay-2">
            <div className="max-w-md mx-auto">
              <div className="grid grid-cols-3 gap-4 mb-10">
                {[
                  { icon: "📊", title: "Bar Chart", desc: "Compare categories" },
                  { icon: "📈", title: "Line Chart", desc: "Track trends" },
                  { icon: "🥧", title: "Pie Chart", desc: "Show proportions" },
                ].map((item) => (
                  <div
                    key={item.title}
                    className="glass-card p-4 text-center"
                  >
                    <div className="text-3xl mb-2">{item.icon}</div>
                    <p className="text-sm font-semibold text-[var(--foreground)]">
                      {item.title}
                    </p>
                    <p className="text-xs text-[var(--text-muted)]">
                      {item.desc}
                    </p>
                  </div>
                ))}
              </div>
              <p className="text-[var(--text-muted)] text-sm">
                Upload your data to get started with AI-powered visualizations
              </p>
            </div>
          </section>
        )}
      </main>

      {/* ─── Footer ──────────────────────────────────────────── */}
      <footer className="mt-auto border-t border-[var(--card-border)] py-6 px-6 backdrop-blur-md bg-[var(--background)]/40">
        <div className="max-w-[1440px] mx-auto flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-[var(--text-muted)]">
          <p>
            © {new Date().getFullYear()} AI Infographic & Dashboard Generator.
            Built with Next.js & Recharts.
          </p>
          <div className="flex items-center gap-4">
            <span className="status-badge processing">
              <span className="w-1.5 h-1.5 rounded-full bg-indigo-400 inline-block animate-pulse" />
              AI-Powered
            </span>
          </div>
        </div>
      </footer>
    </div>
  );
}
