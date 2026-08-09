import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  Plus, Pencil, Palette, Trash2, Archive, ArchiveRestore, LayoutGrid,
  Printer, Search, ArrowRight, X, Check, Minus, Hash, Type,
  ListChecks, FolderClock, BookOpen, FileText, RefreshCw, ClipboardList,
  Pin, PinOff, Copy, RotateCcw, FolderOpen, FileImage, FileSpreadsheet, ListOrdered,
  Share2, Calendar, CalendarCheck, Newspaper, Eraser, CalendarRange,
  Lock, Unlock, ChevronUp, ChevronDown, ChevronLeft, ChevronRight, ImageDown, FileOutput,
  Camera, ImageOff, Settings, Volume2, VolumeX, BarChart3, Users,
  Shuffle, AlertTriangle, MessageSquareWarning, ClipboardCopy, Eye, EyeOff, Award,
  CalendarPlus, Moon, Sun, Filter, ListTodo, HelpCircle, Send, Activity
} from "lucide-react";

// Anon/public key — safe to keep in client code by design (Supabase protects
// data via Row Level Security policies, not by hiding this key). NEVER put
// the service_role/secret key here.
const SUPABASE_URL = "https://layvtwynhrkakqagrjcv.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxheXZ0d3luaHJrYWtxYWdyamN2Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODYyNzQ3NzUsImV4cCI6MjEwMTg1MDc3NX0.bXpib2_qXKlSh2ps4vy30tVhjNqyPMtgB87iR0zvpRA";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const STORAGE_KEY = "mutabaa-app-data-v2";

const COLORS = [
  { name: "أخضر مسجدي", hex: "#0F6B5C", light: "#E3F0ED" },
  { name: "كهرماني", hex: "#C97A2B", light: "#F7E8D8" },
  { name: "نيلي", hex: "#3B4C8C", light: "#E1E4EF" },
  { name: "وردي مطفأ", hex: "#B4526A", light: "#F3E1E6" },
  { name: "زيتي", hex: "#6B7A3A", light: "#EAEDDD" },
  { name: "سماوي", hex: "#2E7DA6", light: "#DCEAF1" },
  { name: "أرجواني", hex: "#7A4E9E", light: "#EBE1F0" },
  { name: "رمادي", hex: "#5B6472", light: "#E5E6E8" },
  { name: "أحمر", hex: "#C0392B", light: "#F5DEDB" },
];
function colorLight(hex) {
  return COLORS.find((c) => c.hex === hex)?.light || "#F3F1E9";
}

const CLASS_EMOJIS = ["📐", "📖", "🔬", "🎨", "⚽", "🕌", "🌍", "💻", "✏️", "📊", "🎵", "🧮", "🔤", "🌱", "⚗️", "📚"];

const INK = "#232622";
const PAPER = "#FAF8F3";
const LINE = "#E4DFD2";
const MUTED = "#7A7768";

function uid() {
  return crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2);
}
function shuffleArr(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}
function nowLabel() {
  try {
    return new Date().toLocaleString("ar-SA-u-nu-latn", { dateStyle: "medium", timeStyle: "short" });
  } catch (e) {
    return new Date().toLocaleString();
  }
}
function nowMeta() {
  const d = new Date();
  const dateKey = d.toISOString().slice(0, 10);
  try {
    return {
      day: d.toLocaleDateString("ar-SA-u-nu-latn", { weekday: "long" }),
      date: d.toLocaleDateString("ar-SA-u-nu-latn", { day: "numeric", month: "long", year: "numeric" }),
      time: d.toLocaleTimeString("ar-SA-u-nu-latn", { hour: "2-digit", minute: "2-digit" }),
      dateKey,
    };
  } catch (e) {
    return { day: "", date: d.toLocaleDateString(), time: d.toLocaleTimeString(), dateKey };
  }
}
// Groups flat report entries into one section per column, preserving first-appearance order.
function groupEntries(entries) {
  const map = new Map();
  entries.forEach((e) => {
    const key = e.colId || "note";
    if (!map.has(key)) map.set(key, { colId: key, colName: e.colName, colColor: e.colColor, items: [] });
    map.get(key).items.push(e);
  });
  return Array.from(map.values());
}
// Compiles a ready-to-copy behavior report draft from one flagged category's entries.
function buildBehaviorDraft(cls, row, group) {
  const lines = group.items.map((e, i) => `${i + 1}. ${e.day ? `${e.day}، ` : ""}${e.date || ""} — ${e.time || ""}: ${e.value}`);
  const today = nowMeta();
  return [
    "مسودة تقرير سلوكي",
    "",
    `الطالب: ${row.name}`,
    `الفصل: ${cls.subject} — ${cls.grade}`,
    `المعلم: ${cls.teacher}`,
    `التصنيف: ${group.colName}`,
    `عدد الملاحظات المسجّلة: ${group.items.length}`,
    "",
    "تفاصيل الملاحظات:",
    ...lines,
    "",
    "يُرجى من إدارة المدرسة / ولي أمر الطالب الاطلاع على ما سبق واتخاذ الإجراء المناسب.",
    "",
    `تاريخ إعداد التقرير: ${today.day}، ${today.date}`,
    `المعلم: ${cls.teacher}`,
  ].join("\n");
}
// Returns the most recent recorded value for a row+column from the report log,
// used so the display board keeps showing scores even after a cell auto-clears.
function lastReportedValue(cls, rowId, colId) {
  const list = cls.reports?.[rowId] || [];
  for (let i = list.length - 1; i >= 0; i--) {
    if (list[i].colId === colId) return list[i].value;
  }
  return null;
}
// Returns the value recorded for a row+column on one specific calendar date
// (used by لوحة العرض once a date is picked), or null if nothing was recorded that day.
function valueOnDate(cls, rowId, colId, dateKey) {
  const list = cls.reports?.[rowId] || [];
  for (let i = list.length - 1; i >= 0; i--) {
    if (list[i].colId === colId && list[i].dateKey === dateKey) return list[i].value;
  }
  return null;
}
// Shared cell-value resolver for the display board / exports: date-scoped
// when a specific date is given, otherwise falls back to the live cell / last-known value.
function boardCellValue(cls, rowId, colId, dateKey) {
  if (dateKey) return valueOnDate(cls, rowId, colId, dateKey) || "";
  return cls.cells[`${rowId}:${colId}`] || lastReportedValue(cls, rowId, colId) || "";
}

// The mandatory "الغياب" column is not a normal user-defined column — it's a
// fixed, always-present part of every class table (as soon as it has rows),
// similar to the report/delete columns. A student with no explicit "absent"
// mark for a date is simply considered present by default.
const ATTEND_COL_ID = "__absence__";
function todayKey() {
  return new Date().toISOString().slice(0, 10);
}
// Short confirmation tone + light vibration when a record is saved. Silent no-op
// if the person has disabled it in settings, or if the browser blocks audio
// (e.g. no user gesture yet) — never throws, just skips quietly.
function playFeedback(enabled) {
  if (!enabled) return;
  try {
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (AudioCtx) {
      const ctx = new AudioCtx();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = "sine";
      osc.frequency.value = 880;
      gain.gain.setValueAtTime(0.16, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.16);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start();
      osc.stop(ctx.currentTime + 0.16);
    }
  } catch (e) { /* ignore */ }
  try {
    if (navigator.vibrate) navigator.vibrate(35);
  } catch (e) { /* ignore */ }
}
function attendanceStatus(cls, rowId, dateKey) {
  return cls.attendance?.[dateKey]?.[rowId] === "absent" ? "absent" : "present";
}
// Pure state-updater factories (used with updateClass both from the class
// table's quick-mark button and from the متابعة الحضور modal).
function markAbsentUpdater(rowId, dateKey) {
  return (c) => {
    const attendance = { ...(c.attendance || {}), [dateKey]: { ...((c.attendance || {})[dateKey] || {}), [rowId]: "absent" } };
    const already = (c.reports?.[rowId] || []).some((e) => e.colId === ATTEND_COL_ID && e.dateKey === dateKey);
    let reports = c.reports || {};
    if (!already) {
      const meta = nowMeta();
      const entry = { id: uid(), colId: ATTEND_COL_ID, colName: "الغياب", colColor: "#C0392B", value: "غائب", dateKey, ...meta };
      reports = { ...reports, [rowId]: [...(reports[rowId] || []), entry] };
    }
    return { ...c, attendance, reports };
  };
}
function clearAbsentUpdater(rowId, dateKey) {
  return (c) => {
    const dayRec = { ...((c.attendance || {})[dateKey] || {}) };
    delete dayRec[rowId];
    const reports = { ...(c.reports || {}) };
    if (reports[rowId]) reports[rowId] = reports[rowId].filter((e) => !(e.colId === ATTEND_COL_ID && e.dateKey === dateKey));
    return { ...c, attendance: { ...(c.attendance || {}), [dateKey]: dayRec }, reports };
  };
}
function wrapCanvasText(ctx, text, maxWidth) {
  const raw = String(text ?? "").trim();
  if (!raw) return ["—"];
  const words = raw.split(/\s+/);
  const lines = [];
  let line = "";
  const breakLongWord = (word) => {
    // Force-splits a single token that's wider than the cell on its own,
    // so it can never overflow past the cell border into a neighboring cell/column.
    let chunk = "";
    for (const ch of word) {
      const test = chunk + ch;
      if (chunk && ctx.measureText(test).width > maxWidth) {
        lines.push(chunk);
        chunk = ch;
      } else {
        chunk = test;
      }
    }
    return chunk;
  };
  words.forEach((w) => {
    const test = line ? `${line} ${w}` : w;
    if (ctx.measureText(test).width <= maxWidth) {
      line = test;
      return;
    }
    if (line) { lines.push(line); line = ""; }
    line = ctx.measureText(w).width <= maxWidth ? w : breakLongWord(w);
  });
  if (line) lines.push(line);
  return lines.length ? lines : ["—"];
}

// Builds a flat table onto an offscreen canvas (title + header row + data
// rows), wrapping long cell text onto multiple lines instead of letting it
// overflow, and rendering at high resolution for crisp PNG/PDF output.
function buildTableCanvas({ title, subtitle, headers, rows }) {
  const cellW = 150, lineH = 18, cellPadV = 12, headerH = 42, pad = 24, titleH = 70;
  const cols = headers.length;
  const width = pad * 2 + cols * cellW;
  const scale = 3;
  const measure = document.createElement("canvas").getContext("2d");
  measure.font = "13px Tahoma, Arial";
  const rowLines = rows.map((r) => r.map((val) => wrapCanvasText(measure, val ? String(val) : "—", cellW - 16)));
  const rowHeights = rowLines.map((cellsLines) => Math.max(...cellsLines.map((lines) => lines.length)) * lineH + cellPadV);
  const height = titleH + headerH + rowHeights.reduce((a, b) => a + b, 0) + pad;
  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = PAPER;
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = INK;
  ctx.font = "bold 20px Tahoma, Arial";
  ctx.fillText(title, width / 2, pad + 14);
  if (subtitle) {
    ctx.font = "13px Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText(subtitle, width / 2, pad + 36);
  }
  let y = titleH;
  ctx.fillStyle = "#F3F1E9";
  ctx.fillRect(pad, y, cols * cellW, headerH);
  headers.forEach((h, i) => {
    ctx.strokeStyle = LINE;
    ctx.strokeRect(pad + i * cellW, y, cellW, headerH);
    ctx.fillStyle = INK;
    ctx.font = "bold 14px Tahoma, Arial";
    ctx.fillText(String(h), pad + i * cellW + cellW / 2, y + headerH / 2);
  });
  y += headerH;
  rows.forEach((r, ri) => {
    const rh = rowHeights[ri];
    ctx.fillStyle = ri % 2 ? "#FBFAF6" : "#ffffff";
    ctx.fillRect(pad, y, cols * cellW, rh);
    r.forEach((val, ci) => {
      ctx.strokeStyle = LINE;
      ctx.strokeRect(pad + ci * cellW, y, cellW, rh);
      ctx.fillStyle = INK;
      ctx.font = "13px Tahoma, Arial";
      const lines = rowLines[ri][ci];
      let ly = y + rh / 2 - (lines.length * lineH) / 2 + lineH / 2;
      lines.forEach((ln) => { ctx.fillText(ln, pad + ci * cellW + cellW / 2, ly); ly += lineH; });
    });
    y += rh;
  });
  return { canvas, logicalWidth: width, logicalHeight: height };
}

// Builds a well-organized, sectioned canvas for a student report: one
// colored section per category (behavior, homework, participation, notes,
// exams...) each with its own mini date/time/value table — legible enough
// to print and hand to a parent.
function buildReportCanvas({ title, subtitle, groups }) {
  const pad = 24, titleH = 70, summaryH = 34, sectionH = 34, colHeaderH = 26, lineH = 17, cellPadV = 12, sectionGap = 18;
  const colWidths = [180, 90, 300];
  const tableW = colWidths.reduce((a, b) => a + b, 0);
  const width = pad * 2 + tableW;
  const scale = 3;
  const measure = document.createElement("canvas").getContext("2d");
  measure.font = "13px Tahoma, Arial";
  const groupRowLines = groups.map((g) => g.items.map((it) => wrapCanvasText(measure, it.value || "—", colWidths[2] - 20)));
  const groupRowHeights = groups.map((g, gi) => g.items.map((it, ii) => Math.max(1, groupRowLines[gi][ii].length) * lineH + cellPadV));
  let height = titleH + (groups.length > 0 ? summaryH : 0);
  groups.forEach((g, gi) => { height += sectionH + colHeaderH + groupRowHeights[gi].reduce((a, b) => a + b, 0) + sectionGap; });
  height += pad + (groups.length === 0 ? 40 : 0);
  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = PAPER;
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = INK;
  ctx.font = "bold 20px Tahoma, Arial";
  ctx.fillText(title, width / 2, pad + 14);
  if (subtitle) {
    ctx.font = "13px Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText(subtitle, width / 2, pad + 36);
  }
  let y = titleH;
  if (groups.length === 0) {
    ctx.font = "14px Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText("لا يوجد رصد لهذا الطالب بعد.", width / 2, y + 30);
  } else {
    // Summary strip: one chip per category with its record count
    ctx.font = "12px Tahoma, Arial";
    const summaryText = groups.map((g) => `${g.colName} (${g.items.length})`).join("   •   ");
    ctx.fillStyle = "#F3F1E9";
    ctx.fillRect(pad, y, tableW, summaryH);
    ctx.fillStyle = INK;
    ctx.font = "bold 12px Tahoma, Arial";
    ctx.fillText(summaryText, width / 2, y + summaryH / 2);
    y += summaryH;
  }
  groups.forEach((g, gi) => {
    ctx.fillStyle = g.colColor || "#0F6B5C";
    ctx.fillRect(pad, y, tableW, sectionH);
    ctx.fillStyle = "#fff";
    ctx.font = "bold 14px Tahoma, Arial";
    ctx.fillText(`${g.colName}  (${g.items.length})`, pad + tableW / 2, y + sectionH / 2);
    y += sectionH;
    const headers = ["اليوم والتاريخ", "الوقت", "القيمة"];
    ctx.fillStyle = "#F3F1E9";
    ctx.fillRect(pad, y, tableW, colHeaderH);
    let x = pad;
    headers.forEach((h, i) => {
      ctx.strokeStyle = LINE;
      ctx.strokeRect(x, y, colWidths[i], colHeaderH);
      ctx.fillStyle = INK;
      ctx.font = "bold 12px Tahoma, Arial";
      ctx.fillText(h, x + colWidths[i] / 2, y + colHeaderH / 2);
      x += colWidths[i];
    });
    y += colHeaderH;
    g.items.forEach((it, ri) => {
      const rh = groupRowHeights[gi][ri];
      ctx.fillStyle = ri % 2 ? "#FBFAF6" : "#fff";
      ctx.fillRect(pad, y, tableW, rh);
      const dayDate = `${it.day ? it.day + "، " : ""}${it.date || ""}`;
      let xx = pad;
      ctx.strokeStyle = LINE;
      ctx.strokeRect(xx, y, colWidths[0], rh);
      ctx.fillStyle = INK;
      ctx.font = "12px Tahoma, Arial";
      ctx.fillText(dayDate || "—", xx + colWidths[0] / 2, y + rh / 2);
      xx += colWidths[0];
      ctx.strokeRect(xx, y, colWidths[1], rh);
      ctx.fillText(it.time || "—", xx + colWidths[1] / 2, y + rh / 2);
      xx += colWidths[1];
      ctx.strokeRect(xx, y, colWidths[2], rh);
      const lines = groupRowLines[gi][ri];
      let ly = y + rh / 2 - (lines.length * lineH) / 2 + lineH / 2;
      lines.forEach((ln) => { ctx.fillText(ln, xx + colWidths[2] / 2, ly); ly += lineH; });
      y += rh;
    });
    y += sectionGap;
  });
  return { canvas, logicalWidth: width, logicalHeight: height };
}

// Draws a simple decorative certificate of appreciation onto a canvas.
function buildCertificateCanvas({ countryName, ministryName, schoolName, logoImageElement, title, studentName, reason, className, teacherName, principalName, date, accentColor }) {
  const width = 1000, height = 760, scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  const accent = accentColor || "#0F6B5C";

  // Warm gradient background instead of a flat fill
  const grad = ctx.createRadialGradient(width / 2, height / 2, 40, width / 2, height / 2, width * 0.72);
  grad.addColorStop(0, "#FFFDF8");
  grad.addColorStop(1, "#FBF3DE");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, width, height);

  // faint dot texture for depth
  ctx.save();
  ctx.globalAlpha = 0.05;
  ctx.fillStyle = accent;
  for (let x = 60; x < width - 60; x += 26) {
    for (let y = 60; y < height - 60; y += 26) {
      ctx.beginPath();
      ctx.arc(x, y, 1.1, 0, Math.PI * 2);
      ctx.fill();
    }
  }
  ctx.restore();

  // decorative border (outer solid, inner dashed, innermost thin solid)
  ctx.strokeStyle = accent;
  ctx.lineWidth = 6;
  ctx.strokeRect(24, 24, width - 48, height - 48);
  ctx.save();
  ctx.setLineDash([2, 6]);
  ctx.lineWidth = 1.5;
  ctx.strokeRect(33, 33, width - 66, height - 66);
  ctx.restore();
  ctx.lineWidth = 1.5;
  ctx.strokeRect(38, 38, width - 76, height - 76);

  // ornamental corner flourishes
  const drawCorner = (cx, cy, sx, sy) => {
    ctx.save();
    ctx.translate(cx, cy);
    ctx.scale(sx, sy);
    ctx.strokeStyle = accent;
    ctx.globalAlpha = 0.55;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.arc(0, 0, 30, Math.PI, 1.5 * Math.PI);
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(0, 0, 18, Math.PI, 1.5 * Math.PI);
    ctx.stroke();
    ctx.globalAlpha = 1;
    ctx.fillStyle = accent;
    ctx.beginPath();
    ctx.moveTo(34, 0); ctx.lineTo(40, 6); ctx.lineTo(46, 0); ctx.lineTo(40, -6);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
  };
  drawCorner(58, 58, 1, 1);
  drawCorner(width - 58, 58, -1, 1);
  drawCorner(58, height - 58, 1, -1);
  drawCorner(width - 58, height - 58, -1, -1);

  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  // Official header block: logo (if provided) + country/ministry/school names.
  const OFFSET = 60; // shifts everything below down to make room for this header
  if (logoImageElement) {
    const logoSize = 56;
    ctx.drawImage(logoImageElement, width / 2 - logoSize / 2, 52, logoSize, logoSize);
  }
  let hy = logoImageElement ? 122 : 62;
  if (countryName) {
    ctx.fillStyle = "#232622";
    ctx.font = "bold 15px Tahoma, Arial";
    ctx.fillText(countryName, width / 2, hy);
    hy += 20;
  }
  if (ministryName) {
    ctx.fillStyle = "#7A7768";
    ctx.font = "14px Tahoma, Arial";
    ctx.fillText(ministryName, width / 2, hy);
    hy += 20;
  }
  if (schoolName) {
    ctx.fillStyle = "#7A7768";
    ctx.font = "16px Tahoma, Arial";
    ctx.fillText(schoolName, width / 2, hy);
  }

  ctx.fillStyle = accent;
  ctx.font = "bold 44px Tahoma, Arial";
  ctx.fillText(title, width / 2, 160 + OFFSET);

  ctx.strokeStyle = accent;
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(width / 2 - 90, 190 + OFFSET);
  ctx.lineTo(width / 2 + 90, 190 + OFFSET);
  ctx.stroke();

  ctx.fillStyle = "#232622";
  ctx.font = "18px Tahoma, Arial";
  ctx.fillText("تُمنح هذه الشهادة إلى", width / 2, 250 + OFFSET);

  ctx.fillStyle = "#232622";
  ctx.font = "bold 40px Tahoma, Arial";
  ctx.fillText(studentName, width / 2, 320 + OFFSET);

  ctx.strokeStyle = "#D8D2C0";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(width / 2 - 220, 350 + OFFSET);
  ctx.lineTo(width / 2 + 220, 350 + OFFSET);
  ctx.stroke();

  const wrapped = wrapCanvasText(ctx, reason, width - 260);
  ctx.font = "20px Tahoma, Arial";
  ctx.fillStyle = "#3A3A34";
  let ry = 400 + OFFSET;
  wrapped.forEach((line) => { ctx.fillText(line, width / 2, ry); ry += 30; });

  if (className) {
    ctx.font = "15px Tahoma, Arial";
    ctx.fillStyle = "#7A7768";
    ctx.fillText(className, width / 2, ry + 20);
  }

  ctx.font = "14px Tahoma, Arial";
  ctx.fillStyle = "#7A7768";
  ctx.textAlign = "right";
  ctx.fillText(`التاريخ: ${date}`, width - 90, height - 90);

  ctx.textAlign = "center";
  const colW = 260;
  if (teacherName) {
    ctx.font = "12px Tahoma, Arial";
    ctx.fillStyle = "#7A7768";
    ctx.fillText("المعلم", width - 120 - colW / 2, height - 130);
    ctx.font = "bold 15px Tahoma, Arial";
    ctx.fillStyle = "#232622";
    ctx.fillText(teacherName, width - 120 - colW / 2, height - 108);
  }
  if (principalName) {
    ctx.font = "12px Tahoma, Arial";
    ctx.fillStyle = "#7A7768";
    ctx.fillText("مدير المدرسة", 120 + colW / 2, height - 130);
    ctx.font = "bold 15px Tahoma, Arial";
    ctx.fillStyle = "#232622";
    ctx.fillText(principalName, 120 + colW / 2, height - 108);
  }

  return { canvas, logicalWidth: width, logicalHeight: height };
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}

// Minimal ZIP writer (stored/uncompressed entries) — no external library
// needed, and "stored" ZIP entries are readable by every standard unzip tool.
function crc32(buf) {
  if (!crc32.table) {
    const table = [];
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    crc32.table = table;
  }
  let crc = 0 ^ -1;
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ crc32.table[(crc ^ buf[i]) & 0xff];
  return (crc ^ -1) >>> 0;
}
function buildZip(files) {
  const encoder = new TextEncoder();
  const now = new Date();
  const dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xffff;
  const dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xffff;
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBytes = encoder.encode(file.name);
    const data = file.data;
    const crc = crc32(data);
    const localHeaderOffset = offset;

    const local = new Uint8Array(30 + nameBytes.length);
    const ldv = new DataView(local.buffer);
    ldv.setUint32(0, 0x04034b50, true);
    ldv.setUint16(4, 20, true);
    ldv.setUint16(6, 0, true);
    ldv.setUint16(8, 0, true);
    ldv.setUint16(10, dosTime, true);
    ldv.setUint16(12, dosDate, true);
    ldv.setUint32(14, crc, true);
    ldv.setUint32(18, data.length, true);
    ldv.setUint32(22, data.length, true);
    ldv.setUint16(26, nameBytes.length, true);
    ldv.setUint16(28, 0, true);
    local.set(nameBytes, 30);
    localParts.push(local, data);
    offset += local.length + data.length;

    const central = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(central.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(4, 20, true);
    cdv.setUint16(6, 20, true);
    cdv.setUint16(8, 0, true);
    cdv.setUint16(10, 0, true);
    cdv.setUint16(12, dosTime, true);
    cdv.setUint16(14, dosDate, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, data.length, true);
    cdv.setUint32(24, data.length, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint32(42, localHeaderOffset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);
  });

  const centralOffset = offset;
  const centralSize = centralParts.reduce((s, p) => s + p.length, 0);
  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(8, files.length, true);
  edv.setUint16(10, files.length, true);
  edv.setUint32(12, centralSize, true);
  edv.setUint32(16, centralOffset, true);

  const parts = [...localParts, ...centralParts, eocd];
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  parts.forEach((p) => { out.set(p, pos); pos += p.length; });
  return new Blob([out], { type: "application/zip" });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

// Tries the native device share sheet first (great on mobile for "sharing" a
// file directly, e.g. to WhatsApp or a parent's number); falls back to a
// normal download if the Web Share API / file sharing isn't available.
async function shareOrDownloadFile(blob, filename, mime) {
  try {
    const file = new File([blob], filename, { type: mime });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: filename });
      return;
    }
  } catch (e) {
    // fall through to plain download below
  }
  downloadBlob(blob, filename);
}

// Builds a minimal, valid single-page PDF that simply embeds a JPEG image.
// This avoids needing any PDF library or embedded font (which can't render
// Arabic without a font file) — the "text" is really the canvas rendering,
// so Arabic shows up exactly as the browser drew it.
function buildPdfFromJpegDataUrl(dataUrl, pixelWidth, pixelHeight, logicalWidth, logicalHeight) {
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

  const widthPt = Math.round(logicalWidth * 0.75);
  const heightPt = Math.round(logicalHeight * 0.75);

  const enc = new TextEncoder();
  const chunks = [];
  const offsets = {};
  let pos = 0;
  const push = (data) => {
    const b = typeof data === "string" ? enc.encode(data) : data;
    chunks.push(b);
    pos += b.length;
  };

  push("%PDF-1.4\n");
  offsets[1] = pos; push("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
  offsets[2] = pos; push("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");
  offsets[3] = pos; push(`3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${widthPt} ${heightPt}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>\nendobj\n`);
  offsets[4] = pos; push(`4 0 obj\n<< /Type /XObject /Subtype /Image /Width ${pixelWidth} /Height ${pixelHeight} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${bytes.length} >>\nstream\n`);
  push(bytes);
  push("\nendstream\nendobj\n");
  const content = `q ${widthPt} 0 0 ${heightPt} 0 0 cm /Im0 Do Q`;
  offsets[5] = pos; push(`5 0 obj\n<< /Length ${content.length} >>\nstream\n${content}\nendstream\nendobj\n`);
  const xrefStart = pos;
  push("xref\n0 6\n0000000000 65535 f \n");
  for (let i = 1; i <= 5; i++) push(`${String(offsets[i]).padStart(10, "0")} 00000 n \n`);
  push(`trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`);

  const total = chunks.reduce((s, c) => s + c.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return new Blob([out], { type: "application/pdf" });
}

function formatDateDisplay(dateKey) {
  try {
    const d = new Date(`${dateKey}T00:00:00`);
    const day = d.toLocaleDateString("ar-SA-u-nu-latn", { weekday: "long" });
    const date = d.toLocaleDateString("ar-SA-u-nu-latn", { day: "numeric", month: "long", year: "numeric" });
    return `${day}، ${date}`;
  } catch (e) {
    return dateKey;
  }
}
function formatDateDisplayEn(dateKey) {
  try {
    const d = new Date(`${dateKey}T00:00:00`);
    return d.toLocaleDateString("en-US", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  } catch (e) {
    return dateKey;
  }
}
function formatDateBilingual(dateKey) {
  return `${formatDateDisplay(dateKey)}  •  ${formatDateDisplayEn(dateKey)}`;
}

// ---- Hijri (تقويم هجري) <-> Gregorian conversion ----
// Standard arithmetic/tabular Islamic calendar via Julian Day Number. This is
// a well-known approximation (not the official Umm al-Qura sighting-based
// calendar), so it can differ by a day from the officially announced date —
// good enough for picking a date, without needing an external service.
const HIJRI_MONTHS = ["محرم", "صفر", "ربيع الأول", "ربيع الآخر", "جمادى الأولى", "جمادى الآخرة", "رجب", "شعبان", "رمضان", "شوال", "ذو القعدة", "ذو الحجة"];
const GREG_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
const WEEKDAY_LETTERS = ["أحد", "اثن", "ثلا", "أرب", "خمي", "جمع", "سبت"];

function gregorianToJD(year, month, day) {
  const a = Math.floor((14 - month) / 12);
  const y = year + 4800 - a;
  const m = month + 12 * a - 3;
  return day + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4) - Math.floor(y / 100) + Math.floor(y / 400) - 32045;
}
function jdToGregorian(jd) {
  const a = jd + 32044;
  const b = Math.floor((4 * a + 3) / 146097);
  const c = a - Math.floor((146097 * b) / 4);
  const d = Math.floor((4 * c + 3) / 1461);
  const e = c - Math.floor((1461 * d) / 4);
  const m = Math.floor((5 * e + 2) / 153);
  const day = e - Math.floor((153 * m + 2) / 5) + 1;
  const month = m + 3 - 12 * Math.floor(m / 10);
  const year = 100 * b + d - 4800 + Math.floor(m / 10);
  return { year, month, day };
}
function islamicToJD(year, month, day) {
  return day + Math.ceil(29.5 * (month - 1)) + (year - 1) * 354 + Math.floor((3 + 11 * year) / 30) + 1948439.5 - 1;
}
function jdToIslamic(jdInput) {
  const jd = Math.floor(jdInput) + 0.5;
  const year = Math.floor((30 * (jd - 1948439.5) + 10646) / 10631);
  let month = Math.min(12, Math.ceil((jd - (29 + islamicToJD(year, 1, 1))) / 29.5) + 1);
  month = Math.max(1, month);
  const day = Math.floor(jd - islamicToJD(year, month, 1)) + 1;
  return { year, month, day };
}
function isoToHijri(iso) {
  const d = new Date(`${iso}T00:00:00`);
  return jdToIslamic(gregorianToJD(d.getFullYear(), d.getMonth() + 1, d.getDate()));
}
function hijriToIso(y, m, d) {
  const jd = Math.round(islamicToJD(y, m, d));
  const g = jdToGregorian(jd);
  return `${String(g.year).padStart(4, "0")}-${String(g.month).padStart(2, "0")}-${String(g.day).padStart(2, "0")}`;
}
function hijriMonthLength(y, m) {
  const ny = m >= 12 ? y + 1 : y;
  const nm = m >= 12 ? 1 : m + 1;
  return Math.round(islamicToJD(ny, nm, 1) - islamicToJD(y, m, 1));
}
function weekdayOfIso(iso) {
  return new Date(`${iso}T00:00:00`).getDay();
}

// Shared: turns a print job (class board, student report, or attendance sheet)
// into a {title, subtitle, headers, rows, filename} table description used by
// the PNG/PDF/Excel exporters.
function jobToTable(job) {
  if (job.type === "class") {
    const cls = job.cls;
    const dateKey = job.dateKey;
    const t = dateKey || todayKey();
    const headers = [...(cls.showRowNumbers ? ["#"] : []), "الاسم", ...(cls.rows.length ? ["الغياب"] : []), ...cls.columns.map((c) => c.name)];
    const rows = cls.rows.map((row, i) => [
      ...(cls.showRowNumbers ? [i + 1] : []),
      row.name,
      attendanceStatus(cls, row.id, t) === "absent" ? "غائب" : "حاضر",
      ...cls.columns.map((col) => boardCellValue(cls, row.id, col.id, dateKey)),
    ]);
    return { title: cls.emoji ? `${cls.emoji} ${cls.subject}` : cls.subject, subtitle: dateKey ? `${cls.grade} • ${formatDateDisplay(dateKey)}` : `${cls.grade} • ${cls.teacher}`, headers, rows, filename: cls.subject || "الفصل" };
  }
  if (job.type === "blank") {
    const cls = job.cls;
    const headers = ["الاسم", ...cls.columns.map((c) => c.name)];
    const rows = cls.rows.map((row) => [row.name, ...cls.columns.map(() => "")]);
    return { title: cls.subject, subtitle: `${cls.grade} • ${cls.teacher} — نسخة فارغة`, headers, rows, filename: `${cls.subject || "الفصل"}-فارغ` };
  }
  if (job.type === "attendance") {
    const { cls, dateKey } = job;
    const headers = ["الاسم", "الحالة"];
    const rows = cls.rows.map((row) => [row.name, attendanceStatus(cls, row.id, dateKey) === "absent" ? "غائب" : "حاضر"]);
    return { title: `الحضور والغياب — ${formatDateDisplay(dateKey)}`, subtitle: `${cls.subject} • ${cls.grade}`, headers, rows, filename: `حضور-${cls.subject || "الفصل"}-${dateKey}` };
  }
  const { cls, row, entries } = job;
  const grouped = groupEntries(entries).flatMap((g) => g.items);
  const headers = ["العمود", "اليوم والتاريخ", "الوقت", "القيمة"];
  const rows = grouped.map((e) => [e.colName, `${e.day ? e.day + "، " : ""}${e.date || ""}`, e.time || "", e.value]);
  return { title: `تقرير الطالب: ${row.name}`, subtitle: `${cls.subject} • ${cls.grade}`, headers, rows, filename: `تقرير-${row.name}` };
}

function jobToCanvas(job) {
  if (job.type === "report") {
    const { cls, row, entries } = job;
    const groups = groupEntries(entries);
    return buildReportCanvas({ title: `تقرير الطالب: ${row.name}`, subtitle: `${cls.subject} • ${cls.grade} • ${cls.teacher}`, groups });
  }
  return buildTableCanvas(jobToTable(job));
}

function exportPng(job) {
  const filename = jobToTable(job).filename;
  const { canvas } = jobToCanvas(job);
  canvas.toBlob((blob) => downloadBlob(blob, `${filename}.png`));
}

// Builds a real PDF (image-wrapped, so Arabic renders correctly) and shares
// it directly via the device share sheet when available, otherwise downloads it.
async function exportPdfShare(job) {
  const filename = jobToTable(job).filename;
  const { canvas, logicalWidth, logicalHeight } = jobToCanvas(job);
  const dataUrl = canvas.toDataURL("image/jpeg", 0.97);
  const blob = buildPdfFromJpegDataUrl(dataUrl, canvas.width, canvas.height, logicalWidth, logicalHeight);
  await shareOrDownloadFile(blob, `${filename}.pdf`, "application/pdf");
}

function exportExcel(job) {
  const t = jobToTable(job);
  const ws = XLSX.utils.aoa_to_sheet([t.headers, ...t.rows]);
  // Give every column a generous width (especially the last one, usually free
  // text like notes/values) so Excel doesn't compress long Arabic text into a
  // narrow column and wrap it letter by letter.
  ws["!cols"] = t.headers.map((_, i) => ({ wch: i === t.headers.length - 1 ? 45 : 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "بيانات");
  XLSX.writeFile(wb, `${t.filename}.xlsx`);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

// Builds a self-contained, static HTML page (no app logic, nothing editable)
// showing the display board — meant to be shared with another teacher or
// admin who should only be able to view the data, not change it.
function buildReadOnlyBoardHtml(cls, dateKey) {
  const t = dateKey || todayKey();
  const headHtml = [
    ...(cls.showRowNumbers ? [`<th style="padding:8px;border:1px solid ${LINE};background:#F3F1E9;width:40px;">#</th>`] : []),
    `<th style="padding:8px;border:1px solid ${LINE};background:#F3F1E9;text-align:right;">الاسم</th>`,
    ...(cls.rows.length ? [`<th style="padding:8px;border:1px solid ${LINE};background:#F3F1E9;">الغياب</th>`] : []),
    ...cls.columns.map((col) => `<th style="padding:8px;border:1px solid ${LINE};background:${colorLight(col.color)};">${escapeHtml(col.name)}</th>`),
  ].join("");
  const rowsHtml = cls.rows.map((row, i) => {
    const status = attendanceStatus(cls, row.id, t);
    const cells = [
      ...(cls.showRowNumbers ? [`<td style="padding:8px;text-align:center;border:1px solid ${LINE};color:${MUTED};font-size:12px;">${i + 1}</td>`] : []),
      `<td style="padding:8px;font-weight:600;border:1px solid ${LINE};border-inline-start:4px solid ${row.color};">${escapeHtml(row.name)}</td>`,
      `<td style="padding:6px;text-align:center;border:1px solid ${LINE};font-weight:700;color:${status === "absent" ? "#C0392B" : "#0F6B5C"};">${status === "absent" ? "غائب" : "حاضر"}</td>`,
      ...cls.columns.map((col) => {
        const val = (dateKey ? valueOnDate(cls, row.id, col.id, dateKey) : (cls.cells[`${row.id}:${col.id}`] || lastReportedValue(cls, row.id, col.id))) || "—";
        return `<td style="padding:6px;text-align:center;border:1px solid ${LINE};">${escapeHtml(String(val))}</td>`;
      }),
    ].join("");
    return `<tr style="background:${i % 2 ? "#FBFAF6" : "#fff"}">${cells}</tr>`;
  }).join("");
  return `<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(cls.subject)} — لوحة العرض</title>
<style>
  body { font-family: Tahoma, Arial, sans-serif; background: ${PAPER}; color: ${INK}; padding: 24px; margin: 0; }
  table { border-collapse: collapse; width: 100%; font-size: 14px; }
  h1 { margin: 0 0 4px; font-size: 22px; }
  p.sub { color: ${MUTED}; margin: 0 0 16px; font-size: 14px; }
  .badge { display: inline-block; padding: 4px 12px; border-radius: 999px; background: #F3F1E9; font-size: 12px; color: ${MUTED}; margin-bottom: 16px; }
</style></head>
<body>
  <div class="badge">📄 نسخة للقراءة فقط — لا يمكن التعديل</div>
  <h1>${escapeHtml(cls.subject)}</h1>
  <p class="sub">${escapeHtml(cls.grade)} • ${escapeHtml(cls.teacher)} • ${escapeHtml(cls.yearHijri || "")} هـ / ${escapeHtml(cls.yearGregorian || "")} م</p>
  <table><thead><tr>${headHtml}</tr></thead><tbody>${rowsHtml}</tbody></table>
</body></html>`;
}

// Static, read-only HTML scoped to ONE student — used to share a single
// student's report with a parent without exposing the rest of the class.
function buildReadOnlyStudentHtml(cls, row, entries) {
  const groups = groupEntries(entries);
  const sectionsHtml = groups.map((g) => {
    const rowsHtml = g.items.map((e, i) => `
      <tr style="background:${i % 2 ? "#FBFAF6" : "#fff"}">
        <td style="padding:8px;border:1px solid ${LINE};">${escapeHtml(e.day ? `${e.day}، ` : "")}${escapeHtml(e.date || "")}</td>
        <td style="padding:8px;text-align:center;border:1px solid ${LINE};">${escapeHtml(e.time || "")}</td>
        <td style="padding:8px;border:1px solid ${LINE};">${escapeHtml(e.value)}</td>
      </tr>`).join("");
    return `
      <div style="border:1px solid ${LINE};border-radius:12px;overflow:hidden;margin-bottom:16px;">
        <div style="padding:10px 14px;background:${g.colColor};color:#fff;font-weight:700;display:flex;justify-content:space-between;">
          <span>${escapeHtml(g.colName)}</span><span>${g.items.length} رصد</span>
        </div>
        <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr>
            <th style="padding:8px;border:1px solid ${LINE};background:#F8F7F2;text-align:right;">اليوم والتاريخ</th>
            <th style="padding:8px;border:1px solid ${LINE};background:#F8F7F2;">الوقت</th>
            <th style="padding:8px;border:1px solid ${LINE};background:#F8F7F2;text-align:right;">القيمة</th>
          </tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>`;
  }).join("");
  return `<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>تقرير الطالب — ${escapeHtml(row.name)}</title>
<style>
  body { font-family: Tahoma, Arial, sans-serif; background: ${PAPER}; color: ${INK}; padding: 24px; margin: 0; max-width: 680px; margin: 0 auto; }
  h1 { margin: 0 0 4px; font-size: 22px; }
  p.sub { color: ${MUTED}; margin: 0 0 16px; font-size: 14px; }
  .badge { display: inline-block; padding: 4px 12px; border-radius: 999px; background: #F3F1E9; font-size: 12px; color: ${MUTED}; margin-bottom: 16px; }
</style></head>
<body>
  <div class="badge">📄 نسخة للقراءة فقط — خاصة بهذا الطالب فقط</div>
  <h1>${escapeHtml(row.name)}</h1>
  <p class="sub">${escapeHtml(cls.subject)} • ${escapeHtml(cls.grade)} • ${escapeHtml(cls.teacher)}</p>
  ${groups.length === 0 ? `<p style="color:${MUTED};">لا يوجد رصد بعد.</p>` : sectionsHtml}
</body></html>`;
}

function useFonts() {
  useEffect(() => {
    if (document.getElementById("mutabaa-fonts")) return;
    const link = document.createElement("link");
    link.id = "mutabaa-fonts";
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&family=IBM+Plex+Sans+Arabic:wght@400;500;600&display=swap";
    document.head.appendChild(link);
  }, []);
}

// ---------- print system (robust, doesn't rely on Tailwind print: variant) ----------
function PrintStyles() {
  return (
    <style>{`
      .app-print-root { display: none; }
      @media print {
        body * { visibility: hidden !important; }
        .app-print-root, .app-print-root * { visibility: visible !important; }
        .app-print-root { display: block !important; position: absolute; inset: 0; padding: 24px; }
      }
      @keyframes tossToTrash {
        0% { opacity: 1; transform: translateY(0) rotate(0deg) scale(1); }
        55% { opacity: 0.75; transform: translateY(50px) rotate(18deg) scale(0.82); }
        100% { opacity: 0; transform: translateY(110px) rotate(45deg) scale(0.2); }
      }
      .trash-toss { animation: tossToTrash 0.45s ease-in forwards; pointer-events: none; }
      @keyframes tickerScrollSingle {
        from { transform: translateX(-100%); }
        to { transform: translateX(100%); }
      }
      .ticker-track-single { position: relative; width: 100%; animation-name: tickerScrollSingle; animation-timing-function: linear; animation-iteration-count: infinite; }
      .ticker-track-single span { white-space: nowrap; display: inline-block; }
      @keyframes spin { to { transform: rotate(360deg); } }
      @keyframes modalBackdropIn { from { opacity: 0; } to { opacity: 1; } }
      @keyframes modalPanelIn {
        from { opacity: 0; transform: translateY(14px) scale(0.97); }
        to { opacity: 1; transform: translateY(0) scale(1); }
      }
      .modal-backdrop-in { animation: modalBackdropIn 0.18s ease-out forwards; }
      .modal-panel-in { animation: modalPanelIn 0.22s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
      @keyframes pageFadeIn {
        from { opacity: 0; transform: translateY(6px); }
        to { opacity: 1; transform: translateY(0); }
      }
      .page-fade-in { animation: pageFadeIn 0.25s ease-out forwards; }
      @keyframes cardIn {
        from { opacity: 0; transform: translateY(8px) scale(0.98); }
        to { opacity: 1; transform: translateY(0) scale(1); }
      }
      .card-in { animation: cardIn 0.3s cubic-bezier(0.16, 1, 0.3, 1) both; }
    `}</style>
  );
}

function usePrint() {
  const [job, setJob] = useState(null);
  useEffect(() => {
    if (!job) return;
    const t = setTimeout(() => { window.print(); }, 80);
    return () => clearTimeout(t);
  }, [job]);
  return [job, setJob];
}

// ---------- generic small components ----------

// One consistent look for every compact icon-only action (table column/row
// controls, small list actions, etc.) — fixed size, padding, and hover state
// so these don't feel scattered across different parts of the app.
function MiniIconBtn({ icon: Icon, onClick, title, color, disabled }) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={title}
      className="p-1 rounded-md hover:bg-black/5 active:scale-90 disabled:opacity-25 disabled:active:scale-100 transition-transform shrink-0 flex items-center justify-center"
      style={{ color: color || MUTED }}
    >
      <Icon size={13} strokeWidth={2} />
    </button>
  );
}

function IconBtn({ icon: Icon, label, onClick, tone = "default" }) {
  const tones = {
    default: { bg: "transparent", fg: INK, border: LINE },
    danger: { bg: "#FBEDEA", fg: "#9A3B2E", border: "#F0D2CB" },
    primary: { bg: "#0F6B5C", fg: "#fff", border: "#0F6B5C" },
  };
  const t = tones[tone];
  return (
    <button
      onClick={onClick}
      className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-sm font-medium transition-all hover:opacity-80 active:scale-95 whitespace-nowrap"
      style={{ background: t.bg, color: t.fg, border: `1px solid ${t.border}` }}
    >
      <Icon size={16} strokeWidth={2} />
      <span>{label}</span>
    </button>
  );
}

// A single calendar popover that lets the person pick the date using either
// the Hijri or Gregorian calendar (toggle inside the popover) — both resolve
// to the same underlying ISO date used everywhere else in the app.
function DateField({ value, onChange }) {
  const [open, setOpen] = useState(false);
  const [mode, setMode] = useState("gregorian");
  const [viewY, setViewY] = useState(2024);
  const [viewM, setViewM] = useState(1);
  const [hViewY, setHViewY] = useState(1445);
  const [hViewM, setHViewM] = useState(1);

  const openPicker = () => {
    const b = value ? new Date(`${value}T00:00:00`) : new Date();
    setViewY(b.getFullYear());
    setViewM(b.getMonth() + 1);
    const h = isoToHijri(value || todayKey());
    setHViewY(h.year);
    setHViewM(h.month);
    setOpen((o) => !o);
  };

  const selectIso = (iso) => { onChange(iso); setOpen(false); };

  const gDaysInMonth = new Date(viewY, viewM, 0).getDate();
  const gFirstWeekday = new Date(viewY, viewM - 1, 1).getDay();
  const gCells = [...Array(gFirstWeekday).fill(null), ...Array.from({ length: gDaysInMonth }, (_, i) => i + 1)];
  const gPrev = () => { let m = viewM - 1, y = viewY; if (m < 1) { m = 12; y -= 1; } setViewM(m); setViewY(y); };
  const gNext = () => { let m = viewM + 1, y = viewY; if (m > 12) { m = 1; y += 1; } setViewM(m); setViewY(y); };
  const gSelect = (d) => selectIso(`${String(viewY).padStart(4, "0")}-${String(viewM).padStart(2, "0")}-${String(d).padStart(2, "0")}`);

  const hDaysInMonth = hijriMonthLength(hViewY, hViewM);
  const hFirstWeekday = weekdayOfIso(hijriToIso(hViewY, hViewM, 1));
  const hCells = [...Array(hFirstWeekday).fill(null), ...Array.from({ length: hDaysInMonth }, (_, i) => i + 1)];
  const hPrev = () => { let m = hViewM - 1, y = hViewY; if (m < 1) { m = 12; y -= 1; } setHViewM(m); setHViewY(y); };
  const hNext = () => { let m = hViewM + 1, y = hViewY; if (m > 12) { m = 1; y += 1; } setHViewM(m); setHViewY(y); };
  const hSelect = (d) => selectIso(hijriToIso(hViewY, hViewM, d));

  return (
    <div className="relative inline-block">
      <button
        type="button"
        onClick={openPicker}
        className="flex items-center gap-1.5 text-sm font-medium px-2.5 py-1.5 rounded-lg"
        style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}
      >
        <Calendar size={15} color={MUTED} />
        {value ? formatDateDisplay(value) : "اختر تاريخًا"}
      </button>
      {open && (
        <div className="absolute z-30 top-full mt-1 rounded-xl shadow-lg p-3" style={{ background: "#fff", border: `1px solid ${LINE}`, width: 280 }}>
          <div className="flex gap-2 mb-3">
            <button onClick={() => setMode("gregorian")} className="flex-1 text-xs font-bold py-1.5 rounded-lg" style={{ background: mode === "gregorian" ? INK : "transparent", color: mode === "gregorian" ? "#fff" : MUTED, border: `1px solid ${mode === "gregorian" ? INK : LINE}` }}>ميلادي</button>
            <button onClick={() => setMode("hijri")} className="flex-1 text-xs font-bold py-1.5 rounded-lg" style={{ background: mode === "hijri" ? INK : "transparent", color: mode === "hijri" ? "#fff" : MUTED, border: `1px solid ${mode === "hijri" ? INK : LINE}` }}>هجري</button>
          </div>
          {mode === "gregorian" ? (
            <>
              <div className="flex items-center justify-between mb-2">
                <button onClick={gPrev} className="p-1 rounded hover:bg-black/5"><ChevronRight size={16} /></button>
                <span className="text-sm font-bold" style={{ color: INK }}>{GREG_MONTHS[viewM - 1]} {viewY}</span>
                <button onClick={gNext} className="p-1 rounded hover:bg-black/5"><ChevronLeft size={16} /></button>
              </div>
              <div className="grid grid-cols-7 gap-1 mb-1">
                {WEEKDAY_LETTERS.map((w) => <div key={w} className="text-center" style={{ fontSize: 10, fontWeight: 700, color: MUTED }}>{w}</div>)}
              </div>
              <div className="grid grid-cols-7 gap-1">
                {gCells.map((d, i) => {
                  const iso = d ? `${String(viewY).padStart(4, "0")}-${String(viewM).padStart(2, "0")}-${String(d).padStart(2, "0")}` : null;
                  const isSel = iso && iso === value;
                  return d ? (
                    <button key={i} onClick={() => gSelect(d)} className="text-xs py-1.5 rounded-md hover:bg-black/5" style={{ background: isSel ? "#0F6B5C" : "transparent", color: isSel ? "#fff" : INK, fontWeight: isSel ? 700 : 400 }}>{d}</button>
                  ) : <div key={i} />;
                })}
              </div>
            </>
          ) : (
            <>
              <div className="flex items-center justify-between mb-2">
                <button onClick={hPrev} className="p-1 rounded hover:bg-black/5"><ChevronRight size={16} /></button>
                <span className="text-sm font-bold" style={{ color: INK }}>{HIJRI_MONTHS[hViewM - 1]} {hViewY}</span>
                <button onClick={hNext} className="p-1 rounded hover:bg-black/5"><ChevronLeft size={16} /></button>
              </div>
              <div className="grid grid-cols-7 gap-1 mb-1">
                {WEEKDAY_LETTERS.map((w) => <div key={w} className="text-center" style={{ fontSize: 10, fontWeight: 700, color: MUTED }}>{w}</div>)}
              </div>
              <div className="grid grid-cols-7 gap-1">
                {hCells.map((d, i) => {
                  const iso = d ? hijriToIso(hViewY, hViewM, d) : null;
                  const isSel = iso && iso === value;
                  return d ? (
                    <button key={i} onClick={() => hSelect(d)} className="text-xs py-1.5 rounded-md hover:bg-black/5" style={{ background: isSel ? "#0F6B5C" : "transparent", color: isSel ? "#fff" : INK, fontWeight: isSel ? 700 : 400 }}>{d}</button>
                  ) : <div key={i} />;
                })}
              </div>
            </>
          )}
          <button onClick={() => selectIso(todayKey())} className="w-full mt-3 text-xs font-bold py-1.5 rounded-lg" style={{ border: `1px solid ${LINE}`, color: "#0F6B5C" }}>اليوم</button>
        </div>
      )}
    </div>
  );
}

function ConfirmDialog({ title, message, onConfirm, onCancel, confirmLabel = "حذف" }) {
  return (
    <div
      className="fixed inset-0 z-[70] flex items-center justify-center p-4 modal-backdrop-in"
      style={{ background: "rgba(35,38,34,0.5)" }}
      onMouseDown={(e) => { if (e.target === e.currentTarget) onCancel(); }}
    >
      <div className="rounded-2xl shadow-2xl w-full max-w-sm p-5 modal-panel-in" style={{ background: PAPER, border: `1px solid ${LINE}` }}>
        <h3 className="font-bold text-base mb-2" style={{ color: INK }}>{title}</h3>
        <p className="text-sm mb-5" style={{ color: MUTED }}>{message}</p>
        <div className="flex justify-end gap-2">
          <button onClick={onCancel} className="px-4 py-2 rounded-lg text-sm font-medium active:scale-95 transition-transform" style={{ color: MUTED }}>إلغاء</button>
          <button onClick={onConfirm} className="px-4 py-2 rounded-lg text-sm font-bold text-white active:scale-95 transition-transform" style={{ background: "#C0392B" }}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

function PrintFormatModal({ onClose, onChoose }) {
  const options = [
    { key: "pdf", label: "PDF", icon: FileText, desc: "مشاركة/حفظ مباشر" },
    { key: "png", label: "PNG", icon: FileImage, desc: "صورة للمشاركة" },
    { key: "excel", label: "Excel", icon: FileSpreadsheet, desc: "ملف بيانات" },
  ];
  return (
    <Modal title="اختر صيغة التصدير" onClose={onClose} zIndex={80}>
      <div className="grid grid-cols-3 gap-3">
        {options.map((o) => (
          <button
            key={o.key}
            onClick={() => onChoose(o.key)}
            className="flex flex-col items-center gap-2 p-4 rounded-xl hover:opacity-80 transition-opacity"
            style={{ border: `1px solid ${LINE}`, background: "#fff" }}
          >
            <o.icon size={26} color="#0F6B5C" />
            <span className="font-bold text-sm" style={{ color: INK }}>{o.label}</span>
            <span className="text-xs text-center" style={{ color: MUTED }}>{o.desc}</span>
          </button>
        ))}
      </div>
    </Modal>
  );
}

function Modal({ title, onClose, children, wide = false, lg = false, xl = false, zIndex = 50 }) {
  const widthClass = xl ? "max-w-6xl" : lg ? "max-w-5xl" : wide ? "max-w-3xl" : "max-w-md";
  return (
    <div
      className="fixed inset-0 flex items-center justify-center p-4 modal-backdrop-in"
      style={{ background: "rgba(35,38,34,0.45)", zIndex }}
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        className={`rounded-2xl shadow-2xl w-full ${widthClass} max-h-[90vh] overflow-y-auto modal-panel-in`}
        style={{ background: PAPER, border: `1px solid ${LINE}` }}
      >
        <div className="flex items-center justify-between px-5 py-4 sticky top-0" style={{ background: PAPER, borderBottom: `1px solid ${LINE}` }}>
          <h3 className="font-bold text-lg" style={{ color: INK }}>{title}</h3>
          <button onClick={onClose} className="p-1.5 rounded-full hover:bg-black/5 active:scale-90 transition-transform">
            <X size={18} color={MUTED} />
          </button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children, hint }) {
  return (
    <div className="mb-4">
      <label className="block text-sm font-semibold mb-1.5" style={{ color: INK }}>{label}</label>
      {children}
      {hint && <p className="text-xs mt-1" style={{ color: MUTED }}>{hint}</p>}
    </div>
  );
}

const inputStyle = {
  width: "100%",
  padding: "9px 12px",
  borderRadius: "10px",
  border: `1px solid ${LINE}`,
  background: "#fff",
  fontFamily: "'IBM Plex Sans Arabic', sans-serif",
  fontSize: "14px",
  color: INK,
  outline: "none",
};

function ColorSwatches({ value, onChange, size = 8 }) {
  return (
    <div className="flex flex-wrap gap-1.5">
      {COLORS.map((c) => (
        <button
          key={c.hex}
          type="button"
          title={c.name}
          onClick={() => onChange(c.hex)}
          className="rounded-full flex items-center justify-center transition-transform hover:scale-110 shrink-0"
          style={{ width: size * 4, height: size * 4, background: c.hex, boxShadow: value === c.hex ? `0 0 0 2px #fff, 0 0 0 4px ${c.hex}` : "none" }}
        >
          {value === c.hex && <Check size={size * 1.5} color="#fff" strokeWidth={3} />}
        </button>
      ))}
    </div>
  );
}

// ---------- option editor (per-option color) ----------

function OptionsEditor({ options, onChange }) {
  const update = (id, patch) => onChange(options.map((o) => (o.id === id ? { ...o, ...patch } : o)));
  const remove = (id) => onChange(options.filter((o) => o.id !== id));
  const add = () => onChange([...options, { id: uid(), label: "", color: COLORS[0].hex }]);
  return (
    <div>
      <div className="space-y-2">
        {options.map((o) => (
          <div key={o.id} className="flex items-center gap-2 p-2 rounded-lg" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <span className="w-3.5 h-3.5 rounded-full shrink-0" style={{ background: o.color }} />
            <input
              value={o.label}
              onChange={(e) => update(o.id, { label: e.target.value })}
              placeholder="نص الخيار"
              className="flex-1 text-sm outline-none bg-transparent"
              style={{ color: INK }}
            />
            <div className="flex gap-1 shrink-0">
              {COLORS.slice(0, 6).map((c) => (
                <button key={c.hex} type="button" onClick={() => update(o.id, { color: c.hex })}
                  className="w-4 h-4 rounded-full" style={{ background: c.hex, boxShadow: o.color === c.hex ? `0 0 0 2px #fff, 0 0 0 3px ${c.hex}` : "none" }} />
              ))}
            </div>
            <button type="button" onClick={() => remove(o.id)} className="p-1 rounded hover:bg-black/5 shrink-0">
              <X size={13} color={MUTED} />
            </button>
          </div>
        ))}
      </div>
      <button type="button" onClick={add} className="mt-2 text-xs font-semibold flex items-center gap-1" style={{ color: "#0F6B5C" }}>
        <Plus size={13} /> إضافة خيار
      </button>
    </div>
  );
}

// ---------- reusable field group for a column/row draft ----------

function TypePicker({ type, setType, labels }) {
  return (
    <div className="flex gap-2">
      {[
        { v: "text", label: labels?.text || "نص حر", icon: Type },
        { v: "counter", label: "عداد (+/-)", icon: Hash },
        { v: "dropdown", label: "قائمة منسدلة", icon: ListChecks },
      ].map((opt) => (
        <button
          key={opt.v}
          type="button"
          onClick={() => setType(opt.v)}
          className="flex-1 flex flex-col items-center gap-1 py-2.5 rounded-xl text-xs font-medium"
          style={{
            border: `1px solid ${type === opt.v ? "#0F6B5C" : LINE}`,
            background: type === opt.v ? "#EAF3F0" : "#fff",
            color: type === opt.v ? "#0F6B5C" : MUTED,
          }}
        >
          <opt.icon size={16} />
          {opt.label}
        </button>
      ))}
    </div>
  );
}

function ColumnDraftForm({ draft, onChange, onRemove, removable }) {
  const set = (patch) => onChange({ ...draft, ...patch });
  return (
    <div className="rounded-xl p-3 mb-3" style={{ border: `1px solid ${LINE}`, background: "#FCFBF7" }}>
      {removable && (
        <div className="flex justify-end mb-1">
          <button onClick={onRemove} className="p-1 rounded hover:bg-black/5"><X size={14} color={MUTED} /></button>
        </div>
      )}
      <Field label="اسم العمود">
        <input style={inputStyle} value={draft.name} onChange={(e) => set({ name: e.target.value })} placeholder="مثال: الواجب اليومي" />
      </Field>
      <Field label="طريقة عمل العمود">
        <TypePicker type={draft.type} setType={(v) => set({ type: v })} />
      </Field>
      {draft.type === "dropdown" && (
        <Field label="خيارات القائمة (لكل خيار لونه الخاص)">
          <OptionsEditor options={draft.options} onChange={(options) => set({ options })} />
        </Field>
      )}
      <Field label="لون العمود">
        <ColorSwatches value={draft.color} onChange={(color) => set({ color })} />
      </Field>
      <label className="flex items-center gap-2 text-sm font-medium mt-1" style={{ color: INK }}>
        <input type="checkbox" checked={!!draft.pinned} onChange={(e) => set({ pinned: e.target.checked })} />
        تثبيت العمود (يبقى ظاهرًا أثناء التمرير الأفقي)
      </label>
      <label className="flex items-center gap-2 text-sm font-medium mt-1" style={{ color: INK }}>
        <input type="checkbox" checked={!!draft.autoRenew} onChange={(e) => set({ autoRenew: e.target.checked })} />
        تفريغ تلقائي لهذا العمود
      </label>
      <p className="text-xs mt-1" style={{ color: MUTED }}>
        يُفرَّغ الرصد تلقائيًا فقط عندما يكون "التجديد التلقائي" مفعّلاً في الصف <b>و</b> في العمود معًا — هذا يتيح لك استثناء عمود معيّن (مثل الدرجة النهائية) من الاختفاء حتى لو كان الصف مفعّلاً فيه التجديد.
      </p>
      <Field label="رصد نفس القيمة لجميع الطلاب فور الإنشاء (اختياري)">
        {draft.type === "dropdown" ? (
          <select value={draft.bulkValue || ""} onChange={(e) => set({ bulkValue: e.target.value })} style={inputStyle}>
            <option value="">بدون رصد جماعي</option>
            {(draft.options || []).filter((o) => o.label.trim()).map((o) => <option key={o.id} value={o.label}>{o.label}</option>)}
          </select>
        ) : (
          <input
            type={draft.type === "counter" ? "number" : "text"}
            value={draft.bulkValue || ""}
            onChange={(e) => set({ bulkValue: e.target.value })}
            style={inputStyle}
            placeholder={draft.type === "counter" ? "مثال: 5" : "اتركه فارغًا لعدم الرصد الجماعي"}
          />
        )}
      </Field>
      <label className="flex items-center gap-2 text-sm font-medium mt-1" style={{ color: INK }}>
        <input type="checkbox" checked={!!draft.behaviorFlag} onChange={(e) => set({ behaviorFlag: e.target.checked })} />
        تتبّع هذا العمود كملاحظات سلوكية سلبية
      </label>
      {draft.behaviorFlag && (
        <Field label="تنبيه تلقائي عند الوصول إلى عدد الملاحظات">
          <input
            type="number"
            min={1}
            value={draft.behaviorThreshold ?? 3}
            onChange={(e) => set({ behaviorThreshold: Math.max(1, Number(e.target.value) || 1) })}
            style={{ ...inputStyle, width: 90 }}
          />
        </Field>
      )}
    </div>
  );
}

function emptyColumnDraft() {
  return { key: uid(), name: "", type: "text", options: [], color: COLORS[2].hex, autoRenew: false, pinned: false, bulkValue: "", behaviorFlag: false, behaviorThreshold: 3 };
}

function ColumnModal({ initial, onClose, onSaveMany, onSaveOne, onDelete }) {
  const isEdit = !!initial;
  const [single, setSingle] = useState(() => (initial ? { ...initial } : null));
  const [drafts, setDrafts] = useState(() => (isEdit ? [] : [emptyColumnDraft()]));
  const [tab, setTab] = useState("detailed");
  const [bulkText, setBulkText] = useState("");
  const [bulkType, setBulkType] = useState("text");
  const [bulkColor, setBulkColor] = useState(COLORS[2].hex);

  const validSingle = single && single.name.trim();
  const validDrafts = drafts.filter((d) => d.name.trim());
  const bulkNames = bulkText.split("\n").map((s) => s.trim()).filter(Boolean);

  return (
    <Modal title={isEdit ? "تعديل العمود" : "إضافة عمود"} onClose={onClose} wide={!isEdit}>
      {isEdit ? (
        <ColumnDraftForm draft={single} onChange={setSingle} removable={false} />
      ) : (
        <>
          <div className="flex gap-2 mb-4">
            <button onClick={() => setTab("detailed")} className="px-3 py-1.5 rounded-full text-xs font-semibold"
              style={{ background: tab === "detailed" ? INK : "transparent", color: tab === "detailed" ? "#fff" : MUTED, border: `1px solid ${tab === "detailed" ? INK : LINE}` }}>
              إضافة تفصيلية
            </button>
            <button onClick={() => setTab("bulk")} className="px-3 py-1.5 rounded-full text-xs font-semibold"
              style={{ background: tab === "bulk" ? INK : "transparent", color: tab === "bulk" ? "#fff" : MUTED, border: `1px solid ${tab === "bulk" ? INK : LINE}` }}>
              إضافة دفعة واحدة
            </button>
          </div>

          {tab === "detailed" ? (
            <>
              {drafts.map((d, i) => (
                <ColumnDraftForm
                  key={d.key}
                  draft={d}
                  removable={drafts.length > 1}
                  onRemove={() => setDrafts(drafts.filter((x) => x.key !== d.key))}
                  onChange={(nd) => setDrafts(drafts.map((x, xi) => (xi === i ? nd : x)))}
                />
              ))}
              <button
                onClick={() => setDrafts([...drafts, emptyColumnDraft()])}
                className="mb-2 text-sm font-semibold flex items-center gap-1"
                style={{ color: "#0F6B5C" }}
              >
                <Plus size={15} /> إضافة عمود آخر
              </button>
            </>
          ) : (
            <>
              <Field label="أسماء الأعمدة (كل اسم في سطر)" hint="مفيد للأعمدة المتشابهة، مثل: اختبار 1، اختبار 2، اختبار 3...">
                <textarea
                  style={{ ...inputStyle, minHeight: "140px", resize: "vertical" }}
                  value={bulkText}
                  onChange={(e) => setBulkText(e.target.value)}
                  placeholder={"اختبار 1\nاختبار 2\nاختبار 3"}
                />
              </Field>
              <Field label="طريقة عمل مشتركة لهذه الأعمدة">
                <TypePicker type={bulkType} setType={setBulkType} />
              </Field>
              <Field label="لون مشترك لهذه الأعمدة">
                <ColorSwatches value={bulkColor} onChange={setBulkColor} />
              </Field>
              <p className="text-xs mt-1" style={{ color: MUTED }}>سيُضاف {bulkNames.length} عمود.</p>
            </>
          )}
        </>
      )}
      <div className="flex justify-between items-center mt-4">
        <div>
          {isEdit && (
            <button onClick={onDelete} className="px-3 py-2 rounded-lg text-sm font-medium" style={{ color: "#9A3B2E" }}>حذف العمود</button>
          )}
        </div>
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
          <button
            disabled={isEdit ? !validSingle : (tab === "detailed" ? validDrafts.length === 0 : bulkNames.length === 0)}
            onClick={() => {
              if (isEdit) return onSaveOne(single);
              if (tab === "detailed") return onSaveMany(validDrafts);
              return onSaveMany(bulkNames.map((name) => ({ name, type: bulkType, options: [], color: bulkColor, autoRenew: false, pinned: false, bulkValue: "" })));
            }}
            className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
            style={{ background: "#0F6B5C" }}
          >تم</button>
        </div>
      </div>
    </Modal>
  );
}

// ---------- Row modal (individual multi-add + bulk quick add) ----------

function emptyRowDraft() {
  return { key: uid(), name: "", type: "text", options: [], color: COLORS[4].hex, autoRenew: false, medicalNote: "" };
}

function RowDraftForm({ draft, onChange, onRemove, removable }) {
  const set = (patch) => onChange({ ...draft, ...patch });
  return (
    <div className="rounded-xl p-3 mb-3" style={{ border: `1px solid ${LINE}`, background: "#FCFBF7" }}>
      {removable && (
        <div className="flex justify-end mb-1">
          <button onClick={onRemove} className="p-1 rounded hover:bg-black/5"><X size={14} color={MUTED} /></button>
        </div>
      )}
      <Field label="اسم الصف">
        <input style={inputStyle} value={draft.name} onChange={(e) => set({ name: e.target.value })} placeholder="مثال: اسم الطالب" />
      </Field>
      <Field label="طريقة عمل الصف" hint="نوع خلايا الجدول الفعلي يُحدَّد حسب نوع كل عمود؛ هذا التصنيف تسمية توضيحية للصف.">
        <TypePicker type={draft.type} setType={(v) => set({ type: v })} labels={{ text: "مربع فارغ" }} />
      </Field>
      {draft.type === "dropdown" && (
        <Field label="خيارات القائمة">
          <OptionsEditor options={draft.options} onChange={(options) => set({ options })} />
        </Field>
      )}
      <Field label="لون الصف">
        <ColorSwatches value={draft.color} onChange={(color) => set({ color })} />
      </Field>
      <label className="flex items-center gap-2 text-sm font-medium mt-1" style={{ color: INK }}>
        <input type="checkbox" checked={draft.autoRenew} onChange={(e) => set({ autoRenew: e.target.checked })} />
        تجديد تلقائي لهذا الصف
      </label>
      <p className="text-xs mt-1" style={{ color: MUTED }}>
        يعمل فقط مع الأعمدة التي فُعِّل فيها "تفريغ تلقائي" أيضًا؛ أي عمود لا يفعّلها يحتفظ برصده لهذا الصف حتى لو كان الصف نفسه مفعّلًا فيه التجديد.
      </p>
      <Field label="ملاحظة طبية / تنبيه خاص (اختياري)" hint="يظهر كأيقونة تنبيه واضحة بجانب اسم الطالب في كل مكان — للحالات التي تحتاج انتباهًا سريعًا.">
        <textarea
          style={{ ...inputStyle, minHeight: "60px", resize: "vertical" }}
          value={draft.medicalNote || ""}
          onChange={(e) => set({ medicalNote: e.target.value })}
          placeholder="مثال: حساسية من الفول السوداني، يحتاج استخدام النظارة..."
        />
      </Field>
    </div>
  );
}

function RowModal({ initial, onClose, onSaveMany, onSaveOne, onDelete, showRowNumbers, onToggleShowRowNumbers }) {
  const isEdit = !!initial;
  const [single, setSingle] = useState(() => (initial ? { ...initial } : null));
  const [drafts, setDrafts] = useState(() => (isEdit ? [] : [emptyRowDraft()]));
  const [tab, setTab] = useState("detailed");
  const [bulkText, setBulkText] = useState("");
  const [bulkColor, setBulkColor] = useState(COLORS[4].hex);
  const [bulkAutoRenew, setBulkAutoRenew] = useState(false);
  const [bulkMedicalNote, setBulkMedicalNote] = useState("");
  const [importNames, setImportNames] = useState([]);
  const [importFileName, setImportFileName] = useState("");
  const [importError, setImportError] = useState("");
  const importInputRef = useRef(null);

  const validSingle = single && single.name.trim();
  const validDrafts = drafts.filter((d) => d.name.trim());
  const bulkNames = bulkText.split("\n").map((s) => s.trim()).filter(Boolean);

  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportFileName(file.name);
    setImportError("");
    setImportNames([]);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const HEADER_WORDS = ["اسم", "الاسم", "اسم الطالب", "الطالب", "name", "student", "students"];
        const names = [];
        rows.forEach((r) => {
          const val = String(r[0] ?? "").trim();
          if (!val) return;
          if (HEADER_WORDS.includes(val.toLowerCase())) return;
          names.push(val);
        });
        if (names.length === 0) setImportError("لم يتم العثور على أي أسماء في العمود الأول من الملف.");
        setImportNames(names);
      } catch (err) {
        setImportError("تعذّرت قراءة هذا الملف. تأكد أنه ملف Excel (.xlsx) أو CSV صالح.");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <Modal title={isEdit ? "تعديل الصف" : "إضافة صف"} onClose={onClose} wide={!isEdit}>
      <label className="flex items-center gap-2 text-sm font-medium mb-4 p-3 rounded-xl" style={{ color: INK, background: "#F3F1E9", border: `1px solid ${LINE}` }}>
        <input type="checkbox" checked={!!showRowNumbers} onChange={onToggleShowRowNumbers} />
        <ListOrdered size={15} color={MUTED} />
        إظهار ترقيم تسلسلي للصفوف (١، ٢، ٣...) — يتحدّث تلقائيًا عند السحب وإعادة الترتيب
      </label>
      {isEdit ? (
        <RowDraftForm draft={single} onChange={setSingle} removable={false} />
      ) : (
        <>
          <div className="flex gap-2 mb-4 flex-wrap">
            <button onClick={() => setTab("detailed")} className="px-3 py-1.5 rounded-full text-xs font-semibold"
              style={{ background: tab === "detailed" ? INK : "transparent", color: tab === "detailed" ? "#fff" : MUTED, border: `1px solid ${tab === "detailed" ? INK : LINE}` }}>
              إضافة تفصيلية
            </button>
            <button onClick={() => setTab("bulk")} className="px-3 py-1.5 rounded-full text-xs font-semibold"
              style={{ background: tab === "bulk" ? INK : "transparent", color: tab === "bulk" ? "#fff" : MUTED, border: `1px solid ${tab === "bulk" ? INK : LINE}` }}>
              إضافة دفعة واحدة
            </button>
            <button onClick={() => setTab("import")} className="px-3 py-1.5 rounded-full text-xs font-semibold flex items-center gap-1"
              style={{ background: tab === "import" ? INK : "transparent", color: tab === "import" ? "#fff" : MUTED, border: `1px solid ${tab === "import" ? INK : LINE}` }}>
              <FileSpreadsheet size={13} /> استيراد من Excel
            </button>
          </div>

          {tab === "detailed" ? (
            <>
              {drafts.map((d, i) => (
                <RowDraftForm
                  key={d.key}
                  draft={d}
                  removable={drafts.length > 1}
                  onRemove={() => setDrafts(drafts.filter((x) => x.key !== d.key))}
                  onChange={(nd) => setDrafts(drafts.map((x, xi) => (xi === i ? nd : x)))}
                />
              ))}
              <button onClick={() => setDrafts([...drafts, emptyRowDraft()])} className="mb-2 text-sm font-semibold flex items-center gap-1" style={{ color: "#0F6B5C" }}>
                <Plus size={15} /> إضافة صف آخر
              </button>
            </>
          ) : tab === "bulk" ? (
            <>
              <Field label="أسماء الصفوف (كل اسم في سطر)">
                <textarea
                  style={{ ...inputStyle, minHeight: "140px", resize: "vertical" }}
                  value={bulkText}
                  onChange={(e) => setBulkText(e.target.value)}
                  placeholder={"محمد أحمد\nسارة خالد\nعبدالله فهد"}
                />
              </Field>
              <Field label="لون مشترك لهذه الصفوف">
                <ColorSwatches value={bulkColor} onChange={setBulkColor} />
              </Field>
              <label className="flex items-center gap-2 text-sm font-medium" style={{ color: INK }}>
                <input type="checkbox" checked={bulkAutoRenew} onChange={(e) => setBulkAutoRenew(e.target.checked)} />
                تفعيل التجديد التلقائي لهذه الصفوف
              </label>
              <Field label="ملاحظة طبية / تنبيه خاص مشتركة (اختياري)" hint="تُطبَّق على كل الصفوف المضافة هنا — إن كانت متشابهة. يمكنك تعديلها لكل طالب لاحقًا من نافذة تعديله.">
                <textarea
                  style={{ ...inputStyle, minHeight: "60px", resize: "vertical" }}
                  value={bulkMedicalNote}
                  onChange={(e) => setBulkMedicalNote(e.target.value)}
                  placeholder="اتركه فارغًا إذا لا يوجد تنبيه مشترك"
                />
              </Field>
              <p className="text-xs mt-2" style={{ color: MUTED }}>سيُضاف {bulkNames.length} صف بنوع "مربع فارغ".</p>
            </>
          ) : (
            <>
              <Field label="ملف Excel أو CSV بأسماء الطلاب (أول عمود)">
                <input ref={importInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} style={{ display: "none" }} />
                <button type="button" onClick={() => importInputRef.current?.click()} className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold" style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}>
                  <FileSpreadsheet size={15} color="#0F6B5C" /> اختر ملفًا
                </button>
                {importFileName && <p className="text-xs mt-1.5" style={{ color: MUTED }}>{importFileName}</p>}
              </Field>
              {importError && <p className="text-xs mb-3 font-medium" style={{ color: "#C0392B" }}>{importError}</p>}
              {importNames.length > 0 && (
                <div className="mb-3 p-2 rounded-lg" style={{ border: `1px solid ${LINE}`, maxHeight: 150, overflowY: "auto", background: "#F8F7F2" }}>
                  <p className="text-xs font-bold mb-1" style={{ color: "#0F6B5C" }}>تم العثور على {importNames.length} اسم:</p>
                  {importNames.map((n, i) => <p key={i} className="text-xs" style={{ color: INK }}>{i + 1}. {n}</p>)}
                </div>
              )}
              <Field label="لون مشترك لهذه الصفوف">
                <ColorSwatches value={bulkColor} onChange={setBulkColor} />
              </Field>
              <label className="flex items-center gap-2 text-sm font-medium" style={{ color: INK }}>
                <input type="checkbox" checked={bulkAutoRenew} onChange={(e) => setBulkAutoRenew(e.target.checked)} />
                تفعيل التجديد التلقائي لهذه الصفوف
              </label>
              <Field label="ملاحظة طبية / تنبيه خاص مشتركة (اختياري)" hint="تُطبَّق على كل الصفوف المستوردة — إن كانت متشابهة. يمكنك تعديلها لكل طالب لاحقًا من نافذة تعديله.">
                <textarea
                  style={{ ...inputStyle, minHeight: "60px", resize: "vertical" }}
                  value={bulkMedicalNote}
                  onChange={(e) => setBulkMedicalNote(e.target.value)}
                  placeholder="اتركه فارغًا إذا لا يوجد تنبيه مشترك"
                />
              </Field>
            </>
          )}
        </>
      )}
      <div className="flex justify-between items-center mt-4">
        <div>
          {isEdit && <button onClick={onDelete} className="px-3 py-2 rounded-lg text-sm font-medium" style={{ color: "#9A3B2E" }}>حذف الصف</button>}
        </div>
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
          <button
            disabled={isEdit ? !validSingle : (tab === "detailed" ? validDrafts.length === 0 : tab === "bulk" ? bulkNames.length === 0 : importNames.length === 0)}
            onClick={() => {
              if (isEdit) return onSaveOne(single);
              if (tab === "detailed") return onSaveMany(validDrafts);
              if (tab === "bulk") return onSaveMany(bulkNames.map((name) => ({ name, type: "text", options: [], color: bulkColor, autoRenew: bulkAutoRenew, medicalNote: bulkMedicalNote.trim() })));
              return onSaveMany(importNames.map((name) => ({ name, type: "text", options: [], color: bulkColor, autoRenew: bulkAutoRenew, medicalNote: bulkMedicalNote.trim() })));
            }}
            className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
            style={{ background: "#0F6B5C" }}
          >تم</button>
        </div>
      </div>
    </Modal>
  );
}

// ---------- Class add/edit modal ----------

const GRADE_OPTIONS = ["الأول الابتدائي","الثاني الابتدائي","الثالث الابتدائي","الرابع الابتدائي","الخامس الابتدائي","السادس الابتدائي","الأول متوسط","الثاني متوسط","الثالث متوسط","الأول ثانوي","الثاني ثانوي","الثالث ثانوي"];

function NewTermModal({ cls, onConfirm, onClose }) {
  const [yearH, setYearH] = useState("");
  const [yearG, setYearG] = useState("");
  return (
    <Modal title="بدء فصل دراسي جديد" onClose={onClose}>
      <p className="text-sm mb-4" style={{ color: MUTED }}>
        سيتم إنشاء نسخة جديدة من "{cls.subject} — {cls.grade}" بنفس الأعمدة وقائمة الطلاب، بدون أي رصد أو حضور أو أحداث سابقة. ستُؤرشف النسخة الحالية تلقائيًا كمرجع دائم لا يُلمس.
      </p>
      <div className="grid grid-cols-2 gap-3">
        <Field label="العام الدراسي الجديد (هجري)"><input style={inputStyle} value={yearH} onChange={(e) => setYearH(e.target.value)} placeholder="1448/1449" /></Field>
        <Field label="العام الدراسي الجديد (ميلادي)"><input style={inputStyle} value={yearG} onChange={(e) => setYearG(e.target.value)} placeholder="2027/2028" /></Field>
      </div>
      <div className="flex justify-end gap-2 mt-4">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button
          onClick={() => onConfirm({ yearHijri: yearH.trim(), yearGregorian: yearG.trim() })}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white"
          style={{ background: "#0F6B5C" }}
        >بدء الفصل الجديد</button>
      </div>
    </Modal>
  );
}

function ClassModal({ initial, onClose, onSave, existingClasses = [], onDuplicateExisting }) {
  const [subject, setSubject] = useState(initial?.subject || "");
  const [grade, setGrade] = useState(initial?.grade || "");
  const [teacher, setTeacher] = useState(initial?.teacher || "");
  const [yearH, setYearH] = useState(initial?.yearHijri || "");
  const [yearG, setYearG] = useState(initial?.yearGregorian || "");
  const [color, setColor] = useState(initial?.color || COLORS[0].hex);
  const [emoji, setEmoji] = useState(initial?.emoji || "");
  const [dupId, setDupId] = useState("");
  const valid = subject.trim() && grade.trim() && teacher.trim();
  return (
    <Modal title={initial ? "تعديل بيانات الفصل" : "إضافة فصل جديد"} onClose={onClose}>
      {!initial && existingClasses.length > 0 && (
        <div className="mb-4 p-3 rounded-xl flex flex-wrap items-center gap-2" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
          <Copy size={15} color={MUTED} className="shrink-0" />
          <span className="text-sm font-medium shrink-0" style={{ color: INK }}>أو كرر فصلًا موجودًا:</span>
          <select value={dupId} onChange={(e) => setDupId(e.target.value)} style={{ ...inputStyle, width: "auto", flex: 1, minWidth: "140px" }}>
            <option value="">اختر فصلًا...</option>
            {existingClasses.map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade}</option>)}
          </select>
          <button
            disabled={!dupId}
            onClick={() => onDuplicateExisting(dupId)}
            className="px-3 py-1.5 rounded-lg text-sm font-bold text-white disabled:opacity-40 shrink-0"
            style={{ background: "#0F6B5C" }}
          >تكرار</button>
        </div>
      )}
      <Field label="اسم المادة"><input style={inputStyle} value={subject} onChange={(e) => setSubject(e.target.value)} placeholder="مثال: الرياضيات" /></Field>
      <Field label="الصف">
        <input style={inputStyle} value={grade} onChange={(e) => setGrade(e.target.value)} placeholder="مثال: الرابع الابتدائي" list="grade-list" />
        <datalist id="grade-list">{GRADE_OPTIONS.map((g) => <option key={g} value={g} />)}</datalist>
      </Field>
      <Field label="اسم المعلم"><input style={inputStyle} value={teacher} onChange={(e) => setTeacher(e.target.value)} placeholder="مثال: أ. محمد العتيبي" /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="العام الدراسي (هجري)"><input style={inputStyle} value={yearH} onChange={(e) => setYearH(e.target.value)} placeholder="1447/1448" /></Field>
        <Field label="العام الدراسي (ميلادي)"><input style={inputStyle} value={yearG} onChange={(e) => setYearG(e.target.value)} placeholder="2026/2027" /></Field>
      </div>
      <Field label="لون الفصل"><ColorSwatches value={color} onChange={setColor} /></Field>
      <Field label="أيقونة الفصل (اختياري)">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {CLASS_EMOJIS.map((em) => (
            <button
              key={em}
              type="button"
              onClick={() => setEmoji(em === emoji ? "" : em)}
              className="w-9 h-9 rounded-lg flex items-center justify-center text-lg"
              style={{ border: `1px solid ${em === emoji ? "#0F6B5C" : LINE}`, background: em === emoji ? "#EAF3F0" : "#fff" }}
            >{em}</button>
          ))}
        </div>
        <input style={{ ...inputStyle, width: 90 }} value={emoji} onChange={(e) => setEmoji(e.target.value.slice(0, 2))} placeholder="أو اكتب رمزًا" />
      </Field>
      <div className="flex justify-end gap-2 mt-5">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button disabled={!valid} onClick={() => onSave({ subject: subject.trim(), grade: grade.trim(), teacher: teacher.trim(), yearHijri: yearH.trim(), yearGregorian: yearG.trim(), color, emoji })}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40" style={{ background: "#0F6B5C" }}>تم</button>
      </div>
    </Modal>
  );
}

// ---------- Dropdown cell with per-option colors ----------

function emptyTestQuestion() {
  return { id: uid(), text: "", options: [{ id: uid(), text: "" }, { id: uid(), text: "" }], correctOptionId: null };
}

function TestBuilderModal({ onSave, onClose }) {
  const [title, setTitle] = useState("");
  const [questions, setQuestions] = useState([emptyTestQuestion()]);

  const updateQuestion = (qid, patch) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, ...patch } : q)));
  const updateOption = (qid, oid, text) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, options: q.options.map((o) => (o.id === oid ? { ...o, text } : o)) } : q)));
  const addOption = (qid) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, options: [...q.options, { id: uid(), text: "" }] } : q)));
  const removeOption = (qid, oid) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, options: q.options.filter((o) => o.id !== oid), correctOptionId: q.correctOptionId === oid ? null : q.correctOptionId } : q)));
  const addQuestion = () => setQuestions((qs) => [...qs, emptyTestQuestion()]);
  const removeQuestion = (qid) => setQuestions((qs) => qs.filter((q) => q.id !== qid));

  const valid = title.trim() && questions.length > 0 && questions.every((q) => q.text.trim() && q.options.length >= 2 && q.options.every((o) => o.text.trim()) && q.correctOptionId);

  return (
    <Modal title="إنشاء اختبار جديد" onClose={onClose} wide>
      <Field label="عنوان الاختبار"><input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="مثال: اختبار الوحدة الأولى" /></Field>
      <div className="space-y-4 mb-4">
        {questions.map((q, qi) => (
          <div key={q.id} className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <div className="flex items-center gap-2 mb-2">
              <span className="text-xs font-bold px-2 py-1 rounded-full shrink-0" style={{ background: "#F3F1E9", color: MUTED }}>س{qi + 1}</span>
              <input style={{ ...inputStyle, flex: 1 }} value={q.text} onChange={(e) => updateQuestion(q.id, { text: e.target.value })} placeholder="نص السؤال" />
              {questions.length > 1 && <button onClick={() => removeQuestion(q.id)} className="p-1.5 rounded hover:bg-black/5 shrink-0"><Trash2 size={14} color="#C0392B" /></button>}
            </div>
            <div className="space-y-1.5 mr-8">
              <p className="text-xs mb-1" style={{ color: MUTED }}>حدد الدائرة أمام الإجابة الصحيحة</p>
              {q.options.map((o) => (
                <div key={o.id} className="flex items-center gap-2 p-1.5 rounded-lg" style={{ background: q.correctOptionId === o.id ? "#EAF3F0" : "transparent" }}>
                  <input type="radio" name={`correct-${q.id}`} checked={q.correctOptionId === o.id} onChange={() => updateQuestion(q.id, { correctOptionId: o.id })} />
                  <input style={{ ...inputStyle, flex: 1, padding: "6px 10px", background: "#fff" }} value={o.text} onChange={(e) => updateOption(q.id, o.id, e.target.value)} placeholder="نص الخيار" />
                  {q.options.length > 2 && <button onClick={() => removeOption(q.id, o.id)} className="p-1 rounded hover:bg-black/5 shrink-0"><X size={12} color={MUTED} /></button>}
                </div>
              ))}
              <button onClick={() => addOption(q.id)} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#0F6B5C" }}><Plus size={12} /> إضافة خيار</button>
            </div>
          </div>
        ))}
      </div>
      <button onClick={addQuestion} className="mb-4 text-sm font-semibold flex items-center gap-1" style={{ color: "#0F6B5C" }}><Plus size={15} /> إضافة سؤال</button>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button
          disabled={!valid}
          onClick={() => onSave({ id: uid(), title: title.trim(), questions, results: [], createdAt: todayKey() })}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
          style={{ background: "#0F6B5C" }}
        >حفظ الاختبار</button>
      </div>
    </Modal>
  );
}

function TestsListModal({ tests, onCreateNew, onGrade, onGradeCamera, onPrint, onDelete, onClose }) {
  return (
    <Modal title="الاختبارات" onClose={onClose} wide>
      <button onClick={onCreateNew} className="mb-4 flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>
        <Plus size={16} /> إنشاء اختبار جديد
      </button>
      {(!tests || tests.length === 0) ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد اختبارات بعد.</p>
      ) : (
        <div className="space-y-2">
          {tests.map((t) => (
            <div key={t.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
              <div className="flex-1">
                <p className="text-sm font-bold" style={{ color: INK }}>{t.title}</p>
                <p className="text-xs" style={{ color: MUTED }}>{t.questions.length} سؤال • {(t.results || []).length} طالب مصحّح</p>
              </div>
              <button onClick={() => onPrint(t.id)} title="طباعة الورقة" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Printer size={15} color={MUTED} /></button>
              <button onClick={() => onGradeCamera(t.id)} title="تصحيح بالكاميرا" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Camera size={15} color={MUTED} /></button>
              <button onClick={() => onGrade(t.id)} className="px-3 py-1.5 rounded-lg text-xs font-bold text-white shrink-0" style={{ background: "#0F6B5C" }}>تصحيح</button>
              <button onClick={() => onDelete(t.id)} className="p-1.5 rounded hover:bg-black/5 shrink-0"><Trash2 size={14} color="#C0392B" /></button>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function AnswerOption({ label, selected, onSelect, color }) {
  const c = color || "#0F6B5C";
  return (
    <button
      type="button"
      onClick={onSelect}
      className="w-full flex items-center gap-2.5 px-3 py-2.5 rounded-xl text-sm text-right transition-colors"
      style={{ border: `1.5px solid ${selected ? c : LINE}`, background: selected ? `${c}1A` : "#fff", color: selected ? c : INK, fontWeight: selected ? 700 : 400 }}
    >
      <span className="w-4 h-4 rounded-full flex items-center justify-center shrink-0" style={{ border: `2px solid ${selected ? c : LINE}` }}>
        {selected && <span className="w-2 h-2 rounded-full" style={{ background: c }} />}
      </span>
      <span className="flex-1">{label}</span>
    </button>
  );
}

// Builds a printable answer-sheet page for a test: title, a name line (either
// blank for students to fill in, or pre-filled for a specific student), and
// every question with empty circular bubbles next to each option to mark.
// Builds an official remedial-plan / weakness-report document: school header
// (logo/country/ministry/school), student meta, structured sections, and a
// teacher/principal signature footer — for either working use or sending to admin.
function buildRemedialPlanCanvas({
  countryName, ministryName, schoolName, logoImageElement,
  title, studentName, showName, className, teacherName, principalName, date,
  subject, weaknessSigns, causes, goals, methods, followUp, notes,
}) {
  const width = 900;
  const scale = 2;
  const measureCanvas = document.createElement("canvas");
  const mctx = measureCanvas.getContext("2d");
  mctx.font = "15px Tahoma, Arial";
  const contentWidth = width - 120;

  const sections = [
    { label: "مظاهر الضعف", value: weaknessSigns },
    { label: "أسباب الضعف", value: causes },
    { label: "الأهداف العلاجية", value: goals },
    { label: "الأساليب والأنشطة العلاجية", value: methods },
    { label: "المتابعة والتقويم", value: followUp },
    ...(notes && notes.trim() ? [{ label: "ملاحظات وتوصيات", value: notes }] : []),
  ].map((s) => ({ ...s, lines: wrapCanvasText(mctx, s.value && s.value.trim() ? s.value : "—", contentWidth) }));

  let headerHeight = 40;
  if (logoImageElement) headerHeight += 62;
  if (countryName) headerHeight += 18;
  if (ministryName) headerHeight += 18;
  if (schoolName) headerHeight += 22;
  headerHeight += 40; // title
  const metaHeight = 4 * 22 + 16;
  const sectionsHeight = sections.reduce((s, b) => s + 24 + b.lines.length * 22 + 18, 0);
  const height = headerHeight + metaHeight + sectionsHeight + 140;

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textAlign = "center";

  let hy = 40;
  if (logoImageElement) {
    ctx.drawImage(logoImageElement, width / 2 - 26, hy - 26, 52, 52);
    hy += 40;
  }
  if (countryName) { ctx.font = "bold 14px Tahoma, Arial"; ctx.fillStyle = "#232622"; ctx.fillText(countryName, width / 2, hy); hy += 18; }
  if (ministryName) { ctx.font = "13px Tahoma, Arial"; ctx.fillStyle = "#7A7768"; ctx.fillText(ministryName, width / 2, hy); hy += 18; }
  if (schoolName) { ctx.font = "14px Tahoma, Arial"; ctx.fillStyle = "#232622"; ctx.fillText(schoolName, width / 2, hy); hy += 24; }

  ctx.font = "bold 22px Tahoma, Arial";
  ctx.fillStyle = "#0F6B5C";
  ctx.fillText(title, width / 2, hy + 12);
  hy += 40;

  ctx.textAlign = "right";
  ctx.font = "14px Tahoma, Arial";
  ctx.fillStyle = "#232622";
  ctx.fillText(showName && studentName ? `اسم الطالب: ${studentName}` : "اسم الطالب: ....................................", width - 60, hy);
  hy += 22;
  ctx.fillText(`المادة: ${subject || "—"}`, width - 60, hy);
  hy += 22;
  ctx.fillText(`الصف: ${className || "—"}`, width - 60, hy);
  hy += 22;
  ctx.fillText(`التاريخ: ${date}`, width - 60, hy);
  hy += 30;
  ctx.strokeStyle = "#D8D2C0";
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(60, hy); ctx.lineTo(width - 60, hy); ctx.stroke();
  hy += 24;

  sections.forEach((s) => {
    ctx.font = "bold 16px Tahoma, Arial";
    ctx.fillStyle = "#0F6B5C";
    ctx.fillText(s.label, width - 60, hy);
    hy += 24;
    ctx.font = "14px Tahoma, Arial";
    ctx.fillStyle = "#232622";
    s.lines.forEach((line) => { ctx.fillText(line, width - 60, hy); hy += 22; });
    hy += 18;
  });

  hy = height - 90;
  ctx.textAlign = "center";
  const colW = 260;
  if (teacherName) {
    ctx.font = "12px Tahoma, Arial"; ctx.fillStyle = "#7A7768";
    ctx.fillText("المعلم", width - 120 - colW / 2, hy);
    ctx.font = "bold 15px Tahoma, Arial"; ctx.fillStyle = "#232622";
    ctx.fillText(teacherName, width - 120 - colW / 2, hy + 22);
  }
  if (principalName) {
    ctx.font = "12px Tahoma, Arial"; ctx.fillStyle = "#7A7768";
    ctx.fillText("مدير المدرسة", 120 + colW / 2, hy);
    ctx.font = "bold 15px Tahoma, Arial"; ctx.fillStyle = "#232622";
    ctx.fillText(principalName, 120 + colW / 2, hy + 22);
  }

  return { canvas, logicalWidth: width, logicalHeight: height };
}

// ---------- Camera-based auto-grading (OMR) ----------
// Computes bubble center positions as fractions (0-1) of the sheet's content
// area, shared by BOTH the print builder and the scanner so printed and
// detected coordinates are always consistent by construction.
function computeOMRGrid(questionCount, optionCount) {
  const headerFrac = 0.14;
  const bottomMargin = 0.03;
  const gridHeight = 1 - headerFrac - bottomMargin;
  const rowHeight = gridHeight / questionCount;
  const leftMargin = 0.08, rightMargin = 0.14;
  const usableWidth = 1 - leftMargin - rightMargin;
  const colWidth = usableWidth / optionCount;
  const rows = [];
  for (let q = 0; q < questionCount; q++) {
    const y = headerFrac + rowHeight * (q + 0.5);
    const options = [];
    for (let o = 0; o < optionCount; o++) {
      // RTL: option 0 sits at the rightmost position
      const x = 1 - rightMargin - colWidth * (o + 0.5);
      options.push({ oIndex: o, x });
    }
    rows.push({ qIndex: q, y, options });
  }
  return { rows };
}

const OMR_SHEET_W = 900, OMR_SHEET_H = 1200, OMR_MARGIN = 22;

// Encodes a student's index within their class roster as a row of small
// printed squares (filled = 1, empty = 0) in a fixed corner — read back the
// same way as answer bubbles, so no text/QR recognition is needed to know
// whose sheet this is. Supports rosters up to 2^ID_BITS - 1 students.
const ID_BITS = 7;
function computeIDMarkerPositions() {
  const positions = [];
  const startX = 0.07, y = 0.045, gap = 0.03;
  for (let i = 0; i < ID_BITS; i++) positions.push({ x: startX + i * gap, y });
  return positions;
}

// Printable OMR answer sheet: outer border the student/teacher aligns to,
// corner marks for visual reference, and empty bubbles (labels sit ABOVE
// each bubble, never inside it, so an unmarked bubble is always pure blank
// white — critical for reliable darkness sampling later).
function buildOMRSheetCanvas({ title, studentName, questionCount, optionCount, studentIndex }) {
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = OMR_SHEET_W * scale;
  canvas.height = OMR_SHEET_H * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, OMR_SHEET_W, OMR_SHEET_H);

  ctx.strokeStyle = "#000";
  ctx.lineWidth = 3;
  ctx.strokeRect(OMR_MARGIN, OMR_MARGIN, OMR_SHEET_W - 2 * OMR_MARGIN, OMR_SHEET_H - 2 * OMR_MARGIN);
  const mk = 20;
  ctx.fillStyle = "#000";
  [[OMR_MARGIN, OMR_MARGIN], [OMR_SHEET_W - OMR_MARGIN - mk, OMR_MARGIN], [OMR_MARGIN, OMR_SHEET_H - OMR_MARGIN - mk], [OMR_SHEET_W - OMR_MARGIN - mk, OMR_SHEET_H - OMR_MARGIN - mk]]
    .forEach(([x, y]) => ctx.fillRect(x, y, mk, mk));

  const cw = OMR_SHEET_W - 2 * OMR_MARGIN, ch = OMR_SHEET_H - 2 * OMR_MARGIN;

  if (typeof studentIndex === "number") {
    const bits = studentIndex.toString(2).padStart(ID_BITS, "0").split("").map(Number);
    computeIDMarkerPositions().forEach((pos, i) => {
      const px = OMR_MARGIN + pos.x * cw, py = OMR_MARGIN + pos.y * ch;
      const size = 11;
      ctx.strokeStyle = "#000"; ctx.lineWidth = 1.2;
      ctx.strokeRect(px - size / 2, py - size / 2, size, size);
      if (bits[i] === 1) {
        ctx.fillStyle = "#000";
        ctx.fillRect(px - size / 2 + 1.5, py - size / 2 + 1.5, size - 3, size - 3);
      }
    });
  }

  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.fillStyle = "#232622";
  ctx.font = "bold 22px Tahoma, Arial";
  ctx.fillText(title, OMR_SHEET_W / 2, OMR_MARGIN + 42);
  ctx.font = "15px Tahoma, Arial";
  ctx.fillText(`اسم الطالب: ${studentName || "......................................"}`, OMR_SHEET_W / 2, OMR_MARGIN + 70);

  const grid = computeOMRGrid(questionCount, optionCount);
  const letters = ["أ", "ب", "ج", "د", "هـ", "و"];
  grid.rows.forEach((row) => {
    const py = OMR_MARGIN + row.y * ch;
    ctx.textAlign = "right";
    ctx.font = "bold 14px Tahoma, Arial";
    ctx.fillStyle = "#232622";
    ctx.fillText(`${row.qIndex + 1}`, OMR_SHEET_W - OMR_MARGIN - 6, py + 5);
    row.options.forEach((opt) => {
      const px = OMR_MARGIN + opt.x * cw;
      ctx.textAlign = "center";
      ctx.font = "11px Tahoma, Arial";
      ctx.fillStyle = "#7A7768";
      ctx.fillText(letters[opt.oIndex] || String(opt.oIndex + 1), px, py - 15);
      ctx.beginPath();
      ctx.arc(px, py, 11, 0, Math.PI * 2);
      ctx.strokeStyle = "#000";
      ctx.lineWidth = 1.5;
      ctx.stroke();
    });
  });

  return { canvas, logicalWidth: OMR_SHEET_W, logicalHeight: OMR_SHEET_H };
}

// Captures the current video frame cropped to a target aspect ratio, using
// the SAME crop math as CSS `object-fit: cover` — so what the user aligned
// on screen is exactly what gets analyzed, regardless of camera resolution.
function captureCroppedFrame(videoEl, targetAspect) {
  const vw = videoEl.videoWidth, vh = videoEl.videoHeight;
  const videoAspect = vw / vh;
  let sx, sy, sw, sh;
  if (videoAspect > targetAspect) {
    sh = vh; sw = vh * targetAspect; sx = (vw - sw) / 2; sy = 0;
  } else {
    sw = vw; sh = vw / targetAspect; sx = 0; sy = (vh - sh) / 2;
  }
  const canvas = document.createElement("canvas");
  canvas.width = OMR_SHEET_W;
  canvas.height = OMR_SHEET_H;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(videoEl, sx, sy, sw, sh, 0, 0, OMR_SHEET_W, OMR_SHEET_H);
  return canvas;
}

function sampleBubbleDarkness(ctx, cx, cy, radius) {
  const x = Math.max(0, Math.round(cx - radius)), y = Math.max(0, Math.round(cy - radius));
  const size = radius * 2;
  let imgData;
  try {
    imgData = ctx.getImageData(x, y, size, size);
  } catch (e) {
    return 0;
  }
  let total = 0, count = 0;
  for (let i = 0; i < imgData.data.length; i += 4) {
    total += (imgData.data[i] + imgData.data[i + 1] + imgData.data[i + 2]) / 3;
    count++;
  }
  return count ? 255 - total / count : 0;
}

// Analyzes a captured OMR sheet frame and returns { qIndex: oIndex|null }.
// null means the two darkest bubbles for that question were too close in
// darkness to confidently tell them apart — flagged for manual review.
function analyzeOMRFrame(canvas, questionCount, optionCount) {
  const ctx = canvas.getContext("2d");
  const cw = OMR_SHEET_W - 2 * OMR_MARGIN, ch = OMR_SHEET_H - 2 * OMR_MARGIN;
  const grid = computeOMRGrid(questionCount, optionCount);
  const results = {};
  grid.rows.forEach((row) => {
    const py = OMR_MARGIN + row.y * ch;
    const darks = row.options.map((opt) => ({
      oIndex: opt.oIndex,
      d: sampleBubbleDarkness(ctx, OMR_MARGIN + opt.x * cw, py, 11),
    })).sort((a, b) => b.d - a.d);
    const best = darks[0], second = darks[1];
    results[row.qIndex] = best && best.d > 40 && (best.d - (second?.d || 0)) > 18 ? best.oIndex : null;
  });
  return results;
}

// Reads back the printed ID marker squares (solid black vs blank) and
// decodes them to the student's roster index. Printed squares have high
// contrast (pure ink vs pure paper), so a simple absolute threshold works
// well here — unlike hand-filled answer bubbles.
function decodeStudentIndex(canvas) {
  const ctx = canvas.getContext("2d");
  const cw = OMR_SHEET_W - 2 * OMR_MARGIN, ch = OMR_SHEET_H - 2 * OMR_MARGIN;
  let bits = "";
  computeIDMarkerPositions().forEach((pos) => {
    const d = sampleBubbleDarkness(ctx, OMR_MARGIN + pos.x * cw, OMR_MARGIN + pos.y * ch, 6);
    bits += d > 60 ? "1" : "0";
  });
  return parseInt(bits, 2);
}

function buildTestPaperCanvas({ title, studentName, questions }) {
  const width = 900;
  const scale = 2;
  const measureCanvas = document.createElement("canvas");
  const mctx = measureCanvas.getContext("2d");
  mctx.font = "16px Tahoma, Arial";
  const contentWidth = width - 120;
  const letters = ["أ", "ب", "ج", "د", "هـ", "و", "ز", "ح"];

  const blocks = questions.map((q, qi) => {
    const qLines = wrapCanvasText(mctx, `${qi + 1}. ${q.text}`, contentWidth);
    return { qLines, optCount: q.options.length };
  });
  const contentHeight = blocks.reduce((s, b) => s + b.qLines.length * 24 + 10 + b.optCount * 30 + 16, 0);
  const height = 160 + contentHeight + 40;

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";

  ctx.textAlign = "center";
  ctx.fillStyle = "#232622";
  ctx.font = "bold 26px Tahoma, Arial";
  ctx.fillText(title, width / 2, 50);

  ctx.textAlign = "right";
  ctx.font = "16px Tahoma, Arial";
  ctx.fillText(`اسم الطالب: ${studentName || "......................................"}`, width - 60, 90);
  ctx.strokeStyle = "#232622";
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(60, 112); ctx.lineTo(width - 60, 112); ctx.stroke();

  let y = 150;
  questions.forEach((q, qi) => {
    const block = blocks[qi];
    ctx.font = "bold 16px Tahoma, Arial";
    ctx.fillStyle = "#232622";
    block.qLines.forEach((line) => { ctx.fillText(line, width - 60, y); y += 24; });
    y += 6;
    ctx.font = "15px Tahoma, Arial";
    q.options.forEach((o, oi) => {
      ctx.beginPath();
      ctx.arc(width - 80, y - 5, 7, 0, Math.PI * 2);
      ctx.stroke();
      ctx.fillStyle = "#232622";
      ctx.fillText(`${letters[oi] || oi + 1}) ${o.text}`, width - 95, y);
      y += 28;
    });
    y += 16;
  });

  return { canvas, logicalWidth: width, logicalHeight: height };
}

function PrintTestModal({ test, classes, onClose }) {
  const [mode, setMode] = useState("blank");
  const [classId, setClassId] = useState(classes.find((c) => !c.archived)?.id || "");
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);
  const optionCount = test.questions[0]?.options.length || 0;
  const omrSupported = test.questions.length > 0 && test.questions.every((q) => q.options.length === optionCount);
  const [omrFormat, setOmrFormat] = useState(false);

  const buildSheet = (studentName, index) => omrFormat
    ? buildOMRSheetCanvas({ title: test.title, studentName, questionCount: test.questions.length, optionCount, studentIndex: index })
    : buildTestPaperCanvas({ title: test.title, studentName, questions: test.questions });

  const downloadBlank = () => {
    const built = buildSheet(null);
    built.canvas.toBlob((blob) => downloadBlob(blob, `${test.title}.png`));
  };

  const downloadPerStudent = async () => {
    const cls = classes.find((c) => c.id === classId);
    if (!cls || cls.rows.length === 0) return;
    if (omrFormat && cls.rows.length > (1 << ID_BITS) - 1) {
      setGenerating(false);
      return;
    }
    setGenerating(true);
    setProgress(0);
    const files = [];
    for (let i = 0; i < cls.rows.length; i++) {
      const row = cls.rows[i];
      const built = buildSheet(row.name, i);
      const blob = await new Promise((resolve) => built.canvas.toBlob(resolve, "image/png"));
      const buf = new Uint8Array(await blob.arrayBuffer());
      files.push({ name: `${row.name}.png`, data: buf });
      setProgress(Math.round(((i + 1) / cls.rows.length) * 100));
    }
    const zipBlob = buildZip(files);
    downloadBlob(zipBlob, `${test.title}-أوراق.zip`);
    setGenerating(false);
    onClose();
  };

  return (
    <Modal title={`طباعة — ${test.title}`} onClose={onClose}>
      {omrSupported && (
        <label className="flex items-start gap-2 text-sm font-medium mb-4 p-3 rounded-xl" style={{ color: INK, background: "#F3F1E9", border: `1px solid ${LINE}` }}>
          <input type="checkbox" checked={omrFormat} onChange={(e) => setOmrFormat(e.target.checked)} className="mt-0.5" />
          <span>
            تنسيق التصحيح بالكاميرا (OMR)
            <span className="block text-xs font-normal mt-0.5" style={{ color: MUTED }}>شبكة فقاعات دقيقة الأبعاد مع إطار محاذاة — استخدمها مع ميزة "تصحيح بالكاميرا".</span>
          </span>
        </label>
      )}
      <div className="flex gap-2 mb-4">
        <button onClick={() => setMode("blank")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: mode === "blank" ? INK : "transparent", color: mode === "blank" ? "#fff" : MUTED, border: `1px solid ${mode === "blank" ? INK : LINE}` }}>نسخة عامة بدون أسماء</button>
        <button onClick={() => setMode("perStudent")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: mode === "perStudent" ? INK : "transparent", color: mode === "perStudent" ? "#fff" : MUTED, border: `1px solid ${mode === "perStudent" ? INK : LINE}` }}>نسخة لكل طالب باسمه</button>
      </div>
      {mode === "blank" ? (
        <button onClick={downloadBlank} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>
          <ImageDown size={16} /> تنزيل نسخة فارغة
        </button>
      ) : (
        <>
          <Field label="الفصل">
            <select value={classId} onChange={(e) => setClassId(e.target.value)} style={inputStyle}>
              {classes.filter((c) => !c.archived).map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade} ({c.rows.length} طالب)</option>)}
            </select>
          </Field>
          {generating ? (
            <div className="p-3 rounded-xl flex items-center gap-2" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
              <div className="w-4 h-4 rounded-full border-2 shrink-0" style={{ borderColor: "#0F6B5C transparent #0F6B5C #0F6B5C", animation: "spin 0.8s linear infinite" }} />
              <p className="text-xs font-semibold" style={{ color: INK }}>جارٍ التجهيز... {progress}%</p>
            </div>
          ) : (
            <button onClick={downloadPerStudent} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>
              <ImageDown size={16} /> تنزيل نسخة لكل طالب (ملف مضغوط)
            </button>
          )}
        </>
      )}
    </Modal>
  );
}

function OMRScanModal({ test, classes, onSaveResult, onApplyToClass, onClose }) {
  const optionCount = test.questions[0]?.options.length || 0;
  const uniform = test.questions.length > 0 && test.questions.every((q) => q.options.length === optionCount);

  const [step, setStep] = useState(uniform ? "setup" : "unsupported");
  const [classId, setClassId] = useState(classes.find((c) => !c.archived)?.id || "");
  const [colId, setColId] = useState("");
  const [valueMode, setValueMode] = useState("score");
  const [cameraError, setCameraError] = useState("");
  const [scanResult, setScanResult] = useState(null); // { studentIndex, answers, score, total, percentage }
  const [manualRowId, setManualRowId] = useState("");
  const [postedCount, setPostedCount] = useState(0);
  const videoRef = useRef(null);
  const streamRef = useRef(null);

  const cls = classes.find((c) => c.id === classId);
  const applicableCols = cls ? cls.columns.filter((c) => c.type !== "dropdown") : [];

  const stopCamera = () => {
    if (streamRef.current) { streamRef.current.getTracks().forEach((t) => t.stop()); streamRef.current = null; }
  };

  const startCamera = async () => {
    setCameraError("");
    setScanResult(null);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      streamRef.current = stream;
      setStep("camera");
    } catch (e) {
      setCameraError("تعذّر الوصول للكاميرا. تأكد من منح الإذن للمتصفح (وأن الصفحة مفتوحة بشاشة كاملة لا داخل نافذة مصغّرة)، أو جرّب من جهاز آخر.");
    }
  };

  useEffect(() => {
    if (step === "camera" && streamRef.current && videoRef.current) {
      videoRef.current.srcObject = streamRef.current;
    }
  }, [step]);

  useEffect(() => () => { if (streamRef.current) streamRef.current.getTracks().forEach((t) => t.stop()); }, []);

  const capture = () => {
    if (!videoRef.current || !cls) return;
    const canvas = captureCroppedFrame(videoRef.current, OMR_SHEET_W / OMR_SHEET_H);
    const idx = decodeStudentIndex(canvas);
    const detected = analyzeOMRFrame(canvas, test.questions.length, optionCount);
    const answers = {};
    test.questions.forEach((q, qi) => {
      const oIdx = detected[qi];
      if (oIdx !== null && oIdx !== undefined && q.options[oIdx]) answers[q.id] = q.options[oIdx].id;
    });
    let score = 0;
    test.questions.forEach((q) => { if (answers[q.id] === q.correctOptionId) score++; });
    const total = test.questions.length;
    const percentage = Math.round((score / total) * 100);
    const matchedRow = cls.rows[idx];
    setScanResult({ answers, score, total, percentage, unclear: Object.keys(answers).length < total });
    setManualRowId(matchedRow ? matchedRow.id : "");
    stopCamera();
    setStep("result");
  };

  const recordAndScanNext = () => {
    const row = cls.rows.find((r) => r.id === manualRowId);
    if (!row || !colId) return;
    const value = valueMode === "score" ? String(scanResult.score) : `${scanResult.percentage}%`;
    onSaveResult({ id: uid(), studentName: row.name, answers: scanResult.answers, score: scanResult.score, total: scanResult.total, percentage: scanResult.percentage, when: nowMeta() });
    onApplyToClass(classId, colId, row.name, value);
    setPostedCount((n) => n + 1);
    startCamera();
  };

  return (
    <Modal title={`تصحيح بالكاميرا — ${test.title}`} onClose={() => { stopCamera(); onClose(); }} wide>
      {step === "unsupported" && (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>
          التصحيح بالكاميرا يتطلب عددًا متساويًا من الخيارات لكل الأسئلة. عدّل الاختبار أو استخدم "تصحيح" العادي بدلًا من ذلك.
        </p>
      )}

      {step === "setup" && (
        <>
          <p className="text-sm mb-4" style={{ color: MUTED }}>
            اطبع أوراق الطلاب بتنسيق OMR (زر الطباعة → "نسخة لكل طالب باسمه" مع تفعيل تنسيق الكاميرا) — كل ورقة تُعرَف تلقائيًا لصاحبها، فلا تحتاج كتابة الاسم يدويًا. اختر الفصل والعمود مرة واحدة، وبعدها صوّر كل الأوراق ورقة بعد ورقة.
          </p>
          <Field label="الفصل">
            <select value={classId} onChange={(e) => { setClassId(e.target.value); setColId(""); }} style={inputStyle}>
              {classes.filter((c) => !c.archived).map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade}</option>)}
            </select>
          </Field>
          <Field label="العمود الذي تُرصَد فيه الدرجة">
            <select value={colId} onChange={(e) => setColId(e.target.value)} style={inputStyle}>
              <option value="">اختر عمودًا...</option>
              {applicableCols.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="القيمة المُرصَدة">
            <div className="flex gap-2">
              <button onClick={() => setValueMode("score")} className="flex-1 text-xs font-semibold py-1.5 rounded-lg" style={{ background: valueMode === "score" ? INK : "transparent", color: valueMode === "score" ? "#fff" : MUTED, border: `1px solid ${valueMode === "score" ? INK : LINE}` }}>الدرجة (مثال: 8/10)</button>
              <button onClick={() => setValueMode("percentage")} className="flex-1 text-xs font-semibold py-1.5 rounded-lg" style={{ background: valueMode === "percentage" ? INK : "transparent", color: valueMode === "percentage" ? "#fff" : MUTED, border: `1px solid ${valueMode === "percentage" ? INK : LINE}` }}>النسبة المئوية</button>
            </div>
          </Field>
          <button disabled={!colId} onClick={startCamera} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white disabled:opacity-40 mt-2" style={{ background: "#0F6B5C" }}>
            <Camera size={16} /> بدء التصحيح
          </button>
          {cameraError && <p className="text-xs mt-2 text-center" style={{ color: "#C0392B" }}>{cameraError}</p>}
        </>
      )}

      {step === "camera" && (
        <>
          {postedCount > 0 && <p className="text-xs text-center mb-2 font-semibold" style={{ color: "#0F6B5C" }}>تم رصد {postedCount} ورقة حتى الآن</p>}
          <p className="text-sm mb-3 text-center" style={{ color: MUTED }}>
            حاذِ إطار الورقة بالكامل مع حدود المستطيل الظاهر، بإضاءة جيدة، ثم التقط.
          </p>
          <div className="relative mx-auto mb-4 rounded-xl overflow-hidden" style={{ width: "100%", maxWidth: 340, aspectRatio: `${OMR_SHEET_W} / ${OMR_SHEET_H}`, background: "#000" }}>
            <video ref={videoRef} autoPlay playsInline muted style={{ width: "100%", height: "100%", objectFit: "cover" }} />
            <div className="absolute inset-2 pointer-events-none rounded-lg" style={{ border: "3px solid #F2C94C" }} />
          </div>
          <div className="flex gap-2">
            <button onClick={() => { stopCamera(); setStep("setup"); }} className="flex-1 py-2.5 rounded-xl text-sm font-medium" style={{ color: MUTED, border: `1px solid ${LINE}` }}>إنهاء</button>
            <button onClick={capture} className="flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>
              <Camera size={16} /> التقط
            </button>
          </div>
        </>
      )}

      {step === "result" && scanResult && cls && (
        <>
          <div className="text-center p-6 rounded-2xl mb-4" style={{ background: "#EAF3F0" }}>
            <p className="text-3xl font-extrabold" style={{ color: "#0F6B5C" }}>{scanResult.score} / {scanResult.total}</p>
            <p className="text-sm mt-1" style={{ color: "#0F6B5C" }}>{scanResult.percentage}%</p>
          </div>
          <Field label="الطالب المكتشف من الورقة" hint={scanResult.unclear ? "تعذّرت قراءة بعض الإجابات بوضوح — تأكد من الدرجة أعلاه." : "تأكد أن الاسم صحيح قبل الرصد؛ صحّحه من القائمة إن لزم."}>
            <select value={manualRowId} onChange={(e) => setManualRowId(e.target.value)} style={inputStyle}>
              <option value="">— لم يُتعرّف على طالب مطابق —</option>
              {cls.rows.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
            </select>
          </Field>
          <div className="flex gap-2 mt-2">
            <button onClick={() => { setStep("camera"); startCamera(); }} className="flex-1 py-2.5 rounded-xl text-sm font-medium" style={{ color: MUTED, border: `1px solid ${LINE}` }}>إعادة التصوير</button>
            <button disabled={!manualRowId} onClick={recordAndScanNext} className="flex-1 py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40" style={{ background: "#0F6B5C" }}>
              رصد ✓ والانتقال للتالي
            </button>
          </div>
        </>
      )}
    </Modal>
  );
}

function GradeTestModal({ test, classes, onSaveResult, onApplyToClass, onClose }) {
  const [studentName, setStudentName] = useState("");
  const [answers, setAnswers] = useState({});
  const [graded, setGraded] = useState(null);
  const [applyClassId, setApplyClassId] = useState("");
  const [applyColId, setApplyColId] = useState("");
  const [valueMode, setValueMode] = useState("score");

  const allNames = Array.from(new Set(classes.flatMap((c) => c.rows.map((r) => r.name))));
  const setAnswer = (qid, oid) => setAnswers((a) => ({ ...a, [qid]: oid }));

  const grade = () => {
    let score = 0;
    test.questions.forEach((q) => { if (answers[q.id] === q.correctOptionId) score++; });
    const total = test.questions.length;
    const percentage = Math.round((score / total) * 100);
    const result = { id: uid(), studentName: studentName.trim(), answers, score, total, percentage, when: nowMeta() };
    setGraded(result);
    onSaveResult(result);
  };

  const applyClass = classes.find((c) => c.id === applyClassId);
  const applicableCols = applyClass ? applyClass.columns.filter((c) => c.type !== "dropdown") : [];

  return (
    <Modal title={`تصحيح — ${test.title}`} onClose={onClose} wide>
      {!graded ? (
        <>
          <Field label="اسم الطالب">
            <input list="test-student-names" style={inputStyle} value={studentName} onChange={(e) => setStudentName(e.target.value)} placeholder="اكتب أو اختر اسم الطالب" />
            <datalist id="test-student-names">{allNames.map((n) => <option key={n} value={n} />)}</datalist>
          </Field>
          <div className="space-y-3 mb-4">
            {test.questions.map((q, qi) => (
              <div key={q.id} className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}` }}>
                <p className="text-sm font-semibold mb-2" style={{ color: INK }}>{qi + 1}. {q.text}</p>
                <div className="space-y-1.5">
                  {q.options.map((o) => (
                    <AnswerOption key={o.id} label={o.text} selected={answers[q.id] === o.id} onSelect={() => setAnswer(q.id, o.id)} />
                  ))}
                </div>
              </div>
            ))}
          </div>
          <button
            disabled={!studentName.trim() || Object.keys(answers).length < test.questions.length}
            onClick={grade}
            className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40"
            style={{ background: "#0F6B5C" }}
          >تصحيح الآن</button>
        </>
      ) : (
        <>
          <div className="text-center p-6 rounded-2xl mb-4" style={{ background: "#EAF3F0" }}>
            <p className="text-3xl font-extrabold" style={{ color: "#0F6B5C" }}>{graded.score} / {graded.total}</p>
            <p className="text-sm mt-1" style={{ color: "#0F6B5C" }}>{graded.percentage}% — {graded.studentName}</p>
          </div>
          <p className="text-sm font-semibold mb-2" style={{ color: INK }}>إضافة النتيجة إلى فصل (اختياري)</p>
          <Field label="الفصل">
            <select value={applyClassId} onChange={(e) => { setApplyClassId(e.target.value); setApplyColId(""); }} style={inputStyle}>
              <option value="">بدون — احتفظ بها هنا فقط</option>
              {classes.filter((c) => !c.archived).map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade}</option>)}
            </select>
          </Field>
          {applyClassId && (
            <>
              <Field label="العمود">
                <select value={applyColId} onChange={(e) => setApplyColId(e.target.value)} style={inputStyle}>
                  <option value="">اختر عمودًا...</option>
                  {applicableCols.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
              </Field>
              <Field label="القيمة المُرصَدة">
                <div className="flex gap-2">
                  <button onClick={() => setValueMode("score")} className="flex-1 text-xs font-semibold py-1.5 rounded-lg" style={{ background: valueMode === "score" ? INK : "transparent", color: valueMode === "score" ? "#fff" : MUTED, border: `1px solid ${valueMode === "score" ? INK : LINE}` }}>الدرجة ({graded.score}/{graded.total})</button>
                  <button onClick={() => setValueMode("percentage")} className="flex-1 text-xs font-semibold py-1.5 rounded-lg" style={{ background: valueMode === "percentage" ? INK : "transparent", color: valueMode === "percentage" ? "#fff" : MUTED, border: `1px solid ${valueMode === "percentage" ? INK : LINE}` }}>النسبة ({graded.percentage}%)</button>
                </div>
              </Field>
              <button
                disabled={!applyColId}
                onClick={() => { onApplyToClass(applyClassId, applyColId, graded.studentName, valueMode === "score" ? String(graded.score) : `${graded.percentage}%`); onClose(); }}
                className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 mt-2"
                style={{ background: "#0F6B5C" }}
              >إضافة إلى الفصل</button>
            </>
          )}
          <button onClick={onClose} className="w-full py-2 mt-2 rounded-xl text-sm font-medium" style={{ color: MUTED }}>إغلاق بدون إضافة</button>
        </>
      )}
    </Modal>
  );
}

function MoveStudentsModal({ currentClassId, allClasses, count, onMove, onClose }) {
  const otherClasses = allClasses.filter((c) => c.id !== currentClassId && !c.archived);
  const [destId, setDestId] = useState(otherClasses[0]?.id || "");
  const [includeGrades, setIncludeGrades] = useState(true);

  if (otherClasses.length === 0) {
    return (
      <Modal title="نقل إلى فصل آخر" onClose={onClose}>
        <p className="text-sm text-center py-6" style={{ color: MUTED }}>لا يوجد فصل آخر نشط تنقل إليه الطلاب. أنشئ فصلًا جديدًا أولًا.</p>
      </Modal>
    );
  }

  return (
    <Modal title="نقل إلى فصل آخر" onClose={onClose}>
      <p className="text-sm mb-4" style={{ color: MUTED }}>سيتم نقل <b>{count}</b> طالب إلى الفصل الذي تختاره أدناه، وحذفهم من هذا الفصل.</p>
      <Field label="الفصل الوجهة">
        <select value={destId} onChange={(e) => setDestId(e.target.value)} style={inputStyle}>
          {otherClasses.map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade}</option>)}
        </select>
      </Field>
      <label className="flex items-start gap-2 text-sm font-medium mt-1" style={{ color: INK }}>
        <input type="checkbox" checked={includeGrades} onChange={(e) => setIncludeGrades(e.target.checked)} className="mt-0.5" />
        <span>
          نقل درجاتهم ورصدهم المسجّل معهم
          <span className="block text-xs font-normal mt-0.5" style={{ color: MUTED }}>يُنقل فقط ما كان له عمود بنفس الاسم في الفصل الوجهة. إن أبقيته غير مفعّل، يُنقل الطالب بصف فارغ.</span>
        </span>
      </label>
      <div className="flex justify-end gap-2 mt-5">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button
          onClick={() => { onMove(destId, includeGrades); onClose(); }}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white"
          style={{ background: "#0F6B5C" }}
        >نقل</button>
      </div>
    </Modal>
  );
}

function BulkRecordModal({ columns, onApply, onClose }) {
  const [colId, setColId] = useState(columns[0]?.id || "");
  const [val, setVal] = useState("");
  const col = columns.find((c) => c.id === colId);
  return (
    <Modal title="رصد جماعي للطلاب المحددين" onClose={onClose}>
      <Field label="العمود">
        <select value={colId} onChange={(e) => { setColId(e.target.value); setVal(""); }} style={inputStyle}>
          {columns.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
      </Field>
      <Field label="القيمة">
        {col?.type === "dropdown" ? (
          <select value={val} onChange={(e) => setVal(e.target.value)} style={inputStyle}>
            <option value="">اختر...</option>
            {(col.options || []).filter((o) => o.label.trim()).map((o) => <option key={o.id} value={o.label}>{o.label}</option>)}
          </select>
        ) : (
          <input type={col?.type === "counter" ? "number" : "text"} value={val} onChange={(e) => setVal(e.target.value)} style={inputStyle} placeholder="القيمة" />
        )}
      </Field>
      <div className="flex justify-end gap-2 mt-4">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button
          disabled={!val.trim() || !col}
          onClick={() => { onApply(col, val); onClose(); }}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
          style={{ background: "#0F6B5C" }}
        >تطبيق</button>
      </div>
    </Modal>
  );
}

function FilterModal({ columns, initial, onApply, onClear, onClose }) {
  const [colId, setColId] = useState(initial?.colId || columns[0]?.id || "");
  const [operator, setOperator] = useState(initial?.operator || "eq");
  const [value, setValue] = useState(initial?.value || "");
  const col = columns.find((c) => c.id === colId);

  const operators = col?.type === "counter"
    ? [{ v: "gt", l: "أكبر من" }, { v: "lt", l: "أصغر من" }, { v: "eq", l: "يساوي" }, { v: "empty", l: "فارغ" }, { v: "notEmpty", l: "غير فارغ" }]
    : [{ v: "eq", l: "يساوي" }, { v: "contains", l: "يحتوي على" }, { v: "empty", l: "فارغ" }, { v: "notEmpty", l: "غير فارغ" }];
  const needsValue = !["empty", "notEmpty"].includes(operator);

  return (
    <Modal title="تصفية الجدول" onClose={onClose}>
      <Field label="العمود">
        <select value={colId} onChange={(e) => { setColId(e.target.value); setValue(""); setOperator("eq"); }} style={inputStyle}>
          {columns.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
      </Field>
      <Field label="الشرط">
        <select value={operator} onChange={(e) => setOperator(e.target.value)} style={inputStyle}>
          {operators.map((o) => <option key={o.v} value={o.v}>{o.l}</option>)}
        </select>
      </Field>
      {needsValue && (
        <Field label="القيمة">
          {col?.type === "dropdown" ? (
            <select value={value} onChange={(e) => setValue(e.target.value)} style={inputStyle}>
              <option value="">اختر...</option>
              {(col.options || []).filter((o) => o.label.trim()).map((o) => <option key={o.id} value={o.label}>{o.label}</option>)}
            </select>
          ) : (
            <input type={col?.type === "counter" ? "number" : "text"} value={value} onChange={(e) => setValue(e.target.value)} style={inputStyle} placeholder="القيمة" />
          )}
        </Field>
      )}
      <div className="flex justify-between items-center mt-4">
        <div>
          {initial && <button onClick={() => { onClear(); onClose(); }} className="px-3 py-2 rounded-lg text-sm font-medium" style={{ color: "#9A3B2E" }}>إزالة التصفية</button>}
        </div>
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
          <button
            disabled={!colId || (needsValue && !value.trim())}
            onClick={() => { onApply({ colId, operator, value }); onClose(); }}
            className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
            style={{ background: "#0F6B5C" }}
          >تطبيق</button>
        </div>
      </div>
    </Modal>
  );
}
// Evaluates whether a row matches the active table filter.
function rowMatchesFilter(cls, row, filter) {
  if (!filter) return true;
  const col = cls.columns.find((c) => c.id === filter.colId);
  if (!col) return true;
  const raw = cls.cells[`${row.id}:${filter.colId}`] || lastReportedValue(cls, row.id, filter.colId) || "";
  switch (filter.operator) {
    case "empty": return !raw;
    case "notEmpty": return !!raw;
    case "gt": return Number(raw || 0) > Number(filter.value || 0);
    case "lt": return Number(raw || 0) < Number(filter.value || 0);
    case "contains": return String(raw).toLowerCase().includes(String(filter.value).toLowerCase());
    default: return String(raw).trim() === String(filter.value).trim();
  }
}

function BulkSetPopover({ column, onApply, onClose }) {
  const [val, setVal] = useState("");
  return (
    <div
      className="absolute z-30 top-full mt-1 p-3 rounded-xl shadow-lg"
      style={{ background: "#fff", border: `1px solid ${LINE}`, width: 200 }}
      onClick={(e) => e.stopPropagation()}
    >
      <p className="text-xs font-semibold mb-2" style={{ color: INK }}>رصد نفس القيمة لجميع الطلاب</p>
      {column.type === "dropdown" ? (
        <select value={val} onChange={(e) => setVal(e.target.value)} style={{ ...inputStyle, marginBottom: 8 }}>
          <option value="">اختر...</option>
          {(column.options || []).filter((o) => o.label.trim()).map((o) => <option key={o.id} value={o.label}>{o.label}</option>)}
        </select>
      ) : (
        <input
          type={column.type === "counter" ? "number" : "text"}
          value={val}
          onChange={(e) => setVal(e.target.value)}
          style={{ ...inputStyle, marginBottom: 8 }}
          placeholder={column.type === "counter" ? "مثال: 5" : "القيمة"}
          autoFocus
        />
      )}
      <div className="flex gap-2">
        <button onClick={onClose} className="flex-1 text-xs py-1.5 rounded-lg" style={{ color: MUTED, border: `1px solid ${LINE}` }}>إلغاء</button>
        <button
          onClick={() => { if (val.trim()) { onApply(val); onClose(); } }}
          className="flex-1 text-xs font-bold py-1.5 rounded-lg text-white"
          style={{ background: "#0F6B5C" }}
        >تطبيق</button>
      </div>
    </div>
  );
}

function DropdownCell({ column, value, onChange }) {
  const [open, setOpen] = useState(false);
  const selected = (column.options || []).find((o) => o.label === value);
  return (
    <div className="relative flex justify-center">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-1.5 text-sm px-2 py-1 rounded-md"
        style={{ border: `1px solid ${LINE}`, background: selected ? `${selected.color}20` : "#fff", color: INK, minWidth: "80px", justifyContent: "center" }}
      >
        {selected && <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: selected.color }} />}
        <span>{selected ? selected.label : "—"}</span>
      </button>
      {open && (
        <div className="absolute z-30 top-full mt-1 rounded-lg shadow-lg py-1 min-w-[130px]" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
          <button onClick={() => { onChange(""); setOpen(false); }} className="w-full text-right px-3 py-1.5 text-xs hover:bg-black/5" style={{ color: MUTED }}>بدون</button>
          {(column.options || []).map((o) => (
            <button key={o.id} onClick={() => { onChange(o.label); setOpen(false); }} className="w-full flex items-center gap-2 px-3 py-1.5 text-sm hover:bg-black/5">
              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: o.color }} />
              <span style={{ color: INK }}>{o.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function Cell({ column, value, onChange }) {
  if (column.type === "counter") return <CounterCell value={value} onChange={onChange} />;
  if (column.type === "dropdown") return <DropdownCell column={column} value={value} onChange={onChange} />;
  return <TextCell value={value} onChange={onChange} />;
}

// Counter cell: +/- buttons and direct typing both only update a local
// number, and only log ONE report entry once the value settles (debounced)
// or when the input loses focus — so reaching "4" via four clicks (or typing
// "5" directly) logs a single record, not one per click.
function CounterCell({ value, onChange }) {
  const [n, setN] = useState(Number(value) || 0);
  useEffect(() => { setN(Number(value) || 0); }, [value]);
  const timerRef = useRef(null);
  useEffect(() => () => { if (timerRef.current) clearTimeout(timerRef.current); }, []);

  const scheduleCommit = (v) => {
    if (timerRef.current) clearTimeout(timerRef.current);
    timerRef.current = setTimeout(() => { onChange(String(v)); timerRef.current = null; }, 700);
  };
  const commitNow = (v) => {
    if (timerRef.current) { clearTimeout(timerRef.current); timerRef.current = null; }
    onChange(String(v));
  };
  const bump = (delta) => {
    const next = n + delta;
    setN(next);
    scheduleCommit(next);
  };

  return (
    <div className="flex items-center justify-center gap-1.5">
      <button onClick={() => bump(-1)} className="w-6 h-6 rounded-md flex items-center justify-center hover:bg-black/5" style={{ border: `1px solid ${LINE}` }}><Minus size={12} /></button>
      <input
        type="number"
        value={n}
        onChange={(e) => { const v = Number(e.target.value) || 0; setN(v); scheduleCommit(v); }}
        onBlur={() => commitNow(n)}
        onKeyDown={(e) => { if (e.key === "Enter") { commitNow(n); e.currentTarget.blur(); } }}
        className="w-9 text-center font-semibold text-sm bg-transparent"
        style={{ outline: "none", color: INK }}
      />
      <button onClick={() => bump(1)} className="w-6 h-6 rounded-md flex items-center justify-center hover:bg-black/5" style={{ border: `1px solid ${LINE}` }}><Plus size={12} /></button>
    </div>
  );
}

// Free-text cell: typing only updates local state (so it feels normal to type
// in), and only commits — i.e. only logs ONE report entry with the finished
// sentence — on blur or Enter. This is what fixes notes being logged letter
// by letter as you type.
function TextCell({ value, onChange }) {
  const [local, setLocal] = useState(value || "");
  useEffect(() => { setLocal(value || ""); }, [value]);
  const commit = () => {
    if (local !== (value || "")) onChange(local);
  };
  return (
    <input
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => { if (e.key === "Enter") { commit(); e.currentTarget.blur(); } }}
      className="w-full text-sm text-center rounded-md px-1.5 py-1 bg-transparent"
      style={{ outline: "none" }}
      placeholder="—"
    />
  );
}

// ---------- Report modal ----------

function trashEntryLabel(entry) {
  if (entry.type === "row") return `صف: ${entry.data.row.name}`;
  if (entry.type === "column") return `عمود: ${entry.data.column.name}`;
  if (entry.type === "bulk") return `حذف شامل (${entry.data.columns.length} عمود، ${entry.data.rows.length} صف)`;
  if (entry.type === "bulkRows") return `حذف جماعي (${entry.data.items.length} طالب)`;
  if (entry.type === "class") return `فصل: ${entry.data.cls.subject} — ${entry.data.cls.grade}`;
  if (entry.type === "classesBulk") return `حذف جماعي (${entry.data.classes.length} فصل)`;
  if (entry.type === "reportEntry") return `${entry.data.entry.colName}: ${entry.data.entry.value}`;
  if (entry.type === "reportBulk") return `حذف جماعي من التقرير (${entry.data.entries.length} رصد)`;
  return "عنصر محذوف";
}

function SettingsModal({ feedback, onToggleFeedback, darkMode, onToggleDarkMode, schoolName, principalName, countryName, ministryName, logoImage, onChangeSchoolInfo, footerContacts, footerBadges, onAddContact, onUpdateContact, onRemoveContact, onAddBadge, onRemoveBadge, userEmail, onSignOut, onClose }) {
  const logoInputRef = useRef(null);
  const badgeInputRef = useRef(null);
  const [confirmRemoveLogo, setConfirmRemoveLogo] = useState(false);
  const handleLogoChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChangeSchoolInfo({ logoImage: reader.result });
    reader.readAsDataURL(file);
    e.target.value = "";
  };
  const handleBadgeUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onAddBadge(reader.result);
    reader.readAsDataURL(file);
    e.target.value = "";
  };
  return (
    <Modal title="الإعدادات" onClose={onClose}>
      {userEmail && (
        <div className="flex items-center justify-between p-3 rounded-xl mb-3" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
          <div>
            <p className="text-xs" style={{ color: MUTED }}>الحساب</p>
            <p className="text-sm font-semibold" style={{ color: INK }}>{userEmail}</p>
          </div>
          <button onClick={onSignOut} className="text-xs font-bold px-3 py-1.5 rounded-lg" style={{ color: "#C0392B", border: "1px solid #F0D2CB" }}>تسجيل الخروج</button>
        </div>
      )}
      <div className="flex items-center justify-between p-3 rounded-xl mb-3" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
        <div className="flex items-center gap-2">
          {darkMode ? <Moon size={18} color="#0F6B5C" /> : <Sun size={18} color={MUTED} />}
          <div>
            <p className="text-sm font-semibold" style={{ color: INK }}>الوضع الليلي</p>
            <p className="text-xs" style={{ color: MUTED }}>ألوان داكنة أريح للعين في الإضاءة الخافتة</p>
          </div>
        </div>
        <button
          onClick={onToggleDarkMode}
          className="w-11 h-6 rounded-full shrink-0 relative transition-colors"
          style={{ background: darkMode ? "#0F6B5C" : LINE }}
        >
          <span className="absolute top-0.5 w-5 h-5 rounded-full bg-white shadow transition-all" style={{ [darkMode ? "right" : "left"]: "2px" }} />
        </button>
      </div>
      <div className="flex items-center justify-between p-3 rounded-xl mb-4" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
        <div className="flex items-center gap-2">
          {feedback ? <Volume2 size={18} color="#0F6B5C" /> : <VolumeX size={18} color={MUTED} />}
          <div>
            <p className="text-sm font-semibold" style={{ color: INK }}>تنبيه صوتي واهتزاز عند الرصد</p>
            <p className="text-xs" style={{ color: MUTED }}>تأكيد سريع (صوت + اهتزاز خفيف) كل مرة تسجّل غيابًا أو قيمة</p>
          </div>
        </div>
        <button
          onClick={onToggleFeedback}
          className="w-11 h-6 rounded-full shrink-0 relative transition-colors"
          style={{ background: feedback ? "#0F6B5C" : LINE }}
        >
          <span className="absolute top-0.5 w-5 h-5 rounded-full bg-white shadow transition-all" style={{ [feedback ? "right" : "left"]: "2px" }} />
        </button>
      </div>
      <div className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
        <p className="text-sm font-semibold mb-1" style={{ color: INK }}>بيانات المدرسة</p>
        <p className="text-xs mb-3" style={{ color: MUTED }}>تُستخدم في شهادات التقدير وأي مستندات رسمية أخرى.</p>
        <Field label="اسم الدولة">
          <input style={inputStyle} value={countryName || ""} onChange={(e) => onChangeSchoolInfo({ countryName: e.target.value })} placeholder="مثال: المملكة العربية السعودية" />
        </Field>
        <Field label="اسم الوزارة">
          <input style={inputStyle} value={ministryName || ""} onChange={(e) => onChangeSchoolInfo({ ministryName: e.target.value })} placeholder="مثال: وزارة التعليم" />
        </Field>
        <Field label="اسم المدرسة">
          <input style={inputStyle} value={schoolName || ""} onChange={(e) => onChangeSchoolInfo({ schoolName: e.target.value })} placeholder="مثال: مدرسة الأمل الابتدائية" />
        </Field>
        <Field label="اسم مدير/ة المدرسة">
          <input style={inputStyle} value={principalName || ""} onChange={(e) => onChangeSchoolInfo({ principalName: e.target.value })} placeholder="مثال: أ. سعد القحطاني" />
        </Field>
        <Field label="شعار الوزارة / المدرسة (اختياري)" hint="ارفع الشعار الرسمي الذي تملكه — لا يمكن للتطبيق توليد الشعارات الحكومية تلقائيًا.">
          <input ref={logoInputRef} type="file" accept="image/*" onChange={handleLogoChange} style={{ display: "none" }} />
          <div className="flex items-center gap-2">
            <button type="button" onClick={() => logoInputRef.current?.click()} className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold" style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}>
              <Camera size={15} color="#0F6B5C" /> {logoImage ? "استبدال الشعار" : "رفع شعار"}
            </button>
            {logoImage && (
              <>
                <img src={logoImage} alt="الشعار" className="w-9 h-9 rounded object-contain dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />
                <button type="button" onClick={() => setConfirmRemoveLogo(true)} title="إزالة الشعار" className="p-1.5 rounded hover:bg-black/5">
                  <ImageOff size={15} color={MUTED} />
                </button>
              </>
            )}
          </div>
        </Field>
      </div>

      <div className="p-3 rounded-xl mt-4" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
        <p className="text-sm font-semibold mb-1" style={{ color: INK }}>تذييل الصفحة الرئيسية</p>
        <p className="text-xs mb-3" style={{ color: MUTED }}>بيانات التواصل وشهادات الثقة/الاعتماد التي تظهر أسفل الصفحة الرئيسية — تتحكم فيها إضافةً وحذفًا بالكامل.</p>

        <p className="text-xs font-semibold mb-2" style={{ color: INK }}>بيانات التواصل</p>
        <div className="space-y-2 mb-2">
          {(footerContacts || []).map((c) => (
            <div key={c.id} className="flex items-center gap-2">
              <input style={{ ...inputStyle, flex: 1 }} value={c.label} onChange={(e) => onUpdateContact(c.id, { label: e.target.value })} placeholder="التسمية (مثال: الهاتف)" />
              <input style={{ ...inputStyle, flex: 1 }} value={c.value} onChange={(e) => onUpdateContact(c.id, { value: e.target.value })} placeholder="القيمة" />
              <button onClick={() => onRemoveContact(c.id)} title="حذف" className="p-1.5 rounded hover:bg-black/5 shrink-0"><X size={14} color={MUTED} /></button>
            </div>
          ))}
        </div>
        <button onClick={onAddContact} className="text-xs font-semibold flex items-center gap-1 mb-4" style={{ color: "#0F6B5C" }}>
          <Plus size={13} /> إضافة بيانات تواصل
        </button>

        <p className="text-xs font-semibold mb-2" style={{ color: INK }}>شهادات الثقة / الاعتماد الرسمية</p>
        <div className="flex flex-wrap gap-2 mb-2">
          {(footerBadges || []).map((b) => (
            <div key={b.id} className="relative">
              <img src={b.image} alt="شهادة" className="w-16 h-16 object-contain rounded-lg dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />
              <button onClick={() => onRemoveBadge(b.id)} title="حذف" className="absolute -top-1.5 -left-1.5 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <X size={10} color="#C0392B" />
              </button>
            </div>
          ))}
        </div>
        <input ref={badgeInputRef} type="file" accept="image/*" onChange={handleBadgeUpload} style={{ display: "none" }} />
        <button onClick={() => badgeInputRef.current?.click()} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#0F6B5C" }}>
          <Plus size={13} /> إضافة شهادة/شعار ثقة
        </button>
      </div>
      {confirmRemoveLogo && (
        <ConfirmDialog
          title="إزالة الشعار"
          message="سيُحذف الشعار المرفوع من الإعدادات، وستحتاج لرفعه مجددًا لاستخدامه في الشهادات. متابعة؟"
          confirmLabel="إزالة"
          onCancel={() => setConfirmRemoveLogo(false)}
          onConfirm={() => { onChangeSchoolInfo({ logoImage: null }); setConfirmRemoveLogo(false); }}
        />
      )}
    </Modal>
  );
}

const WHEEL_COLORS = ["#E63958", "#8A93A0", "#1E9E7C", "#3B6FD4", "#8B3FD1", "#F0862E", "#F2B705", "#2EA8C9", "#C23FA0", "#4CAF7D"];

function RandomPickerModal({ rows, onClose }) {
  const [pool, setPool] = useState(() => shuffleArr(rows.map((r) => r.id)));
  const [picked, setPicked] = useState(null);
  const [history, setHistory] = useState([]);
  const [rotation, setRotation] = useState(0);
  const [spinning, setSpinning] = useState(false);

  const segmentAngle = 360 / (rows.length || 1);
  const wheelColors = rows.map((r, i) => WHEEL_COLORS[i % WHEEL_COLORS.length]);
  const conicGradient = `conic-gradient(${wheelColors.map((c, i) => `${c} ${i * segmentAngle}deg ${(i + 1) * segmentAngle}deg`).join(", ")})`;

  const spinToIndex = (index) => {
    const midAngle = index * segmentAngle + segmentAngle / 2;
    const current = rotation % 360;
    const target = (360 - midAngle) % 360;
    const delta = (target - current + 360) % 360;
    setRotation(rotation + delta + 5 * 360);
    setSpinning(true);
    setTimeout(() => setSpinning(false), 3600);
  };

  const pickNext = () => {
    let currentPool = pool;
    if (currentPool.length === 0) currentPool = shuffleArr(rows.map((r) => r.id));
    const id = currentPool[0];
    setPool(currentPool.slice(1));
    const row = rows.find((r) => r.id === id);
    if (!row) return;
    const index = rows.findIndex((r) => r.id === id);
    setPicked(null);
    spinToIndex(index);
    setTimeout(() => {
      setPicked(row);
      setHistory((h) => [row, ...h].slice(0, 20));
    }, 3600);
  };

  useEffect(() => {
    if (rows.length > 0) {
      const t = setTimeout(() => pickNext(), 50);
      return () => clearTimeout(t);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (rows.length === 0) {
    return (
      <Modal title="عجلة الأسماء" onClose={onClose}>
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد طلاب في هذا الفصل بعد.</p>
      </Modal>
    );
  }

  return (
    <Modal title="عجلة الأسماء" onClose={onClose}>
      <p className="text-sm text-center mb-4" style={{ color: MUTED }}>اضغط على العجلة لاختيار طالب</p>
      <div className="relative mx-auto mb-5" style={{ width: 300, height: 300 }}>
        <div style={{ position: "absolute", top: -8, left: "50%", transform: "translateX(-50%)", zIndex: 10, filter: "drop-shadow(0 3px 3px rgba(0,0,0,0.25))" }}>
          <svg width="36" height="44" viewBox="0 0 34 42">
            <path d="M17 0C7.6 0 0 7.6 0 17c0 12.75 17 25 17 25s17-12.25 17-25C34 7.6 26.4 0 17 0z" fill="#232622" />
            <circle cx="17" cy="17" r="7" fill="#fff" />
          </svg>
        </div>
        <div
          className="rounded-full relative overflow-hidden"
          style={{
            width: 300, height: 300,
            background: conicGradient,
            border: "5px solid #fff",
            boxShadow: "0 8px 28px rgba(0,0,0,0.2)",
            transform: `rotate(${rotation}deg)`,
            transition: spinning ? "transform 3.5s cubic-bezier(0.17, 0.67, 0.12, 0.99)" : "none",
          }}
        >
          {rows.map((row, i) => {
            const midAngle = i * segmentAngle + segmentAngle / 2;
            const flip = midAngle > 90 && midAngle < 270;
            return (
              <div key={row.id} style={{ position: "absolute", top: "50%", left: "50%", width: "50%", height: 0, transform: `rotate(${midAngle}deg)`, transformOrigin: "0 0" }}>
                <span style={{
                  position: "absolute", left: 22, top: -8,
                  fontSize: 11, fontWeight: 700, color: "#fff",
                  whiteSpace: "nowrap", maxWidth: 80, overflow: "hidden", textOverflow: "ellipsis",
                  transform: flip ? "rotate(180deg)" : "none", transformOrigin: "left center",
                }}>{row.name}</span>
              </div>
            );
          })}
        </div>
        <button
          onClick={pickNext}
          disabled={spinning}
          className="rounded-full flex items-center justify-center disabled:opacity-90"
          style={{
            position: "absolute", top: "50%", left: "50%", transform: "translate(-50%,-50%)",
            width: 86, height: 86, background: "#181818", color: "#fff",
            border: "4px solid #fff", boxShadow: "0 4px 16px rgba(0,0,0,0.28)",
            fontSize: 14, fontWeight: 800,
          }}
        >{spinning ? "..." : "دوران"}</button>
      </div>

      {picked && !spinning && (
        <div className="rounded-2xl p-4 mb-4 text-center" style={{ background: `${wheelColors[rows.findIndex((r) => r.id === picked.id)]}18`, border: `2px solid ${wheelColors[rows.findIndex((r) => r.id === picked.id)]}` }}>
          <p className="text-xl font-extrabold" style={{ color: INK }}>{picked.name}</p>
        </div>
      )}

      <p className="text-xs mb-2 text-center" style={{ color: MUTED }}>لن يتكرر نفس الطالب حتى يُختار الجميع مرة واحدة. تبقّى {pool.length} من {rows.length} في الدورة الحالية.</p>
      {history.length > 1 && (
        <div className="pt-2" style={{ borderTop: `1px solid ${LINE}` }}>
          <p className="text-xs font-semibold mb-1" style={{ color: MUTED }}>آخر من تم اختيارهم:</p>
          <p className="text-xs" style={{ color: INK }}>{history.slice(1, 6).map((r) => r.name).join("، ")}</p>
        </div>
      )}
    </Modal>
  );
}

function RandomGroupsModal({ rows, onClose }) {
  const [mode, setMode] = useState("count"); // 'count' = number of groups, 'size' = students per group
  const [num, setNum] = useState(4);
  const [groups, setGroups] = useState(null);

  const generate = () => {
    const shuffled = shuffleArr(rows);
    const groupCount = mode === "count" ? Math.max(1, Math.min(num, rows.length)) : Math.max(1, Math.ceil(rows.length / Math.max(1, num)));
    const result = Array.from({ length: groupCount }, () => []);
    shuffled.forEach((row, i) => result[i % groupCount].push(row));
    setGroups(result);
  };

  return (
    <Modal title="مجموعات عشوائية" onClose={onClose} wide>
      {rows.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد طلاب في هذا الفصل بعد.</p>
      ) : (
        <>
          <div className="flex flex-wrap items-end gap-3 mb-4">
            <div className="flex gap-2">
              <button onClick={() => setMode("count")} className="px-3 py-1.5 rounded-full text-xs font-semibold" style={{ background: mode === "count" ? INK : "transparent", color: mode === "count" ? "#fff" : MUTED, border: `1px solid ${mode === "count" ? INK : LINE}` }}>عدد المجموعات</button>
              <button onClick={() => setMode("size")} className="px-3 py-1.5 rounded-full text-xs font-semibold" style={{ background: mode === "size" ? INK : "transparent", color: mode === "size" ? "#fff" : MUTED, border: `1px solid ${mode === "size" ? INK : LINE}` }}>عدد الطلاب لكل مجموعة</button>
            </div>
            <input type="number" min={1} value={num} onChange={(e) => setNum(Math.max(1, Number(e.target.value) || 1))} style={{ ...inputStyle, width: 80 }} />
            <button onClick={generate} className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>
              <Shuffle size={15} /> توليد
            </button>
          </div>
          {groups && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {groups.map((g, gi) => (
                <div key={gi} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
                  <div className="px-3 py-2 font-bold text-sm text-white" style={{ background: COLORS[gi % COLORS.length].hex }}>المجموعة {gi + 1} ({g.length})</div>
                  <div className="p-2">
                    {g.map((row) => (
                      <p key={row.id} className="text-sm py-1 px-1" style={{ color: INK }}>{row.name}</p>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}
    </Modal>
  );
}

function BulkCertificateModal({ cls, rows, schoolName, principalName, countryName, ministryName, logoImage, onClose }) {
  const [title, setTitle] = useState("شهادة تقدير");
  const [reason, setReason] = useState(`تقديرًا لتميّزه وتفوّقه في مادة ${cls.subject}`);
  const [generating, setGenerating] = useState(false);
  const [progress, setProgress] = useState(0);

  const generateAndDownloadZip = async () => {
    setGenerating(true);
    setProgress(0);
    let logoEl = null;
    if (logoImage) {
      try { logoEl = await loadImage(logoImage); } catch (e) { logoEl = null; }
    }
    const files = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const built = buildCertificateCanvas({
        countryName, ministryName, schoolName, logoImageElement: logoEl,
        title, studentName: row.name, reason,
        className: `${cls.subject} — ${cls.grade}`, teacherName: cls.teacher, principalName,
        date: formatDateDisplay(todayKey()), accentColor: row.color,
      });
      const blob = await new Promise((resolve) => built.canvas.toBlob(resolve, "image/png"));
      const buf = new Uint8Array(await blob.arrayBuffer());
      files.push({ name: `${row.name}.png`, data: buf });
      setProgress(Math.round(((i + 1) / rows.length) * 100));
    }
    const zipBlob = buildZip(files);
    downloadBlob(zipBlob, `شهادات-${cls.subject}.zip`);
    setGenerating(false);
    onClose();
  };

  return (
    <Modal title="شهادات تقدير جماعية" onClose={onClose}>
      <p className="text-sm mb-3" style={{ color: MUTED }}>
        سيتم توليد شهادة مستقلة لكل طالب من <b>{rows.length}</b> طالب محدد، بنفس النوع والسبب أدناه (يظهر اسم كل طالب تلقائيًا في شهادته)، ثم تنزيلها جميعًا كملف مضغوط واحد.
      </p>
      <Field label="نوع الشهادة">
        <select value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle}>
          <option>شهادة تقدير</option>
          <option>شهادة شكر</option>
          <option>شهادة تفوّق</option>
          <option>شهادة تميّز</option>
        </select>
      </Field>
      <Field label="سبب التكريم (نفس النص لجميع الطلاب المحددين)">
        <textarea style={{ ...inputStyle, minHeight: 80 }} value={reason} onChange={(e) => setReason(e.target.value)} />
      </Field>
      {generating ? (
        <div className="p-3 rounded-xl flex items-center gap-2" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
          <div className="w-4 h-4 rounded-full border-2 shrink-0" style={{ borderColor: "#0F6B5C transparent #0F6B5C #0F6B5C", animation: "spin 0.8s linear infinite" }} />
          <p className="text-xs font-semibold" style={{ color: INK }}>جارٍ توليد الشهادات... {progress}%</p>
        </div>
      ) : (
        <button onClick={generateAndDownloadZip} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>
          <ImageDown size={16} /> توليد وتنزيل الكل ({rows.length} شهادة)
        </button>
      )}
    </Modal>
  );
}

const REMEDIAL_TEMPLATES = [
  {
    id: "reading_writing",
    name: "القراءة والكتابة (لغتي)",
    weaknessSigns: "صعوبة في القراءة الجهرية بطلاقة، أخطاء متكررة في تركيب الكلمات والجمل، ضعف في الإملاء ورسم الحروف، وبطء واضح في نسخ النصوص.",
    causes: "ضعف الأساس في مرحلة تعلم الحروف والمقاطع، قلة القراءة والممارسة المنزلية، وضعف في التمييز السمعي أو البصري بين الحروف المتشابهة.",
    goals: "أن يقرأ الطالب نصًا قصيرًا بطلاقة مناسبة لمستواه خلال فترة المتابعة، وأن يقل عدد الأخطاء الإملائية في الكتابة الحرة بشكل ملحوظ.",
    methods: "تدريب يومي قصير على القراءة الجهرية مع التصحيح الفوري، تحليل الكلمات إلى مقاطعها الصوتية، ألعاب لتمييز الحروف المتشابهة، تكليفات إملاء تدريجية الصعوبة، وربط القراءة بقصص قصيرة تناسب اهتمام الطالب.",
  },
  {
    id: "math",
    name: "الرياضيات",
    weaknessSigns: "صعوبة في إتقان العمليات الحسابية الأساسية، بطء في الحل، أخطاء متكررة في فهم المسائل اللفظية، وضعف في حفظ واستدعاء الحقائق العددية.",
    causes: "ضعف الأساس في المراحل السابقة، اعتماد على الحفظ دون الفهم، وقلق أو توتر تجاه المادة يؤثر على الأداء.",
    goals: "أن يتقن الطالب العملية الحسابية المستهدفة بدقة مقبولة، وأن يحل مسائل لفظية بسيطة بخطوات صحيحة ومرتبة.",
    methods: "تقسيم المهارة إلى خطوات صغيرة متدرجة، استخدام وسائل محسوسة وأمثلة عملية من الحياة اليومية، تدريبات قصيرة ومتكررة بدل الجرعة الكبيرة الواحدة، ألعاب تعليمية لتقوية سرعة الاستدعاء الذهني، ومراجعة فورية لكل خطأ مع تصحيحه.",
  },
  {
    id: "science",
    name: "العلوم",
    weaknessSigns: "صعوبة في فهم المفاهيم العلمية المجردة، ضعف في ربط المفاهيم بتطبيقاتها، وتراجع في حفظ المصطلحات العلمية.",
    causes: "الاعتماد على الحفظ دون الفهم العميق، قلة الأنشطة العملية والتجريبية، وضعف في المهارات القرائية يؤثر على استيعاب النصوص العلمية.",
    goals: "أن يفسّر الطالب المفهوم العلمي المستهدف بأسلوبه الخاص، وأن يربطه بمثال أو تطبيق من واقعه.",
    methods: "استخدام تجارب وأنشطة عملية مبسّطة، خرائط مفاهيم لتوضيح العلاقات بين الأفكار، ربط الدرس بمواقف حياتية مألوفة للطالب، وتلخيص كل درس بجمل قصيرة بلغة الطالب نفسه.",
  },
  {
    id: "english",
    name: "اللغة الإنجليزية",
    weaknessSigns: "ضعف في حصيلة المفردات، صعوبة في نطق وقراءة الكلمات، وتراجع في فهم القواعد الأساسية.",
    causes: "قلة التعرض للغة خارج الحصة الدراسية، ضعف الأساس في المراحل الأولى، وعدم الممارسة المستمرة.",
    goals: "أن يكتسب الطالب عددًا محددًا من المفردات الجديدة أسبوعيًا، وأن يقرأ جملًا بسيطة بنطق صحيح.",
    methods: "تكرار قصير ويومي للمفردات المستهدفة، بطاقات تعليمية مصوّرة، تمارين استماع ونطق قصيرة، وربط الكلمات بمواقف عملية يستخدمها الطالب فعليًا.",
  },
  {
    id: "general",
    name: "تحصيل دراسي عام",
    weaknessSigns: "تراجع عام في الأداء عبر أكثر من مادة، ضعف التركيز أثناء الحصة، وعدم إكمال الواجبات المنزلية بانتظام.",
    causes: "ظروف أسرية أو نفسية مؤثرة، ضعف في المهارات الأساسية للتعلم، وعدم انتظام في المتابعة الدراسية اليومية.",
    goals: "تحسين مستوى الانتباه والمشاركة داخل الحصة، وانتظام الطالب في تسليم الواجبات المطلوبة.",
    methods: "متابعة يومية قصيرة لتحفيز الطالب، تقسيم المهام لأجزاء صغيرة يسهل إنجازها، تعزيز إيجابي فوري لأي تحسن ولو بسيط، والتواصل المستمر مع ولي الأمر لدعم المتابعة المنزلية.",
  },
];

function RemedialPlanModal({ cls, row, schoolName, principalName, countryName, ministryName, logoImage, onClose }) {
  const [docType, setDocType] = useState("plan"); // 'plan' | 'report'
  const [showName, setShowName] = useState(true);
  const [templateId, setTemplateId] = useState("");
  const [subject, setSubject] = useState(cls.subject || "");
  const [weaknessSigns, setWeaknessSigns] = useState("");
  const [causes, setCauses] = useState("");
  const [goals, setGoals] = useState("");
  const [methods, setMethods] = useState("");
  const [followUp, setFollowUp] = useState("");
  const [notes, setNotes] = useState("");
  const [previewUrl, setPreviewUrl] = useState("");
  const canvasDataRef = useRef(null);

  const applyTemplate = (id) => {
    setTemplateId(id);
    const t = REMEDIAL_TEMPLATES.find((x) => x.id === id);
    if (!t) return;
    setWeaknessSigns(t.weaknessSigns);
    setCauses(t.causes);
    setGoals(t.goals);
    setMethods(t.methods);
  };

  const title = docType === "plan" ? "خطة علاجية" : "تقرير حالة ضعف دراسي وخطة علاجية";

  useEffect(() => {
    let cancelled = false;
    (async () => {
      let logoEl = null;
      if (logoImage) { try { logoEl = await loadImage(logoImage); } catch (e) { logoEl = null; } }
      if (cancelled) return;
      const built = buildRemedialPlanCanvas({
        countryName, ministryName, schoolName, logoImageElement: logoEl,
        title, studentName: row.name, showName, className: cls.grade, teacherName: cls.teacher, principalName,
        date: formatDateDisplay(todayKey()), subject, weaknessSigns, causes, goals, methods, followUp, notes,
      });
      canvasDataRef.current = built;
      setPreviewUrl(built.canvas.toDataURL("image/png"));
    })();
    return () => { cancelled = true; };
  }, [docType, showName, subject, weaknessSigns, causes, goals, methods, followUp, notes, schoolName, principalName, countryName, ministryName, logoImage, cls, row, title]);

  const downloadPng = () => canvasDataRef.current.canvas.toBlob((blob) => downloadBlob(blob, `${title}-${row.name}.png`));
  const sharePdf = async () => {
    const { canvas, logicalWidth, logicalHeight } = canvasDataRef.current;
    const dataUrl = canvas.toDataURL("image/jpeg", 0.97);
    const blob = buildPdfFromJpegDataUrl(dataUrl, canvas.width, canvas.height, logicalWidth, logicalHeight);
    await shareOrDownloadFile(blob, `${title}-${row.name}.pdf`, "application/pdf");
  };

  return (
    <Modal title="خطة علاجية" onClose={onClose} xl>
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div>
          <div className="flex gap-2 mb-3">
            <button onClick={() => setDocType("plan")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: docType === "plan" ? INK : "transparent", color: docType === "plan" ? "#fff" : MUTED, border: `1px solid ${docType === "plan" ? INK : LINE}` }}>خطة علاجية للطالب</button>
            <button onClick={() => setDocType("report")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: docType === "report" ? INK : "transparent", color: docType === "report" ? "#fff" : MUTED, border: `1px solid ${docType === "report" ? INK : LINE}` }}>تقرير رسمي للإدارة</button>
          </div>
          <label className="flex items-center gap-2 text-sm font-medium mb-3" style={{ color: INK }}>
            <input type="checkbox" checked={showName} onChange={(e) => setShowName(e.target.checked)} />
            إظهار اسم الطالب في المستند
          </label>
          <Field label="اختر خطة جاهزة (اختياري)" hint="تعبّئ الحقول أدناه تلقائيًا بمحتوى مقترح — عدّله بما يناسب حالة الطالب.">
            <div className="flex flex-wrap gap-1.5">
              {REMEDIAL_TEMPLATES.map((t) => (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => applyTemplate(t.id)}
                  className="px-2.5 py-1.5 rounded-full text-xs font-semibold"
                  style={{ background: templateId === t.id ? "#0F6B5C" : "transparent", color: templateId === t.id ? "#fff" : INK, border: `1px solid ${templateId === t.id ? "#0F6B5C" : LINE}` }}
                >{t.name}</button>
              ))}
            </div>
          </Field>
          <Field label="المادة"><input style={inputStyle} value={subject} onChange={(e) => setSubject(e.target.value)} /></Field>
          <Field label="مظاهر الضعف"><textarea style={{ ...inputStyle, minHeight: 60 }} value={weaknessSigns} onChange={(e) => setWeaknessSigns(e.target.value)} placeholder="المهارات التي لم يتقنها الطالب بعد..." /></Field>
          <Field label="أسباب الضعف"><textarea style={{ ...inputStyle, minHeight: 60 }} value={causes} onChange={(e) => setCauses(e.target.value)} placeholder="أسرية، صحية، نفسية، ضعف أساسي سابق..." /></Field>
          <Field label="الأهداف العلاجية"><textarea style={{ ...inputStyle, minHeight: 60 }} value={goals} onChange={(e) => setGoals(e.target.value)} placeholder="أهداف محددة وقابلة للقياس..." /></Field>
          <Field label="الأساليب والأنشطة العلاجية" hint="تُعبَّأ تلقائيًا عند اختيار خطة جاهزة أعلاه — عدّلها بما يناسب حالة الطالب">
            <textarea style={{ ...inputStyle, minHeight: 90 }} value={methods} onChange={(e) => setMethods(e.target.value)} placeholder="اختر خطة جاهزة أعلاه أو اكتب الأساليب المناسبة..." />
          </Field>
          <Field label="المتابعة والتقويم"><textarea style={{ ...inputStyle, minHeight: 60 }} value={followUp} onChange={(e) => setFollowUp(e.target.value)} placeholder="كيف ومتى سيُقاس التقدم..." /></Field>
          <Field label="ملاحظات وتوصيات (اختياري)"><textarea style={{ ...inputStyle, minHeight: 50 }} value={notes} onChange={(e) => setNotes(e.target.value)} /></Field>
          {(!countryName && !schoolName) && <p className="text-xs mb-2" style={{ color: MUTED }}>لم تُدخل بيانات المدرسة بعد — أضفها من ⚙️ الإعدادات ليظهر الترويسة الرسمية.</p>}
          <div className="flex gap-2 flex-wrap mt-2">
            <IconBtn icon={ImageDown} label="تنزيل كصورة" onClick={downloadPng} />
            <IconBtn icon={Share2} label="مشاركة/حفظ PDF" onClick={sharePdf} />
          </div>
        </div>
        <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}`, maxHeight: "75vh", overflowY: "auto" }}>
          {previewUrl && <img src={previewUrl} alt="معاينة" className="w-full dark-mode-img-fix" />}
        </div>
      </div>
    </Modal>
  );
}

function CertificateModal({ cls, row, schoolName, principalName, countryName, ministryName, logoImage, onClose }) {
  const [title, setTitle] = useState("شهادة تقدير");
  const [reason, setReason] = useState(`تقديرًا لتميّزه وتفوّقه في مادة ${cls.subject}`);
  const [previewUrl, setPreviewUrl] = useState("");
  const canvasDataRef = useRef(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      let logoEl = null;
      if (logoImage) {
        try { logoEl = await loadImage(logoImage); } catch (e) { logoEl = null; }
      }
      if (cancelled) return;
      const built = buildCertificateCanvas({
        countryName,
        ministryName,
        schoolName,
        logoImageElement: logoEl,
        title,
        studentName: row.name,
        reason,
        className: `${cls.subject} — ${cls.grade}`,
        teacherName: cls.teacher,
        principalName,
        date: formatDateDisplay(todayKey()),
        accentColor: row.color,
      });
      canvasDataRef.current = built;
      setPreviewUrl(built.canvas.toDataURL("image/png"));
    })();
    return () => { cancelled = true; };
  }, [title, reason, schoolName, principalName, countryName, ministryName, logoImage, cls, row]);

  const downloadPng = () => {
    canvasDataRef.current.canvas.toBlob((blob) => downloadBlob(blob, `شهادة-${row.name}.png`));
  };
  const sharePdf = async () => {
    const { canvas, logicalWidth, logicalHeight } = canvasDataRef.current;
    const dataUrl = canvas.toDataURL("image/jpeg", 0.97);
    const blob = buildPdfFromJpegDataUrl(dataUrl, canvas.width, canvas.height, logicalWidth, logicalHeight);
    await shareOrDownloadFile(blob, `شهادة-${row.name}.pdf`, "application/pdf");
  };

  return (
    <Modal title="شهادة تقدير" onClose={onClose} wide>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <Field label="نوع الشهادة">
            <select value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle}>
              <option>شهادة تقدير</option>
              <option>شهادة شكر</option>
              <option>شهادة تفوّق</option>
              <option>شهادة تميّز</option>
            </select>
          </Field>
          <Field label="سبب التكريم">
            <textarea style={{ ...inputStyle, minHeight: 80 }} value={reason} onChange={(e) => setReason(e.target.value)} />
          </Field>
          {!countryName && <p className="text-xs mb-1" style={{ color: MUTED }}>لم تُدخل اسم الدولة بعد — أضفه من ⚙️ الإعدادات.</p>}
          {!ministryName && <p className="text-xs mb-1" style={{ color: MUTED }}>لم تُدخل اسم الوزارة بعد — أضفه من ⚙️ الإعدادات.</p>}
          {!schoolName && <p className="text-xs mb-1" style={{ color: MUTED }}>لم تُدخل اسم المدرسة بعد — أضفه من ⚙️ الإعدادات.</p>}
          {!principalName && <p className="text-xs mb-1" style={{ color: MUTED }}>لم تُدخل اسم مدير المدرسة بعد — أضفه من ⚙️ الإعدادات.</p>}
          {!logoImage && <p className="text-xs mb-2" style={{ color: MUTED }}>لم ترفع شعارًا بعد — يمكنك رفعه من ⚙️ الإعدادات.</p>}
          <div className="flex gap-2 flex-wrap mt-2">
            <IconBtn icon={ImageDown} label="تنزيل كصورة" onClick={downloadPng} />
            <IconBtn icon={Share2} label="مشاركة/حفظ PDF" onClick={sharePdf} />
          </div>
        </div>
        <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
          {previewUrl && <img src={previewUrl} alt="معاينة الشهادة" className="w-full dark-mode-img-fix" />}
        </div>
      </div>
    </Modal>
  );
}

function BehaviorDraftModal({ text, onClose }) {
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (e) { /* ignore */ }
  };
  const share = async () => {
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    await shareOrDownloadFile(blob, "مسودة-تقرير-سلوكي.txt", "text/plain");
  };
  return (
    <Modal title="مسودة تقرير سلوكي" onClose={onClose} wide>
      <textarea
        readOnly
        value={text}
        dir="rtl"
        className="w-full text-sm p-3 rounded-xl mb-3"
        style={{ minHeight: 320, border: `1px solid ${LINE}`, background: "#F8F7F2", color: INK, resize: "vertical", fontFamily: "inherit" }}
      />
      <div className="flex gap-2 justify-end">
        <IconBtn icon={ClipboardCopy} label={copied ? "تم النسخ ✓" : "نسخ"} onClick={copy} />
        <IconBtn icon={Share2} label="مشاركة" onClick={share} />
      </div>
    </Modal>
  );
}

function TrashModal({ trash, onRestore, onClose, onClearAll }) {
  const [confirming, setConfirming] = useState(false);
  return (
    <Modal title="سلة المحذوفات" onClose={onClose} wide>
      {(!trash || trash.length === 0) ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا توجد عناصر محذوفة حاليًا.</p>
      ) : (
        <>
          <div className="flex justify-end mb-3">
            <button onClick={() => setConfirming(true)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium" style={{ color: "#9A3B2E", border: "1px solid #F0D2CB", background: "#FBEDEA" }}>
              <Trash2 size={14} /> حذف الكل نهائيًا
            </button>
          </div>
          <div className="space-y-2">
            {trash.map((entry) => (
              <div key={entry.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                <div className="flex-1">
                  <p className="text-sm font-semibold" style={{ color: INK }}>{trashEntryLabel(entry)}</p>
                  <p className="text-xs mt-0.5" style={{ color: MUTED }}>{entry.when?.day ? `${entry.when.day}، ` : ""}{entry.when?.date} — {entry.when?.time}</p>
                </div>
                <button onClick={() => onRestore(entry.id)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-bold text-white shrink-0" style={{ background: "#0F6B5C" }}>
                  <RotateCcw size={14} /> استعادة
                </button>
              </div>
            ))}
          </div>
        </>
      )}
      {confirming && (
        <ConfirmDialog
          title="حذف كل عناصر السلة"
          message="سيتم حذف جميع العناصر الموجودة في سلة المحذوفات نهائيًا، ولن تستطيع استعادتها. متابعة؟"
          confirmLabel="حذف الكل"
          onCancel={() => setConfirming(false)}
          onConfirm={() => { onClearAll(); setConfirming(false); }}
        />
      )}
    </Modal>
  );
}

function AttendanceModal({ cls, updateClass, onClose, onPrint, onShare }) {
  const [dateKey, setDateKey] = useState(todayKey());

  const setAbsent = (rowId) => updateClass(markAbsentUpdater(rowId, dateKey));
  const setPresent = (rowId) => updateClass(clearAbsentUpdater(rowId, dateKey));

  const statuses = cls.rows.map((row) => attendanceStatus(cls, row.id, dateKey));
  const absentCount = statuses.filter((s) => s === "absent").length;
  const presentCount = statuses.length - absentCount;

  return (
    <Modal title="متابعة الحضور" onClose={onClose} wide>
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <DateField value={dateKey} onChange={setDateKey} />
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-xs px-2.5 py-1 rounded-full font-bold" style={{ background: "#E3F0ED", color: "#0F6B5C" }}>حاضر: {presentCount}</span>
          <span className="text-xs px-2.5 py-1 rounded-full font-bold" style={{ background: "#F5DEDB", color: "#C0392B" }}>غائب: {absentCount}</span>
          <IconBtn icon={Printer} label="طباعة" onClick={() => onPrint(dateKey)} />
          <IconBtn icon={Share2} label="مشاركة" onClick={() => onShare(dateKey)} />
        </div>
      </div>
      <p className="text-xs mb-3" style={{ color: MUTED }}>أي طالب لا يُحدَّد له غياب يُعتبر حاضرًا تلقائيًا لهذا اليوم.</p>
      {cls.rows.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>أضف صفوفًا (طلابًا) أولًا من صفحة الفصل.</p>
      ) : (
        <div className="space-y-2">
          {cls.rows.map((row) => {
            const status = attendanceStatus(cls, row.id, dateKey);
            return (
              <div key={row.id} className="flex items-center gap-3 p-2.5 rounded-lg" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: row.color }} />
                <span className="flex-1 text-sm font-medium" style={{ color: INK }}>{row.name}</span>
                <span
                  className="text-xs font-bold px-2.5 py-1 rounded-full"
                  style={{ background: status === "present" ? "#E3F0ED" : "#F5DEDB", color: status === "present" ? "#0F6B5C" : "#C0392B" }}
                >{status === "present" ? "حاضر" : "غائب"}</span>
                {status === "absent" ? (
                  <>
                    <button title="تعديل (تراجع عن الغياب)" onClick={() => setPresent(row.id)} className="p-1.5 rounded-lg hover:bg-black/5">
                      <Pencil size={14} color={MUTED} />
                    </button>
                    <button title="حذف رصد الغياب" onClick={() => setPresent(row.id)} className="p-1.5 rounded-lg hover:bg-black/5">
                      <Trash2 size={14} color="#C0392B" />
                    </button>
                  </>
                ) : (
                  <button onClick={() => setAbsent(row.id)} className="px-3 py-1.5 rounded-lg text-xs font-bold text-white" style={{ background: "#C0392B" }}>تحديد كغائب</button>
                )}
              </div>
            );
          })}
        </div>
      )}
    </Modal>
  );
}

function EventsModal({ events, speed, onChangeSpeed, onClose, onAdd, onUpdate, onDelete, onMove }) {
  const [newText, setNewText] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editText, setEditText] = useState("");

  const submitNew = () => {
    if (!newText.trim()) return;
    onAdd(newText.trim());
    setNewText("");
  };
  const startEdit = (ev) => { setEditingId(ev.id); setEditText(ev.text); };
  const saveEdit = () => {
    if (editText.trim()) onUpdate(editingId, editText.trim());
    setEditingId(null);
  };

  return (
    <Modal title="الأحداث" onClose={onClose} wide>
      <div className="flex gap-2 mb-4">
        <input
          style={inputStyle}
          value={newText}
          onChange={(e) => setNewText(e.target.value)}
          placeholder="اكتب خبرًا أو حدثًا جديدًا..."
          onKeyDown={(e) => e.key === "Enter" && submitNew()}
        />
        <button onClick={submitNew} className="px-4 py-2 rounded-lg text-sm font-bold text-white shrink-0" style={{ background: "#0F6B5C" }}>إضافة</button>
      </div>
      <div className="flex items-center gap-3 mb-4 p-3 rounded-xl" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
        <span className="text-sm font-semibold shrink-0" style={{ color: INK }}>سرعة الشريط</span>
        <span className="text-xs shrink-0" style={{ color: MUTED }}>أبطأ</span>
        <input
          type="range"
          min="6"
          max="28"
          step="1"
          value={30 - speed}
          onChange={(e) => onChangeSpeed(30 - Number(e.target.value))}
          className="flex-1"
        />
        <span className="text-xs shrink-0" style={{ color: MUTED }}>أسرع</span>
      </div>
      {(!events || events.length === 0) ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا توجد أحداث بعد. أضف أول خبر أعلاه.</p>
      ) : (
        <div className="space-y-2">
          {events.map((ev, i) => (
            <div key={ev.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
              <div className="flex flex-col gap-0.5 shrink-0">
                <button title="نقل لأعلى" disabled={i === 0} onClick={() => onMove(ev.id, -1)} className="p-0.5 rounded hover:bg-black/5 disabled:opacity-25"><ChevronUp size={13} color={MUTED} /></button>
                <button title="نقل لأسفل" disabled={i === events.length - 1} onClick={() => onMove(ev.id, 1)} className="p-0.5 rounded hover:bg-black/5 disabled:opacity-25"><ChevronDown size={13} color={MUTED} /></button>
              </div>
              {editingId === ev.id ? (
                <input
                  autoFocus
                  style={{ ...inputStyle, flex: 1 }}
                  value={editText}
                  onChange={(e) => setEditText(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && saveEdit()}
                />
              ) : (
                <span className="flex-1 text-sm font-medium" style={{ color: INK }}>{ev.text}</span>
              )}
              {editingId === ev.id ? (
                <button onClick={saveEdit} title="حفظ" className="p-1.5 rounded-lg hover:bg-black/5"><Check size={15} color="#0F6B5C" /></button>
              ) : (
                <button onClick={() => startEdit(ev)} title="تعديل" className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={14} color={MUTED} /></button>
              )}
              <button onClick={() => onDelete(ev.id)} title="حذف" className="p-1.5 rounded-lg hover:bg-black/5"><Trash2 size={14} color="#C0392B" /></button>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

// Single-pass scrolling news ticker (no duplicated/overlapping copy of the
// text): it enters fully off-screen on one side and exits fully off-screen on
// the other before looping. Speed is configurable from the الأحداث modal.
function EventsTicker({ events, speed }) {
  const [paused, setPaused] = useState(false);
  if (!events || events.length === 0) return null;
  const text = events.map((e) => e.text).join("      •      ");
  return (
    <div
      className="mb-3 rounded-xl overflow-hidden flex items-center"
      style={{ background: INK, height: 40 }}
      onMouseEnter={() => setPaused(true)}
      onMouseLeave={() => setPaused(false)}
      onTouchStart={() => setPaused(true)}
      onTouchEnd={() => setPaused(false)}
      onTouchCancel={() => setPaused(false)}
    >
      <span className="px-3 shrink-0 flex items-center" style={{ color: "#F2C94C" }}><Newspaper size={16} /></span>
      <div className="overflow-hidden flex-1 h-full flex items-center">
        <div className="ticker-track-single" style={{ animationDuration: `${speed || 14}s`, animationPlayState: paused ? "paused" : "running" }}>
          <span className="text-sm font-semibold" style={{ color: "#fff" }}>{text}</span>
        </div>
      </div>
    </div>
  );
}

const WEEK_DAYS = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"];
const WEEK_PERIODS = [1, 2, 3, 4, 5, 6, 7];
const scheduleKey = (period, day) => `${period}_${day}`;

function ScheduleModal({ schedule, onCellChange, image, onSetImage, onClear, onDeleteAll, onClose }) {
  const [editing, setEditing] = useState(() => !schedule || Object.values(schedule).every((v) => !v || !v.trim()));
  const [confirmType, setConfirmType] = useState(null); // 'clear' | 'delete' | 'removeImage'
  const fileInputRef = useRef(null);

  const handlePhotoChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onSetImage(reader.result);
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  return (
    <Modal title="الجدول الدراسي" onClose={onClose} xl>
      <div className="flex items-center justify-end gap-2 mb-4 flex-wrap">
        <input ref={fileInputRef} type="file" accept="image/*" onChange={handlePhotoChange} style={{ display: "none" }} />
        <IconBtn icon={Camera} label={image ? "استبدال الصورة" : "تصوير أو رفع صورة الجدول"} onClick={() => fileInputRef.current?.click()} />
        {image ? (
          <IconBtn icon={ImageOff} label="إزالة الصورة والعودة لجدول قابل للكتابة" tone="danger" onClick={() => setConfirmType("removeImage")} />
        ) : (
          <>
            <IconBtn icon={Pencil} label={editing ? "إنهاء التعديل" : "تعديل"} onClick={() => setEditing((s) => !s)} />
            <IconBtn icon={Eraser} label="تفريغ" tone="danger" onClick={() => setConfirmType("clear")} />
          </>
        )}
        <IconBtn icon={Trash2} label="حذف" tone="danger" onClick={() => setConfirmType("delete")} />
      </div>

      {image ? (
        <div className="rounded-2xl overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
          <img src={image} alt="الجدول الدراسي" className="w-full object-contain dark-mode-img-fix" style={{ background: "#000", maxHeight: 780 }} />
        </div>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-base">
            <thead>
              <tr>
                <th className="p-3 text-center" style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, width: 110 }}>اليوم</th>
                {WEEK_PERIODS.map((p) => (
                  <th key={p} className="p-3 text-center" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>الحصة {p}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {WEEK_DAYS.map((d, i) => (
                <tr key={d} style={{ background: i % 2 ? "#FBFAF6" : "#fff" }}>
                  <td className="p-3 text-center font-bold" style={{ border: `1px solid ${LINE}`, color: MUTED }}>{d}</td>
                  {WEEK_PERIODS.map((p) => {
                    const key = scheduleKey(p, d);
                    const val = schedule?.[key] || "";
                    return (
                      <td key={p} className="p-2 text-center" style={{ border: `1px solid ${LINE}`, minWidth: 140, height: 56 }}>
                        {editing ? (
                          <input
                            value={val}
                            onChange={(e) => onCellChange(key, e.target.value)}
                            className="w-full text-sm text-center bg-transparent outline-none py-1.5"
                            placeholder="—"
                          />
                        ) : (
                          <span className="text-sm" style={{ color: val ? INK : MUTED }}>{val || "—"}</span>
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {confirmType && (
        <ConfirmDialog
          title={confirmType === "clear" ? "تفريغ الجدول" : confirmType === "removeImage" ? "إزالة الصورة" : "حذف الجدول"}
          message={
            confirmType === "clear"
              ? "سيتم مسح كل خلايا الجدول لإعادة كتابته من جديد. متابعة؟"
              : confirmType === "removeImage"
              ? "سيتم حذف صورة الجدول والرجوع لجدول قابل للكتابة يدويًا. متابعة؟"
              : "سيتم حذف الجدول الدراسي بالكامل (الصورة أو الجدول اليدوي) من الصفحة الرئيسية. متابعة؟"
          }
          confirmLabel={confirmType === "clear" ? "تفريغ" : confirmType === "removeImage" ? "إزالة" : "حذف"}
          onCancel={() => setConfirmType(null)}
          onConfirm={() => {
            if (confirmType === "clear") { onClear(); setEditing(true); }
            else if (confirmType === "removeImage") { onSetImage(null); }
            else { onDeleteAll(); onClose(); }
            setConfirmType(null);
          }}
        />
      )}
    </Modal>
  );
}

function SiteFooter({ contacts, badges }) {
  const hasContacts = contacts && contacts.some((c) => c.label.trim() || c.value.trim());
  const hasBadges = badges && badges.length > 0;
  if (!hasContacts && !hasBadges) return null;
  return (
    <div className="mt-10 pt-6 pb-4" style={{ borderTop: `1px solid ${LINE}` }}>
      {hasBadges && (
        <div className="flex flex-wrap gap-4 justify-center items-center mb-4">
          {badges.map((b) => (
            <img key={b.id} src={b.image} alt="شهادة ثقة" className="h-14 object-contain dark-mode-img-fix" />
          ))}
        </div>
      )}
      {hasContacts && (
        <div className="flex flex-wrap gap-x-6 gap-y-1.5 justify-center text-xs" style={{ color: MUTED }}>
          {contacts.filter((c) => c.label.trim() || c.value.trim()).map((c) => (
            <span key={c.id}><b style={{ color: INK }}>{c.label}</b>{c.label && c.value ? ": " : ""}{c.value}</span>
          ))}
        </div>
      )}
    </div>
  );
}

function TodayActivityModal({ classes, onClose }) {
  const t = todayKey();
  const groups = classes.map((cls) => {
    const items = [];
    Object.entries(cls.reports || {}).forEach(([rowId, entries]) => {
      const row = cls.rows.find((r) => r.id === rowId);
      if (!row) return;
      (entries || []).forEach((e) => {
        if (e.dateKey === t) items.push({ ...e, studentName: row.name });
      });
    });
    return { cls, items };
  }).filter((g) => g.items.length > 0);
  const totalCount = groups.reduce((s, g) => s + g.items.length, 0);

  return (
    <Modal title="نشاطي اليوم" onClose={onClose} wide>
      <p className="text-sm mb-4" style={{ color: MUTED }}>{formatDateDisplay(t)} — {totalCount} رصد إجمالي عبر {groups.length} فصل</p>
      {groups.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لم تسجّل أي شيء اليوم بعد.</p>
      ) : (
        <div className="space-y-4">
          {groups.map(({ cls, items }) => (
            <div key={cls.id} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
              <div className="px-3 py-2 flex items-center gap-2" style={{ background: cls.color }}>
                {cls.emoji && <span>{cls.emoji}</span>}
                <span className="font-bold text-sm" style={{ color: "#fff" }}>{cls.subject} — {cls.grade}</span>
                <span className="text-xs mr-auto px-2 py-0.5 rounded-full font-semibold" style={{ background: "rgba(255,255,255,0.25)", color: "#fff" }}>{items.length}</span>
              </div>
              <div>
                {items.map((it, i) => (
                  <div key={it.id} className="flex items-center gap-3 px-3 py-2" style={{ borderTop: i > 0 ? `1px solid ${LINE}` : "none" }}>
                    <span className="w-2 h-2 rounded-full shrink-0" style={{ background: it.colColor }} />
                    <span className="flex-1 text-sm" style={{ color: INK }}><b>{it.studentName}</b> — {it.colName}: {it.value}</span>
                    <span className="text-xs shrink-0" style={{ color: MUTED }}>{it.time}</span>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function GuideModal({ onClose }) {
  const steps = [
    { icon: Plus, title: "أنشئ فصلك الأول", text: "اضغط علامة + بالصفحة الرئيسية، وأدخل اسم المادة والصف والمعلم والعام الدراسي." },
    { icon: Plus, title: "أضف الأعمدة والطلاب", text: "داخل الفصل، من مجموعة \"إدارة الجدول\" أضف أعمدة (واجب، سلوك، اختبار...) ثم أضف الطلاب — يدويًا، دفعة واحدة، أو استيرادًا من Excel." },
    { icon: CalendarCheck, title: "سجّل يوميًا", text: "من مجموعة \"أدوات الحصة\" تقدر تتابع الحضور، تسجّل أحداثًا، أو تختار طالبًا عشوائيًا أثناء الحصة." },
    { icon: FileText, title: "التقرير والشهادات", text: "لكل طالب تقرير مفصّل يجمع كل ما رُصد له، وتقدر تولّد منه شهادة تقدير جاهزة للطباعة." },
    { icon: Printer, title: "اطبع أو شارك", text: "من مجموعة \"العرض والطباعة\" تصدّر أي فصل كـ PDF أو صورة أو Excel، أو تشارك نسخة للقراءة فقط." },
    { icon: RotateCcw, title: "لا تقلق من الخطأ", text: "أي حذف (صف، عمود، فصل) يُحفظ في سلة المحذوفات، وزر \"تراجع\" يرجّع آخر شيء حذفته بضغطة واحدة." },
  ];
  return (
    <Modal title="كيف أبدأ؟" onClose={onClose} wide>
      <div className="space-y-3">
        {steps.map((s, i) => (
          <div key={i} className="flex items-start gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: "#EAF3F0" }}>
              <s.icon size={16} color="#0F6B5C" />
            </div>
            <div>
              <p className="text-sm font-bold mb-0.5" style={{ color: INK }}>{s.title}</p>
              <p className="text-xs" style={{ color: MUTED }}>{s.text}</p>
            </div>
          </div>
        ))}
      </div>
      <button onClick={onClose} className="w-full mt-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{ background: "#0F6B5C" }}>فهمت، ابدأ الآن</button>
    </Modal>
  );
}

function DashboardStrip({ classes }) {
  if (!classes || classes.length === 0) return null;
  const t = todayKey();
  const weekAgo = new Date();
  weekAgo.setDate(weekAgo.getDate() - 7);

  let totalStudents = 0, totalPresent = 0, notesThisWeek = 0;
  classes.forEach((cls) => {
    (cls.rows || []).forEach((row) => {
      totalStudents += 1;
      if (attendanceStatus(cls, row.id, t) === "present") totalPresent += 1;
    });
    Object.values(cls.reports || {}).forEach((list) => {
      (list || []).forEach((e) => {
        if (!e.dateKey) return;
        const d = new Date(`${e.dateKey}T00:00:00`);
        if (d >= weekAgo) notesThisWeek += 1;
      });
    });
  });
  const attendanceRate = totalStudents > 0 ? Math.round((totalPresent / totalStudents) * 100) : null;

  const cards = [
    { label: "الفصول النشطة", value: classes.length, icon: BookOpen, color: "#0F6B5C" },
    { label: "نسبة الحضور اليوم", value: attendanceRate === null ? "—" : `${attendanceRate}%`, icon: CalendarCheck, color: attendanceRate === null || attendanceRate >= 90 ? "#0F6B5C" : attendanceRate >= 75 ? "#C97A2B" : "#C0392B" },
    { label: "ملاحظات هذا الأسبوع", value: notesThisWeek, icon: FileText, color: "#3B4C8C" },
  ];

  return (
    <div className="grid grid-cols-3 gap-3 mb-5">
      {cards.map((c) => (
        <div key={c.label} className="rounded-2xl p-3 flex items-center gap-2.5" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
          <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: `${c.color}1A` }}>
            <c.icon size={16} color={c.color} />
          </div>
          <div className="min-w-0">
            <p className="text-lg font-extrabold leading-tight" style={{ color: INK }}>{c.value}</p>
            <p className="text-xs truncate" style={{ color: MUTED }}>{c.label}</p>
          </div>
        </div>
      ))}
    </div>
  );
}

function ScheduleMiniCard({ schedule, image, onOpen }) {
  const hasData = image || (schedule && Object.values(schedule).some((v) => v && v.trim()));
  return (
    <button
      onClick={onOpen}
      className="text-right rounded-2xl p-4 mb-5 hover:opacity-90 transition-opacity block"
      style={{ background: "#fff", border: `1px solid ${LINE}`, width: "100%" }}
    >
      <div className="flex items-center gap-2 mb-3">
        <CalendarRange size={18} color="#0F6B5C" />
        <span className="font-bold text-base" style={{ color: INK }}>الجدول الدراسي</span>
      </div>
      {!hasData ? (
        <p className="text-sm" style={{ color: MUTED }}>اضغط لإنشاء جدولك الدراسي الأسبوعي أو رفع صورته</p>
      ) : image ? (
        <img src={image} alt="الجدول الدراسي" className="w-full rounded-lg dark-mode-img-fix" style={{ maxHeight: 420, objectFit: "contain", background: "#F3F1E9" }} />
      ) : (
        <table className="w-full border-collapse" style={{ fontSize: 14 }}>
          <thead>
            <tr>
              <th style={{ padding: 8, border: `1px solid ${LINE}`, background: "#F3F1E9" }}></th>
              {WEEK_PERIODS.map((p) => (
                <th key={p} style={{ padding: 8, border: `1px solid ${LINE}`, background: "#F3F1E9", color: MUTED }}>{p}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {WEEK_DAYS.map((d) => (
              <tr key={d}>
                <td style={{ padding: 8, border: `1px solid ${LINE}`, textAlign: "center", color: MUTED, fontWeight: 700 }}>{d.slice(2, 4)}</td>
                {WEEK_PERIODS.map((p) => {
                  const val = schedule?.[scheduleKey(p, d)] || "";
                  return (
                    <td key={p} style={{ padding: 8, border: `1px solid ${LINE}`, textAlign: "center", color: INK, maxWidth: 90, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {val}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </button>
  );
}

function ReportModal({ cls, row, entries, reportTrash, schoolName, principalName, countryName, ministryName, logoImage, onClose, onEditEntry, onDeleteEntry, onDeleteCategory, onDeleteAllEntries, onAddNote, onRestoreLatest, onRestoreEntry, onClearTrash, onPrint }) {
  const [editing, setEditing] = useState(false);
  const [noteText, setNoteText] = useState("");
  const [showTrash, setShowTrash] = useState(false);
  const [behaviorDraftGroup, setBehaviorDraftGroup] = useState(null);
  const [showMedicalBanner, setShowMedicalBanner] = useState(true);
  const [showCertificate, setShowCertificate] = useState(false);
  const [showBehaviorBanner, setShowBehaviorBanner] = useState(true);
  const [sharingStudent, setSharingStudent] = useState(false);
  const [confirmDeleteAll, setConfirmDeleteAll] = useState(false);
  const [confirmDeleteCategory, setConfirmDeleteCategory] = useState(null);
  const [showRemedialPlan, setShowRemedialPlan] = useState(false);

  const addNote = () => {
    if (!noteText.trim()) return;
    const meta = nowMeta();
    onAddNote({ id: uid(), colId: "note", colName: "ملاحظة", colColor: COLORS[7].hex, value: noteText.trim(), ...meta, manual: true });
    setNoteText("");
  };

  const shareStudentReadOnly = async () => {
    setSharingStudent(true);
    try {
      const html = buildReadOnlyStudentHtml(cls, row, entries);
      const blob = new Blob([html], { type: "text/html" });
      await shareOrDownloadFile(blob, `تقرير-${row.name}.html`, "text/html");
    } finally {
      setSharingStudent(false);
    }
  };

  const groups = groupEntries(entries);
  const flaggedGroups = groups.filter((g) => {
    const col = cls.columns.find((c) => c.id === g.colId);
    return col?.behaviorFlag && g.items.length >= (col.behaviorThreshold || 3);
  });

  return (
    <Modal title={`تقرير الطالب — ${row.name}`} onClose={onClose} lg>
      <div className="rounded-2xl p-4 mb-4 flex flex-wrap items-center gap-4" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
        <div className="w-12 h-12 rounded-full flex items-center justify-center font-bold text-lg shrink-0" style={{ background: row.color, color: "#fff" }}>
          {(row.name || "؟").trim().charAt(0)}
        </div>
        <div className="flex-1" style={{ minWidth: 160 }}>
          <p className="font-bold text-base" style={{ color: INK }}>{row.name}</p>
          <p className="text-xs" style={{ color: MUTED }}>{cls.subject} • {cls.grade} • {cls.teacher}</p>
        </div>
      </div>

      {(groups.length > 0 || (row.medicalNote && row.medicalNote.trim())) && (
        <div className="flex gap-3 overflow-x-auto mb-4 pb-1">
          {row.medicalNote && row.medicalNote.trim() && (
            <div className="rounded-2xl p-3 text-center shrink-0 relative" style={{ background: "#FBEDEA", border: "2px solid #C0392B", minWidth: 92, maxWidth: 170 }}>
              <button
                onClick={() => setShowMedicalBanner((s) => !s)}
                title={showMedicalBanner ? "إخفاء التفاصيل" : "إظهار التفاصيل"}
                className="absolute top-1.5 left-1.5 p-1 rounded-full hover:bg-black/10"
              >
                {showMedicalBanner ? <Eye size={12} color="#9A3B2E" /> : <EyeOff size={12} color="#9A3B2E" />}
              </button>
              <div className="w-9 h-9 rounded-full mx-auto mb-2 flex items-center justify-center" style={{ background: "#C0392B" }}>
                <AlertTriangle size={16} color="#fff" strokeWidth={2.5} />
              </div>
              <p className="text-xs font-extrabold mb-1" style={{ color: "#9A3B2E" }}>تنبيه خاص</p>
              {showMedicalBanner && (
                <p className="text-[11px] leading-snug" style={{ color: "#7A2E22", wordBreak: "break-word", whiteSpace: "pre-wrap" }}>{row.medicalNote}</p>
              )}
            </div>
          )}
          {groups.map((g) => (
            <div key={g.colId} className="rounded-2xl p-3 text-center shrink-0" style={{ background: `${g.colColor}14`, minWidth: 92 }}>
              <div className="w-9 h-9 rounded-full mx-auto mb-2 flex items-center justify-center" style={{ background: g.colColor }}>
                <Check size={16} color="#fff" strokeWidth={3} />
              </div>
              <p className="text-2xl font-extrabold" style={{ color: INK }}>{g.items.length}</p>
              <p className="text-xs" style={{ color: MUTED }}>{g.colName}</p>
            </div>
          ))}
        </div>
      )}

      {flaggedGroups.map((g) => (
        showBehaviorBanner ? (
          <div key={g.colId} className="flex items-center justify-between gap-3 p-3 rounded-xl mb-4 flex-wrap" style={{ background: "#FBEDEA", border: "1px solid #F0D2CB" }}>
            <div className="flex items-start gap-2">
              <MessageSquareWarning size={16} color="#9A3B2E" className="shrink-0 mt-0.5" />
              <p className="text-sm font-semibold" style={{ color: "#7A2E22" }}>
                وصل عدد ملاحظات "{g.colName}" إلى {g.items.length} — يمكنك توليد مسودة تقرير سلوكي جاهزة.
              </p>
            </div>
            <div className="flex items-center gap-2 shrink-0">
              <button onClick={() => setBehaviorDraftGroup(g)} className="px-3 py-1.5 rounded-lg text-xs font-bold text-white" style={{ background: "#9A3B2E" }}>
                توليد مسودة تقرير
              </button>
              <button onClick={() => setShowBehaviorBanner(false)} title="إخفاء التنبيه" className="p-1 rounded hover:bg-black/5">
                <Eye size={15} color="#9A3B2E" />
              </button>
            </div>
          </div>
        ) : null
      ))}
      {!showBehaviorBanner && flaggedGroups.length > 0 && (
        <button onClick={() => setShowBehaviorBanner(true)} className="flex items-center gap-1.5 text-xs font-semibold mb-4 px-2.5 py-1 rounded-full" style={{ color: "#9A3B2E", border: "1px solid #F0D2CB" }}>
          <EyeOff size={13} /> تنبيهات السلوك مخفية — إظهار
        </button>
      )}

      <p className="text-sm font-semibold mb-2" style={{ color: INK }}>{entries.length} رصد إجمالي عبر {groups.length} تصنيف</p>

      <div className="rounded-xl p-2.5 mb-2.5 flex flex-wrap items-center gap-2" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
        <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>التقرير</span>
        <IconBtn icon={Pencil} label={editing ? "إنهاء التعديل" : "تعديل"} onClick={() => setEditing((s) => !s)} />
        <div className="flex items-center rounded-lg overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
          <button onClick={onRestoreLatest} title="استعادة آخر محذوف من هذا التقرير" className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium hover:opacity-80" style={{ color: INK }}>
            <RotateCcw size={15} /> استعادة
          </button>
          <button onClick={() => setShowTrash(true)} title="سجل المحذوفات من هذا التقرير" className="px-2.5 py-2 hover:bg-black/5" style={{ borderInlineStart: `1px solid ${LINE}` }}>
            <FolderOpen size={15} color={MUTED} />
          </button>
        </div>
        <IconBtn icon={Activity} label="خطة علاجية" onClick={() => setShowRemedialPlan(true)} />
      </div>

      <div className="rounded-xl p-2.5 mb-4 flex flex-wrap items-center gap-2" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
        <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>المخرجات</span>
        <IconBtn icon={Printer} label="طباعة" onClick={() => onPrint()} />
        <IconBtn icon={Award} label="شهادة تقدير" onClick={() => setShowCertificate(true)} />
        <IconBtn icon={Share2} label={sharingStudent ? "جارٍ التجهيز..." : "مشاركة تقرير هذا الطالب فقط"} onClick={shareStudentReadOnly} />
        {entries.length > 0 && (
          <IconBtn icon={Trash2} label="حذف كل التقرير" tone="danger" onClick={() => setConfirmDeleteAll(true)} />
        )}
      </div>

      {groups.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد رصد لهذا الطالب بعد.</p>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4 items-start">
          {groups.map((g) => (
            <div key={g.colId} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
              <div className="px-4 py-2.5 flex items-center gap-2" style={{ background: g.colColor }}>
                <span className="font-bold text-sm" style={{ color: "#fff" }}>{g.colName}</span>
                <span className="text-xs px-2 py-0.5 rounded-full font-semibold" style={{ background: "rgba(255,255,255,0.25)", color: "#fff" }}>{g.items.length} رصد</span>
                <button onClick={() => setConfirmDeleteCategory(g)} title={`حذف كل رصد "${g.colName}"`} className="mr-auto p-1 rounded hover:bg-black/10">
                  <Trash2 size={14} color="#fff" />
                </button>
              </div>
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr>
                    <th className="p-2 text-right text-xs font-semibold" style={{ background: "#F8F7F2", color: MUTED, width: "40%" }}>اليوم والتاريخ</th>
                    <th className="p-2 text-center text-xs font-semibold" style={{ background: "#F8F7F2", color: MUTED, width: "18%" }}>الوقت</th>
                    <th className="p-2 text-right text-xs font-semibold" style={{ background: "#F8F7F2", color: MUTED }}>القيمة</th>
                    <th className="p-2 text-center text-xs font-semibold" style={{ background: "#F8F7F2", color: MUTED, width: 40 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {g.items.map((e, i) => (
                    <tr key={e.id} style={{ background: i % 2 ? "#FBFAF6" : "#fff" }}>
                      <td className="p-2 text-xs" style={{ borderTop: `1px solid ${LINE}`, color: INK }}>{e.day ? `${e.day}، ` : ""}{e.date}</td>
                      <td className="p-2 text-xs text-center" style={{ borderTop: `1px solid ${LINE}`, color: MUTED }}>{e.time}</td>
                      <td className="p-2" style={{ borderTop: `1px solid ${LINE}`, wordBreak: "break-word", whiteSpace: "pre-wrap" }}>
                        {editing ? (
                          <input style={{ ...inputStyle, padding: "5px 8px" }} value={e.value} onChange={(ev) => onEditEntry(e.id, ev.target.value)} />
                        ) : (
                          <span className="text-sm font-medium" style={{ color: INK }}>{e.value}</span>
                        )}
                      </td>
                      <td className="p-2 text-center" style={{ borderTop: `1px solid ${LINE}` }}>
                        <button title="حذف هذا الرصد" onClick={() => onDeleteEntry(e.id)} className="p-1 rounded hover:bg-black/5">
                          <Trash2 size={13} color="#C0392B" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))}
        </div>
      )}

      {editing && (
        <div className="flex gap-2">
          <input style={inputStyle} value={noteText} onChange={(e) => setNoteText(e.target.value)} placeholder="أضف ملاحظة يدوية للتقرير..." />
          <button onClick={addNote} className="px-4 py-2 rounded-lg text-sm font-bold text-white shrink-0" style={{ background: "#0F6B5C" }}>إضافة</button>
        </div>
      )}

      {showTrash && (
        <TrashModal
          trash={reportTrash}
          onClose={() => setShowTrash(false)}
          onRestore={(id) => onRestoreEntry(id)}
          onClearAll={onClearTrash}
        />
      )}
      {behaviorDraftGroup && (
        <BehaviorDraftModal
          text={buildBehaviorDraft(cls, row, behaviorDraftGroup)}
          onClose={() => setBehaviorDraftGroup(null)}
        />
      )}
      {showCertificate && (
        <CertificateModal
          cls={cls}
          row={row}
          schoolName={schoolName}
          principalName={principalName}
          countryName={countryName}
          ministryName={ministryName}
          logoImage={logoImage}
          onClose={() => setShowCertificate(false)}
        />
      )}
      {showRemedialPlan && (
        <RemedialPlanModal
          cls={cls}
          row={row}
          schoolName={schoolName}
          principalName={principalName}
          countryName={countryName}
          ministryName={ministryName}
          logoImage={logoImage}
          onClose={() => setShowRemedialPlan(false)}
        />
      )}
      {confirmDeleteAll && (
        <ConfirmDialog
          title="حذف كل التقرير"
          message={`سيتم حذف كل رصد الطالب "${row.name}" (${entries.length} عنصر). يمكن استعادتها كلها دفعة واحدة من سجل المحذوفات. متابعة؟`}
          onCancel={() => setConfirmDeleteAll(false)}
          onConfirm={() => { onDeleteAllEntries(); setConfirmDeleteAll(false); }}
        />
      )}
      {confirmDeleteCategory && (
        <ConfirmDialog
          title={`حذف تصنيف "${confirmDeleteCategory.colName}"`}
          message={`سيتم حذف كل رصد "${confirmDeleteCategory.colName}" لهذا الطالب (${confirmDeleteCategory.items.length} عنصر). يمكن استعادتها دفعة واحدة من سجل المحذوفات. متابعة؟`}
          onCancel={() => setConfirmDeleteCategory(null)}
          onConfirm={() => { onDeleteCategory(confirmDeleteCategory.colId); setConfirmDeleteCategory(null); }}
        />
      )}
    </Modal>
  );
}

// ---------- Board table (shared by board view & printing) ----------

function BoardTable({ cls, dateKey }) {
  if (cls.columns.length === 0 || cls.rows.length === 0) {
    return <p className="text-sm py-10 text-center" style={{ color: MUTED }}>لا توجد بيانات لعرضها بعد.</p>;
  }
  const t = dateKey || todayKey();
  return (
    <table className="w-full border-collapse text-sm">
      <thead>
        <tr>
          {cls.showRowNumbers && <th className="p-2 text-center" style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, width: 40 }}>#</th>}
          <th className="p-2 text-right" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>الاسم</th>
          {cls.rows.length > 0 && <th className="p-2 text-center" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>الغياب</th>}
          {cls.columns.map((col) => (
            <th key={col.id} className="p-2 text-center" style={{ background: colorLight(col.color), border: `1px solid ${LINE}`, color: INK }}>{col.name}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {cls.rows.map((row, i) => {
          const status = attendanceStatus(cls, row.id, t);
          return (
            <tr key={row.id} style={{ background: i % 2 ? "#FBFAF6" : "#fff" }}>
              {cls.showRowNumbers && <td className="p-2 text-center text-xs font-semibold" style={{ border: `1px solid ${LINE}`, color: MUTED }}>{i + 1}</td>}
              <td className="p-2 font-medium" style={{ border: `1px solid ${LINE}`, color: INK, borderInlineStart: `4px solid ${row.color}` }}>{row.name}</td>
              <td className="p-1 text-center font-bold" style={{ border: `1px solid ${LINE}`, color: status === "absent" ? "#C0392B" : "#0F6B5C" }}>{status === "absent" ? "غائب" : "حاضر"}</td>
              {cls.columns.map((col) => {
                const val = dateKey ? valueOnDate(cls, row.id, col.id, dateKey) : (cls.cells[`${row.id}:${col.id}`] || lastReportedValue(cls, row.id, col.id));
                return <td key={col.id} className="p-1 text-center" style={{ border: `1px solid ${LINE}` }}>{val || "—"}</td>;
              })}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function DisplayBoard({ cls, onClose, onPrint }) {
  const [sharing, setSharing] = useState(false);
  const [boardDate, setBoardDate] = useState(todayKey());
  const shareReadOnly = async () => {
    setSharing(true);
    try {
      const html = buildReadOnlyBoardHtml(cls, boardDate);
      const blob = new Blob([html], { type: "text/html" });
      await shareOrDownloadFile(blob, `${cls.subject || "الفصل"}-لوحة-العرض.html`, "text/html");
    } finally {
      setSharing(false);
    }
  };
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 modal-backdrop-in" style={{ background: "rgba(35,38,34,0.45)" }}>
      <div className="rounded-2xl shadow-2xl w-full max-w-5xl max-h-[92vh] overflow-hidden flex flex-col modal-panel-in" style={{ background: "#fff" }}>
        <div className="flex items-center justify-between px-5 py-4 flex-wrap gap-2" style={{ borderBottom: `1px solid ${LINE}` }}>
          <h3 className="font-bold text-lg flex items-center gap-2" style={{ color: INK }}>{cls.emoji && <span>{cls.emoji}</span>}لوحة العرض — {cls.subject}</h3>
          <div className="flex items-center flex-wrap gap-2">
            <DateField value={boardDate} onChange={setBoardDate} />
            <IconBtn icon={Printer} label="طباعة" onClick={() => onPrint(boardDate)} />
            <IconBtn icon={Share2} label={sharing ? "جارٍ التجهيز..." : "مشاركة نسخة للقراءة فقط"} onClick={shareReadOnly} />
            <button onClick={onClose} className="p-1.5 rounded-full hover:bg-black/5 active:scale-90 transition-transform"><X size={18} color={MUTED} /></button>
          </div>
        </div>
        <div className="p-5 overflow-auto">
          <div className="mb-3 text-sm" style={{ color: MUTED }}>{cls.grade} • {cls.teacher} • {formatDateDisplay(boardDate)}</div>
          <BoardTable cls={cls} dateKey={boardDate} />
        </div>
      </div>
    </div>
  );
}

// ---------- Printable content ----------

function PrintContent({ job }) {
  if (!job) return null;
  if (job.type === "class") {
    const cls = job.cls;
    return (
      <div>
        <h1 className="text-xl font-bold mb-1">{cls.subject}</h1>
        <p className="text-sm mb-4" style={{ color: "#555" }}>{cls.grade} • {cls.teacher} • {cls.yearHijri} هـ / {cls.yearGregorian} م</p>
        <BoardTable cls={cls} />
      </div>
    );
  }
  if (job.type === "report") {
    const { cls, row, entries } = job;
    const groups = groupEntries(entries);
    return (
      <div>
        <h1 className="text-xl font-bold mb-1">تقرير الطالب: {row.name}</h1>
        <p className="text-sm mb-4" style={{ color: "#555" }}>{cls.subject} • {cls.grade} • {cls.teacher} • {cls.yearHijri} هـ / {cls.yearGregorian} م</p>
        {groups.map((g) => (
          <div key={g.colId} className="mb-4">
            <h2 className="text-base font-bold mb-1" style={{ color: g.colColor }}>{g.colName}</h2>
            <table className="w-full border-collapse text-sm mb-2">
              <thead>
                <tr>
                  <th className="p-2 text-right" style={{ border: "1px solid #ddd", background: "#f3f1e9" }}>اليوم والتاريخ</th>
                  <th className="p-2 text-right" style={{ border: "1px solid #ddd", background: "#f3f1e9" }}>الوقت</th>
                  <th className="p-2 text-right" style={{ border: "1px solid #ddd", background: "#f3f1e9" }}>القيمة</th>
                </tr>
              </thead>
              <tbody>
                {g.items.map((e) => (
                  <tr key={e.id}>
                    <td className="p-2" style={{ border: "1px solid #ddd" }}>{e.day ? `${e.day}، ` : ""}{e.date}</td>
                    <td className="p-2" style={{ border: "1px solid #ddd" }}>{e.time}</td>
                    <td className="p-2" style={{ border: "1px solid #ddd" }}>{e.value}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ))}
      </div>
    );
  }
  return null;
}

// ---------- Home page ----------

function ClassCard({ cls, onOpen, onEdit, onColor, onDelete, onArchive, onDuplicate, onTogglePin, onToggleLock, onMove, onNewTerm, animating, isFirst, isLast }) {
  const [showColors, setShowColors] = useState(false);
  const locked = !!cls.locked;
  return (
    <div
      className={`rounded-2xl shadow-sm overflow-hidden relative card-in transition-shadow hover:shadow-md ${animating ? "trash-toss" : ""}`}
      style={{ background: "#fff", border: cls.pinned ? `2px solid #0F6B5C` : `1px solid ${LINE}` }}
    >
      <div className="h-2" style={{ background: cls.color }} />
      <div className="p-4 cursor-pointer relative active:scale-[0.98] transition-transform" onClick={() => !locked && onOpen(cls.id)} style={{ cursor: locked ? "not-allowed" : "pointer" }}>
        {(cls.pinned || locked) && (
          <div className="absolute top-2 left-2 flex gap-1">
            {cls.pinned && <span className="p-1 rounded-full" style={{ background: "#EAF3F0" }}><Pin size={11} color="#0F6B5C" /></span>}
            {locked && <span className="p-1 rounded-full" style={{ background: "#FBEDEA" }}><Lock size={11} color="#9A3B2E" /></span>}
          </div>
        )}
        <div className="flex items-center gap-2 mb-1">
          {cls.emoji ? <span className="text-base leading-none">{cls.emoji}</span> : <BookOpen size={16} color={cls.color} />}
          <h3 className="font-bold text-base" style={{ color: INK }}>{cls.subject}</h3>
        </div>
        <p className="text-sm mb-0.5" style={{ color: MUTED }}>{cls.grade}</p>
        <p className="text-sm mb-0.5" style={{ color: MUTED }}>{cls.teacher}</p>
        <p className="text-xs" style={{ color: MUTED }}>{cls.yearHijri}{cls.yearHijri && cls.yearGregorian ? " هـ  •  " : ""}{cls.yearGregorian}{cls.yearGregorian ? " م" : ""}</p>
        {locked && <p className="text-xs mt-1 font-semibold" style={{ color: "#9A3B2E" }}>مقفل — اضغط أيقونة القفل للدخول</p>}
      </div>
      <div className="flex items-center gap-1 px-3 py-2 relative flex-wrap" style={{ borderTop: `1px solid ${LINE}` }}>
        <button title={cls.pinned ? "إلغاء التثبيت" : "تثبيت الفصل"} onClick={() => onTogglePin(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5">
          <Pin size={15} color={cls.pinned ? "#0F6B5C" : MUTED} />
        </button>
        <button title={locked ? "فتح القفل" : "قفل الفصل"} onClick={() => onToggleLock(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5">
          {locked ? <Lock size={15} color="#9A3B2E" /> : <Unlock size={15} color={MUTED} />}
        </button>
        {!cls.pinned && (
          <>
            <button title="نقل لأعلى" disabled={isFirst} onClick={() => onMove(cls.id, -1)} className="p-1.5 rounded-lg hover:bg-black/5 disabled:opacity-30"><ChevronUp size={15} color={MUTED} /></button>
            <button title="نقل لأسفل" disabled={isLast} onClick={() => onMove(cls.id, 1)} className="p-1.5 rounded-lg hover:bg-black/5 disabled:opacity-30"><ChevronDown size={15} color={MUTED} /></button>
          </>
        )}
        <button title="تعديل" onClick={() => onEdit(cls)} className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={15} color={MUTED} /></button>
        <button title="الوان" onClick={() => setShowColors((s) => !s)} className="p-1.5 rounded-lg hover:bg-black/5"><Palette size={15} color={MUTED} /></button>
        <button title="تكرار الفصل" onClick={() => onDuplicate(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5"><Copy size={15} color={MUTED} /></button>
        <button title="بدء فصل دراسي جديد" onClick={() => onNewTerm(cls)} className="p-1.5 rounded-lg hover:bg-black/5"><CalendarPlus size={15} color={MUTED} /></button>
        <button title={cls.archived ? "إلغاء الأرشفة" : "أرشفة"} onClick={() => onArchive(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5">
          {cls.archived ? <ArchiveRestore size={15} color={MUTED} /> : <Archive size={15} color={MUTED} />}
        </button>
        <button title="حذف" onClick={() => onDelete(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5 mr-auto"><Trash2 size={15} color="#9A3B2E" /></button>
        {showColors && (
          <div className="absolute bottom-12 right-3 z-20 p-2 rounded-xl shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <ColorSwatches value={cls.color} onChange={(hex) => { onColor(cls.id, hex); setShowColors(false); }} />
          </div>
        )}
      </div>
    </div>
  );
}

function HomePage({ data, setData, onOpen, userEmail, onSignOut }) {
  const [modal, setModal] = useState(null);
  const [tab, setTab] = useState("active");
  const [showSettings, setShowSettings] = useState(false);
  const [showTodayActivity, setShowTodayActivity] = useState(false);
  const [showTestsList, setShowTestsList] = useState(false);
  const [showTestBuilder, setShowTestBuilder] = useState(false);
  const [gradingTestId, setGradingTestId] = useState(null);
  const [printingTestId, setPrintingTestId] = useState(null);
  const [omrTestId, setOmrTestId] = useState(null);
  const addTest = (test) => setData((d) => ({ ...d, tests: [...(d.tests || []), test] }));
  const deleteTest = (id) => setData((d) => ({ ...d, tests: (d.tests || []).filter((t) => t.id !== id) }));
  const saveTestResult = (testId, result) => setData((d) => ({ ...d, tests: (d.tests || []).map((t) => (t.id === testId ? { ...t, results: [...(t.results || []), result] } : t)) }));
  const applyResultToClass = (classId, colId, studentName, value) => {
    setData((d) => ({
      ...d,
      classes: d.classes.map((c) => {
        if (c.id !== classId) return c;
        let row = c.rows.find((r) => r.name.trim() === studentName.trim());
        let rows = c.rows;
        if (!row) {
          row = { id: uid(), name: studentName.trim(), type: "text", options: [], color: COLORS[4].hex, autoRenew: false };
          rows = [...rows, row];
        }
        const meta = nowMeta();
        const col = c.columns.find((cc) => cc.id === colId);
        const cells = { ...c.cells, [`${row.id}:${colId}`]: String(value) };
        const entry = { id: uid(), colId, colName: col?.name || "", colColor: col?.color || "#666", value: String(value), ...meta };
        const reports = { ...(c.reports || {}), [row.id]: [...(c.reports?.[row.id] || []), entry] };
        return { ...c, rows, cells, reports };
      }),
    }));
  };
  const [showGuide, setShowGuide] = useState(false);
  useEffect(() => {
    if (data.classes.length === 0 && !data.settings?.guideShown) {
      setShowGuide(true);
      setData((d) => ({ ...d, settings: { ...(d.settings || {}), guideShown: true } }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const addFooterContact = () => setData((d) => ({ ...d, settings: { ...(d.settings || {}), footerContacts: [...((d.settings || {}).footerContacts || []), { id: uid(), label: "", value: "" }] } }));
  const updateFooterContact = (id, patch) => setData((d) => ({ ...d, settings: { ...(d.settings || {}), footerContacts: ((d.settings || {}).footerContacts || []).map((c) => (c.id === id ? { ...c, ...patch } : c)) } }));
  const removeFooterContact = (id) => setData((d) => ({ ...d, settings: { ...(d.settings || {}), footerContacts: ((d.settings || {}).footerContacts || []).filter((c) => c.id !== id) } }));
  const addFooterBadge = (image) => setData((d) => ({ ...d, settings: { ...(d.settings || {}), footerBadges: [...((d.settings || {}).footerBadges || []), { id: uid(), image }] } }));
  const removeFooterBadge = (id) => setData((d) => ({ ...d, settings: { ...(d.settings || {}), footerBadges: ((d.settings || {}).footerBadges || []).filter((b) => b.id !== id) } }));
  const classes = data.classes.filter((c) => (tab === "active" ? !c.archived : c.archived));
  const displayClasses = [...classes.filter((c) => c.pinned), ...classes.filter((c) => !c.pinned)];
  const unpinnedIds = classes.filter((c) => !c.pinned).map((c) => c.id);

  const saveClass = (payload) => {
    setData((d) => {
      if (modal.mode === "edit") return { ...d, classes: d.classes.map((c) => (c.id === modal.data.id ? { ...c, ...payload } : c)) };
      const newClass = { id: uid(), ...payload, archived: false, columns: [], rows: [], cells: {}, reports: {}, showRowNumbers: false, emoji: payload.emoji || "" };
      return { ...d, classes: [newClass, ...d.classes] };
    });
    setModal(null);
  };
  const updateColor = (id, hex) => setData((d) => ({ ...d, classes: d.classes.map((c) => (c.id === id ? { ...c, color: hex } : c)) }));
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [animatingClassId, setAnimatingClassId] = useState(null);
  const deleteClass = (id) => setConfirmDeleteId(id);
  const confirmDeleteClass = () => {
    const id = confirmDeleteId;
    setConfirmDeleteId(null);
    setAnimatingClassId(id);
    setTimeout(() => {
      setData((d) => {
        const idx = d.classes.findIndex((c) => c.id === id);
        if (idx === -1) return d;
        const cls = d.classes[idx];
        const trashEntry = { id: uid(), type: "class", when: nowMeta(), data: { cls, index: idx } };
        return { ...d, classes: d.classes.filter((c) => c.id !== id), trash: [trashEntry, ...(d.trash || [])].slice(0, 30) };
      });
      setAnimatingClassId(null);
    }, 450);
  };
  const toggleArchive = (id) => setData((d) => ({ ...d, classes: d.classes.map((c) => (c.id === id ? { ...c, archived: !c.archived } : c)) }));
  const archiveAllClasses = () => {
    const ids = classes.map((c) => c.id);
    if (ids.length === 0) return;
    setData((d) => ({ ...d, classes: d.classes.map((c) => (ids.includes(c.id) ? { ...c, archived: true } : c)) }));
  };
  const [confirmDeleteAllClasses, setConfirmDeleteAllClasses] = useState(false);
  const deleteAllClasses = () => {
    const ids = classes.map((c) => c.id);
    if (ids.length === 0) return;
    setData((d) => {
      const removed = d.classes.filter((c) => ids.includes(c.id));
      const remaining = d.classes.filter((c) => !ids.includes(c.id));
      const trashEntry = { id: uid(), type: "classesBulk", when: nowMeta(), data: { classes: removed } };
      return { ...d, classes: remaining, trash: [trashEntry, ...(d.trash || [])].slice(0, 30) };
    });
    setConfirmDeleteAllClasses(false);
  };
  const duplicateClass = (id) => {
    setData((d) => {
      const original = d.classes.find((c) => c.id === id);
      if (!original) return d;
      const copy = JSON.parse(JSON.stringify(original));
      copy.id = uid();
      copy.subject = `${original.subject} (نسخة)`;
      copy.archived = false;
      return { ...d, classes: [copy, ...d.classes] };
    });
  };
  const [newTermModalCls, setNewTermModalCls] = useState(null);
  const startNewTerm = (id, years) => {
    setData((d) => {
      const original = d.classes.find((c) => c.id === id);
      if (!original) return d;
      const fresh = JSON.parse(JSON.stringify(original));
      fresh.id = uid();
      fresh.cells = {};
      fresh.reports = {};
      fresh.reportTrash = {};
      fresh.attendance = {};
      fresh.trash = [];
      fresh.events = [];
      fresh.archived = false;
      fresh.locked = false;
      fresh.pinned = false;
      if (years?.yearHijri) fresh.yearHijri = years.yearHijri;
      if (years?.yearGregorian) fresh.yearGregorian = years.yearGregorian;
      const classes = d.classes.map((c) => (c.id === id ? { ...c, archived: true } : c));
      return { ...d, classes: [fresh, ...classes] };
    });
    setNewTermModalCls(null);
  };
  const togglePinClass = (id) => {
    setData((d) => {
      const target = d.classes.find((c) => c.id === id);
      if (!target) return d;
      const updated = { ...target, pinned: !target.pinned };
      const others = d.classes.filter((c) => c.id !== id);
      const pinnedOthers = others.filter((c) => c.pinned);
      const rest = others.filter((c) => !c.pinned);
      return { ...d, classes: [...pinnedOthers, updated, ...rest] };
    });
  };
  const toggleLockClass = (id) => setData((d) => ({ ...d, classes: d.classes.map((c) => (c.id === id ? { ...c, locked: !c.locked } : c)) }));
  const moveClass = (id, direction) => {
    setData((d) => {
      const target = d.classes.find((c) => c.id === id);
      if (!target || target.pinned) return d;
      const groupIds = d.classes.filter((c) => c.archived === target.archived && !c.pinned).map((c) => c.id);
      const pos = groupIds.indexOf(id);
      const swapWithId = groupIds[pos + direction];
      if (!swapWithId) return d;
      const classes = [...d.classes];
      const idxA = classes.findIndex((c) => c.id === id);
      const idxB = classes.findIndex((c) => c.id === swapWithId);
      [classes[idxA], classes[idxB]] = [classes[idxB], classes[idxA]];
      return { ...d, classes };
    });
  };

  const restoreClassEntry = (entryId) => {
    setData((d) => {
      const entry = (d.trash || []).find((t) => t.id === entryId);
      if (!entry) return d;
      let classes;
      if (entry.type === "classesBulk") {
        classes = [...entry.data.classes, ...d.classes];
      } else {
        const idx = Math.min(entry.data.index, d.classes.length);
        classes = [...d.classes];
        classes.splice(idx, 0, entry.data.cls);
      }
      return { ...d, classes, trash: d.trash.filter((t) => t.id !== entryId) };
    });
  };
  const restoreLatestClass = () => {
    if (data.trash && data.trash.length > 0) restoreClassEntry(data.trash[0].id);
  };
  const [showClassTrash, setShowClassTrash] = useState(false);

  const [showSchedule, setShowSchedule] = useState(false);
  const setScheduleCell = (key, val) => setData((d) => ({ ...d, schedule: { ...(d.schedule || {}), [key]: val } }));
  const setScheduleImage = (image) => setData((d) => ({ ...d, scheduleImage: image }));
  const clearSchedule = () => setData((d) => ({ ...d, schedule: {} }));
  const deleteSchedule = () => setData((d) => ({ ...d, schedule: {}, scheduleImage: null }));

  const [showSearch, setShowSearch] = useState(false);
  const [globalSearch, setGlobalSearch] = useState("");
  const searchResults = globalSearch.trim()
    ? data.classes.flatMap((c) => (c.rows || []).filter((r) => r.name.toLowerCase().includes(globalSearch.trim().toLowerCase())).map((r) => ({ cls: c, row: r })))
    : [];

  return (
    <div className="max-w-5xl mx-auto px-4 py-6 page-fade-in">
      <div className="sticky top-0 z-20 pb-2" style={{ background: PAPER }}>
        <div className="flex items-center justify-between mb-5">
          <div>
            <h1 className="text-2xl font-extrabold" style={{ color: INK, fontFamily: "'Cairo', sans-serif" }}>دفتر المتابعة</h1>
            <p className="text-sm mt-1" style={{ color: MUTED }}>فصولك الدراسية في مكان واحد</p>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={() => setShowGuide(true)} title="كيف أبدأ؟" className="w-10 h-10 rounded-full flex items-center justify-center hover:bg-black/5" style={{ border: `1px solid ${LINE}` }}>
              <HelpCircle size={18} color={MUTED} />
            </button>
            <button onClick={() => setShowSettings(true)} title="الإعدادات" className="w-10 h-10 rounded-full flex items-center justify-center hover:bg-black/5" style={{ border: `1px solid ${LINE}` }}>
              <Settings size={18} color={MUTED} />
            </button>
            <button onClick={() => setModal({ mode: "add" })} className="w-12 h-12 rounded-full flex items-center justify-center shadow-md hover:opacity-90 active:scale-95 transition-all" style={{ background: "#0F6B5C" }}>
              <Plus size={22} color="#fff" strokeWidth={2.5} />
            </button>
          </div>
        </div>

        <DashboardStrip classes={data.classes.filter((c) => !c.archived)} />

        <ScheduleMiniCard schedule={data.schedule} image={data.scheduleImage} onOpen={() => setShowSchedule(true)} />

        <div className="flex flex-wrap items-center gap-2 mb-2.5">
          <button onClick={() => setTab("active")} className="px-4 py-1.5 rounded-full text-sm font-semibold"
            style={{ background: tab === "active" ? INK : "transparent", color: tab === "active" ? "#fff" : MUTED, border: `1px solid ${tab === "active" ? INK : LINE}` }}>
            النشطة ({data.classes.filter((c) => !c.archived).length})
          </button>
          <button onClick={() => setTab("archived")} className="px-4 py-1.5 rounded-full text-sm font-semibold flex items-center gap-1.5"
            style={{ background: tab === "archived" ? INK : "transparent", color: tab === "archived" ? "#fff" : MUTED, border: `1px solid ${tab === "archived" ? INK : LINE}` }}>
            <FolderClock size={14} />المؤرشفة ({data.classes.filter((c) => c.archived).length})
          </button>
        </div>

        <div className="rounded-xl p-2.5 flex flex-wrap items-center gap-2" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>أدوات</span>
          <IconBtn icon={RotateCcw} label="تراجع" onClick={restoreLatestClass} />
          <IconBtn icon={FolderOpen} label="استعادة" onClick={() => setShowClassTrash(true)} />
          <IconBtn icon={Search} label="بحث عن طالب في كل الفصول" onClick={() => setShowSearch((s) => !s)} />
          <IconBtn icon={ListTodo} label="نشاطي اليوم" onClick={() => setShowTodayActivity(true)} />
          <IconBtn icon={ListChecks} label="الاختبارات" onClick={() => setShowTestsList(true)} />
          {classes.length > 0 && (
            <>
              <IconBtn icon={Archive} label={`أرشفة الكل (${classes.length})`} onClick={archiveAllClasses} />
              <IconBtn icon={Trash2} label={`حذف الكل (${classes.length})`} tone="danger" onClick={() => setConfirmDeleteAllClasses(true)} />
            </>
          )}
        </div>
      </div>

      <div className="mt-5">
      {showSearch && (
        <div className="mb-5">
          <div className="relative" style={{ maxWidth: "320px" }}>
            <input
              autoFocus
              value={globalSearch}
              onChange={(e) => setGlobalSearch(e.target.value)}
              placeholder="اكتب اسم الطالب..."
              style={{ ...inputStyle, width: "100%", paddingInlineEnd: "28px" }}
            />
            {globalSearch && (
              <button
                onClick={() => setGlobalSearch("")}
                title="مسح البحث"
                className="absolute top-1/2 -translate-y-1/2 p-0.5 rounded-full hover:bg-black/5"
                style={{ insetInlineEnd: "6px" }}
              >
                <X size={14} color={MUTED} />
              </button>
            )}
          </div>
          {globalSearch.trim() && (
            <div className="mt-3 space-y-2">
              {searchResults.length === 0 ? (
                <p className="text-sm p-3 rounded-xl" style={{ color: MUTED, background: "#fff", border: `1px solid ${LINE}` }}>لا يوجد طالب بهذا الاسم في أي فصل.</p>
              ) : (
                searchResults.map(({ cls, row }) => {
                  const status = attendanceStatus(cls, row.id, todayKey());
                  return (
                    <div key={`${cls.id}-${row.id}`} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                      <button
                        onClick={() => onOpen(cls.id)}
                        title="فتح الفصل"
                        className="w-full flex items-center gap-3 px-3 py-2.5 text-right hover:bg-black/5"
                        style={{ background: "#F3F1E9" }}
                      >
                        <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: row.color }} />
                        <span className="flex-1">
                          <span className="block text-sm font-bold" style={{ color: INK }}>{row.name}</span>
                          <span className="block text-xs" style={{ color: MUTED }}>{cls.emoji ? `${cls.emoji} ` : ""}{cls.subject} • {cls.grade}</span>
                        </span>
                        <span className="text-xs font-bold px-2 py-0.5 rounded-full shrink-0" style={{ background: status === "absent" ? "#F5DEDB" : "#E3F0ED", color: status === "absent" ? "#C0392B" : "#0F6B5C" }}>
                          {status === "absent" ? "غائب اليوم" : "حاضر اليوم"}
                        </span>
                      </button>
                      {cls.columns.length > 0 && (
                        <div className="overflow-x-auto">
                          <table className="w-full border-collapse" style={{ fontSize: 12 }}>
                            <thead>
                              <tr>
                                {cls.columns.map((col) => (
                                  <th key={col.id} className="p-2 text-center font-semibold" style={{ background: colorLight(col.color), border: `1px solid ${LINE}`, color: INK, whiteSpace: "nowrap" }}>{col.name}</th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              <tr>
                                {cls.columns.map((col) => (
                                  <td key={col.id} className="p-2 text-center" style={{ border: `1px solid ${LINE}`, color: INK }}>{boardCellValue(cls, row.id, col.id) || "—"}</td>
                                ))}
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          )}
        </div>
      )}

      {classes.length === 0 ? (
        <div className="text-center py-20" style={{ color: MUTED }}>
          {tab === "active" ? (<><p className="font-semibold mb-1">لا توجد فصول بعد</p><p className="text-sm">اضغط على علامة + لإضافة أول فصل دراسي</p></>) : (<p className="text-sm">لا توجد فصول مؤرشفة</p>)}
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {displayClasses.map((cls) => (
            <ClassCard
              key={cls.id}
              cls={cls}
              onOpen={onOpen}
              onEdit={(c) => setModal({ mode: "edit", data: c })}
              onColor={updateColor}
              onDelete={deleteClass}
              onArchive={toggleArchive}
              onDuplicate={duplicateClass}
              onNewTerm={(c) => setNewTermModalCls(c)}
              onTogglePin={togglePinClass}
              onToggleLock={toggleLockClass}
              onMove={moveClass}
              isFirst={unpinnedIds.indexOf(cls.id) === 0}
              isLast={unpinnedIds.indexOf(cls.id) === unpinnedIds.length - 1}
              animating={cls.id === animatingClassId}
            />
          ))}
        </div>
      )}
      </div>
      {modal && (
        <ClassModal
          initial={modal.mode === "edit" ? modal.data : null}
          onClose={() => setModal(null)}
          onSave={saveClass}
          existingClasses={modal.mode === "add" ? data.classes : []}
          onDuplicateExisting={(id) => { duplicateClass(id); setModal(null); }}
        />
      )}
      {confirmDeleteId && (
        <ConfirmDialog
          title="حذف الفصل"
          message="سيتم حذف هذا الفصل وجميع بياناته نهائيًا. هل أنت متأكد؟"
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={confirmDeleteClass}
        />
      )}
      {confirmDeleteAllClasses && (
        <ConfirmDialog
          title="حذف كل الفصول"
          message={`سيتم حذف ${classes.length} فصل وكل بياناتها. يمكن استعادتها كلها دفعة واحدة من "استعادة". متابعة؟`}
          onCancel={() => setConfirmDeleteAllClasses(false)}
          onConfirm={deleteAllClasses}
        />
      )}
      {showClassTrash && (
        <TrashModal
          trash={data.trash}
          onClose={() => setShowClassTrash(false)}
          onRestore={(id) => restoreClassEntry(id)}
          onClearAll={() => setData((d) => ({ ...d, trash: [] }))}
        />
      )}
      {showSchedule && (
        <ScheduleModal
          schedule={data.schedule}
          onCellChange={setScheduleCell}
          image={data.scheduleImage}
          onSetImage={setScheduleImage}
          onClear={clearSchedule}
          onDeleteAll={deleteSchedule}
          onClose={() => setShowSchedule(false)}
        />
      )}
      {showSettings && (
        <SettingsModal
          feedback={data.settings?.feedback !== false}
          onToggleFeedback={() => setData((d) => ({ ...d, settings: { ...(d.settings || {}), feedback: !(d.settings?.feedback !== false) } }))}
          darkMode={!!data.settings?.darkMode}
          onToggleDarkMode={() => setData((d) => ({ ...d, settings: { ...(d.settings || {}), darkMode: !d.settings?.darkMode } }))}
          schoolName={data.settings?.schoolName}
          principalName={data.settings?.principalName}
          countryName={data.settings?.countryName}
          ministryName={data.settings?.ministryName}
          logoImage={data.settings?.logoImage}
          onChangeSchoolInfo={(patch) => setData((d) => ({ ...d, settings: { ...(d.settings || {}), ...patch } }))}
          footerContacts={data.settings?.footerContacts}
          footerBadges={data.settings?.footerBadges}
          onAddContact={addFooterContact}
          onUpdateContact={updateFooterContact}
          onRemoveContact={removeFooterContact}
          onAddBadge={addFooterBadge}
          onRemoveBadge={removeFooterBadge}
          userEmail={userEmail}
          onSignOut={onSignOut}
          onClose={() => setShowSettings(false)}
        />
      )}
      {newTermModalCls && (
        <NewTermModal
          cls={newTermModalCls}
          onClose={() => setNewTermModalCls(null)}
          onConfirm={(years) => startNewTerm(newTermModalCls.id, years)}
        />
      )}
      {showTodayActivity && <TodayActivityModal classes={data.classes.filter((c) => !c.archived)} onClose={() => setShowTodayActivity(false)} />}
      {showGuide && <GuideModal onClose={() => setShowGuide(false)} />}
      {showTestsList && (
        <TestsListModal
          tests={data.tests || []}
          onCreateNew={() => { setShowTestsList(false); setShowTestBuilder(true); }}
          onGrade={(id) => { setShowTestsList(false); setGradingTestId(id); }}
          onGradeCamera={(id) => { setShowTestsList(false); setOmrTestId(id); }}
          onPrint={(id) => { setShowTestsList(false); setPrintingTestId(id); }}
          onDelete={deleteTest}
          onClose={() => setShowTestsList(false)}
        />
      )}
      {showTestBuilder && (
        <TestBuilderModal
          onSave={(test) => { addTest(test); setShowTestBuilder(false); setShowTestsList(true); }}
          onClose={() => { setShowTestBuilder(false); setShowTestsList(true); }}
        />
      )}
      {gradingTestId && (
        <GradeTestModal
          test={(data.tests || []).find((t) => t.id === gradingTestId)}
          classes={data.classes}
          onSaveResult={(result) => saveTestResult(gradingTestId, result)}
          onApplyToClass={applyResultToClass}
          onClose={() => { setGradingTestId(null); setShowTestsList(true); }}
        />
      )}
      {omrTestId && (
        <OMRScanModal
          test={(data.tests || []).find((t) => t.id === omrTestId)}
          classes={data.classes}
          onSaveResult={(result) => saveTestResult(omrTestId, result)}
          onApplyToClass={applyResultToClass}
          onClose={() => { setOmrTestId(null); setShowTestsList(true); }}
        />
      )}
      {printingTestId && (
        <PrintTestModal
          test={(data.tests || []).find((t) => t.id === printingTestId)}
          classes={data.classes}
          onClose={() => { setPrintingTestId(null); setShowTestsList(true); }}
        />
      )}
      <SiteFooter contacts={data.settings?.footerContacts} badges={data.settings?.footerBadges} />
    </div>
  );
}

// ---------- Class detail page ----------

function ClassPage({ cls, updateClass, onBack, requestPrint, feedbackEnabled, schoolName, principalName, countryName, ministryName, logoImage, allClasses, onMoveRowsToClass }) {
  const [colModal, setColModal] = useState(null);
  const [rowModal, setRowModal] = useState(null);
  const [showSearch, setShowSearch] = useState(false);
  const [search, setSearch] = useState("");
  const [showBoard, setShowBoard] = useState(false);
  const [reportRowId, setReportRowId] = useState(null);
  const [confirmAction, setConfirmAction] = useState(null); // { type: 'deleteRow'|'deleteAll', row? }
  const [animatingRowId, setAnimatingRowId] = useState(null);
  const [animatingColId, setAnimatingColId] = useState(null);
  const [tableFading, setTableFading] = useState(false);
  const [printChoice, setPrintChoice] = useState(null);
  const [showAttendance, setShowAttendance] = useState(false);
  const [showEvents, setShowEvents] = useState(false);
  const [showRandomPicker, setShowRandomPicker] = useState(false);
  const [showRandomGroups, setShowRandomGroups] = useState(false);
  const [selectedRowIds, setSelectedRowIds] = useState([]);
  const [activeFilter, setActiveFilter] = useState(null);
  const [showFilterModal, setShowFilterModal] = useState(false);
  const [showBulkRecordModal, setShowBulkRecordModal] = useState(false);
  const [showBulkCertificateModal, setShowBulkCertificateModal] = useState(false);
  const [showMoveModal, setShowMoveModal] = useState(false);
  const [confirmBulkDelete, setConfirmBulkDelete] = useState(false);
  const [blinkRowId, setBlinkRowId] = useState(null);
  const timers = useRef({});

  const setCell = (row, col, value) => {
    updateClass((c) => {
      const next = { ...c, cells: { ...c.cells, [`${row.id}:${col.id}`]: value } };
      if (value && value.trim()) {
        const meta = nowMeta();
        const entry = { id: uid(), colId: col.id, colName: col.name, colColor: col.color, value, ...meta };
        const list = next.reports?.[row.id] || [];
        next.reports = { ...(next.reports || {}), [row.id]: [...list, entry] };
      }
      return next;
    });
    if (value && value.trim()) playFeedback(feedbackEnabled);
    if (row.autoRenew && col.autoRenew && value && value.trim()) {
      const key = `${row.id}:${col.id}`;
      if (timers.current[key]) clearTimeout(timers.current[key]);
      timers.current[key] = setTimeout(() => {
        updateClass((c) => ({ ...c, cells: { ...c.cells, [key]: "" } }));
      }, 2500);
    }
  };

  // Quick "record for all" action from the column header: sets the same
  // value for every student's cell in that column in one shot.
  const bulkSetColumnValue = (col, value) => {
    const v = (value || "").trim();
    if (!v) return;
    updateClass((c) => {
      const cells = { ...c.cells };
      const reports = { ...(c.reports || {}) };
      const meta = nowMeta();
      c.rows.forEach((row) => {
        cells[`${row.id}:${col.id}`] = v;
        const entry = { id: uid(), colId: col.id, colName: col.name, colColor: col.color, value: v, ...meta };
        reports[row.id] = [...(reports[row.id] || []), entry];
      });
      return { ...c, cells, reports };
    });
    playFeedback(feedbackEnabled);
  };
  const [bulkSetColId, setBulkSetColId] = useState(null);

  const toggleRowSelection = (rowId) => {
    setSelectedRowIds((ids) => (ids.includes(rowId) ? ids.filter((id) => id !== rowId) : [...ids, rowId]));
  };
  const clearSelection = () => setSelectedRowIds([]);
  const selectAllRows = () => setSelectedRowIds(cls.rows.map((r) => r.id));
  const bulkSetColumnValueForSelected = (col, value) => {
    const v = (value || "").trim();
    if (!v || selectedRowIds.length === 0) return;
    updateClass((c) => {
      const cells = { ...c.cells };
      const reports = { ...(c.reports || {}) };
      const meta = nowMeta();
      selectedRowIds.forEach((rowId) => {
        cells[`${rowId}:${col.id}`] = v;
        const entry = { id: uid(), colId: col.id, colName: col.name, colColor: col.color, value: v, ...meta };
        reports[rowId] = [...(reports[rowId] || []), entry];
      });
      return { ...c, cells, reports };
    });
    playFeedback(feedbackEnabled);
  };
  const bulkDeleteSelected = () => {
    const ids = selectedRowIds;
    if (ids.length === 0) return;
    updateClass((c) => {
      const items = [];
      ids.forEach((rowId) => {
        const idx = c.rows.findIndex((r) => r.id === rowId);
        if (idx === -1) return;
        const row = c.rows[idx];
        const cellsForRow = Object.fromEntries(Object.entries(c.cells).filter(([k]) => k.startsWith(`${rowId}:`)));
        const reportsForRow = c.reports?.[rowId];
        items.push({ row, index: idx, cells: cellsForRow, reports: reportsForRow });
      });
      const rows = c.rows.filter((r) => !ids.includes(r.id));
      const cells = Object.fromEntries(Object.entries(c.cells).filter(([k]) => !ids.some((id) => k.startsWith(`${id}:`))));
      const reports = { ...(c.reports || {}) };
      ids.forEach((id) => delete reports[id]);
      const next = { ...c, rows, cells, reports };
      return pushTrash(next, { id: uid(), type: "bulkRows", when: nowMeta(), data: { items } });
    });
    clearSelection();
  };

  const saveColumnsMany = (drafts) => {
    updateClass((c) => {
      const newCols = drafts.map((d) => ({ id: uid(), name: d.name.trim(), type: d.type, options: d.options, color: d.color, autoRenew: !!d.autoRenew, pinned: !!d.pinned, behaviorFlag: !!d.behaviorFlag, behaviorThreshold: d.behaviorThreshold || 3 }));
      const cells = { ...c.cells };
      const reports = { ...(c.reports || {}) };
      let touched = false;
      newCols.forEach((col, i) => {
        const bulkValue = (drafts[i].bulkValue || "").trim();
        if (!bulkValue) return;
        touched = true;
        const meta = nowMeta();
        c.rows.forEach((row) => {
          cells[`${row.id}:${col.id}`] = bulkValue;
          const entry = { id: uid(), colId: col.id, colName: col.name, colColor: col.color, value: bulkValue, ...meta };
          reports[row.id] = [...(reports[row.id] || []), entry];
        });
      });
      if (touched) playFeedback(feedbackEnabled);
      return { ...c, columns: [...c.columns, ...newCols], cells, reports };
    });
    setColModal(null);
  };
  const saveColumnOne = (payload) => {
    updateClass((c) => {
      const columns = c.columns.map((col) => (col.id === colModal.data.id ? { ...col, ...payload, name: payload.name.trim() } : col));
      const bulkValue = (payload.bulkValue || "").trim();
      const cells = { ...c.cells };
      const reports = { ...(c.reports || {}) };
      if (bulkValue) {
        const col = columns.find((cc) => cc.id === colModal.data.id);
        const meta = nowMeta();
        c.rows.forEach((row) => {
          cells[`${row.id}:${col.id}`] = bulkValue;
          const entry = { id: uid(), colId: col.id, colName: col.name, colColor: col.color, value: bulkValue, ...meta };
          reports[row.id] = [...(reports[row.id] || []), entry];
        });
        playFeedback(feedbackEnabled);
      }
      return { ...c, columns, cells, reports };
    });
    setColModal(null);
  };
  const pushTrash = (c, entry) => ({ ...c, trash: [entry, ...(c.trash || [])].slice(0, 30) });

  const deleteColumn = () => {
    const column = colModal.data;
    setColModal(null);
    setConfirmAction({ type: "deleteColumn", column });
  };

  const saveRowsMany = (drafts) => {
    updateClass((c) => ({ ...c, rows: [...c.rows, ...drafts.map((d) => ({ id: uid(), name: d.name.trim(), type: d.type, options: d.options, color: d.color, autoRenew: !!d.autoRenew, medicalNote: d.medicalNote || "" }))] }));
    setRowModal(null);
  };
  const saveRowOne = (payload) => {
    updateClass((c) => ({ ...c, rows: c.rows.map((r) => (r.id === rowModal.data.id ? { ...r, ...payload, name: payload.name.trim() } : r)) }));
    setRowModal(null);
  };
  const removeRowById = (rowId) => {
    updateClass((c) => {
      const idx = c.rows.findIndex((r) => r.id === rowId);
      if (idx === -1) return c;
      const row = c.rows[idx];
      const cellsForRow = Object.fromEntries(Object.entries(c.cells).filter(([k]) => k.startsWith(`${rowId}:`)));
      const reportsForRow = c.reports?.[rowId];
      const rest = { ...(c.reports || {}) };
      delete rest[rowId];
      const next = {
        ...c,
        rows: c.rows.filter((r) => r.id !== rowId),
        cells: Object.fromEntries(Object.entries(c.cells).filter(([k]) => !k.startsWith(`${rowId}:`))),
        reports: rest,
      };
      return pushTrash(next, { id: uid(), type: "row", when: nowMeta(), data: { row, index: idx, cells: cellsForRow, reports: reportsForRow } });
    });
  };
  const markAbsentToday = (rowId) => {
    updateClass(markAbsentUpdater(rowId, todayKey()));
    playFeedback(feedbackEnabled);
    setBlinkRowId(rowId);
    setTimeout(() => setBlinkRowId(null), 2500);
  };

  const addEvent = (text) => updateClass((c) => ({ ...c, events: [...(c.events || []), { id: uid(), text }] }));
  const updateEvent = (id, text) => updateClass((c) => ({ ...c, events: (c.events || []).map((e) => (e.id === id ? { ...e, text } : e)) }));
  const deleteEvent = (id) => updateClass((c) => ({ ...c, events: (c.events || []).filter((e) => e.id !== id) }));
  const moveEvent = (id, direction) => {
    updateClass((c) => {
      const list = c.events || [];
      const idx = list.findIndex((e) => e.id === id);
      const newIdx = idx + direction;
      if (idx === -1 || newIdx < 0 || newIdx >= list.length) return c;
      const events = [...list];
      [events[idx], events[newIdx]] = [events[newIdx], events[idx]];
      return { ...c, events };
    });
  };

  const duplicateRowById = (rowId) => {
    updateClass((c) => {
      const idx = c.rows.findIndex((r) => r.id === rowId);
      if (idx === -1) return c;
      const orig = c.rows[idx];
      const newId = uid();
      const newRow = { ...orig, id: newId, name: `${orig.name} (نسخة)` };
      const rows = [...c.rows];
      rows.splice(idx + 1, 0, newRow);
      const copiedCells = {};
      Object.entries(c.cells).forEach(([k, v]) => {
        if (k.startsWith(`${rowId}:`)) {
          const colId = k.split(":")[1];
          copiedCells[`${newId}:${colId}`] = v;
        }
      });
      return { ...c, rows, cells: { ...c.cells, ...copiedCells } };
    });
  };
  const deleteRow = () => {
    const row = rowModal.data;
    setRowModal(null);
    setConfirmAction({ type: "deleteRow", row });
  };
  const quickDeleteRow = (row) => setConfirmAction({ type: "deleteRow", row });

  const deleteAll = () => setConfirmAction({ type: "deleteAll" });

  const runConfirmedAction = () => {
    if (!confirmAction) return;
    if (confirmAction.type === "deleteRow") {
      const id = confirmAction.row.id;
      setConfirmAction(null);
      setAnimatingRowId(id);
      setTimeout(() => { removeRowById(id); setAnimatingRowId(null); }, 450);
      return;
    }
    if (confirmAction.type === "deleteColumn") {
      const colId = confirmAction.column.id;
      setConfirmAction(null);
      setAnimatingColId(colId);
      setTimeout(() => {
        updateClass((c) => {
          const idx = c.columns.findIndex((col) => col.id === colId);
          if (idx === -1) return c;
          const column = c.columns[idx];
          const cellsForCol = Object.fromEntries(Object.entries(c.cells).filter(([k]) => k.endsWith(`:${colId}`)));
          const next = {
            ...c,
            columns: c.columns.filter((col) => col.id !== colId),
            cells: Object.fromEntries(Object.entries(c.cells).filter(([k]) => !k.endsWith(`:${colId}`))),
          };
          return pushTrash(next, { id: uid(), type: "column", when: nowMeta(), data: { column, index: idx, cells: cellsForCol } });
        });
        setAnimatingColId(null);
      }, 450);
      return;
    }
    if (confirmAction.type === "deleteAll") {
      setConfirmAction(null);
      setTableFading(true);
      setTimeout(() => {
        updateClass((c) => {
          const next = { ...c, columns: [], rows: [], cells: {}, reports: {} };
          return pushTrash(next, { id: uid(), type: "bulk", when: nowMeta(), data: { columns: c.columns, rows: c.rows, cells: c.cells, reports: c.reports } });
        });
        setTableFading(false);
      }, 450);
      return;
    }
    setConfirmAction(null);
  };

  const restoreEntry = (entryId) => {
    updateClass((c) => {
      const entry = (c.trash || []).find((t) => t.id === entryId);
      if (!entry) return c;
      const next = { ...c, trash: c.trash.filter((t) => t.id !== entryId) };
      if (entry.type === "row") {
        const idx = Math.min(entry.data.index, next.rows.length);
        const rows = [...next.rows];
        rows.splice(idx, 0, entry.data.row);
        next.rows = rows;
        next.cells = { ...next.cells, ...entry.data.cells };
        if (entry.data.reports) next.reports = { ...(next.reports || {}), [entry.data.row.id]: entry.data.reports };
      } else if (entry.type === "column") {
        const idx = Math.min(entry.data.index, next.columns.length);
        const columns = [...next.columns];
        columns.splice(idx, 0, entry.data.column);
        next.columns = columns;
        next.cells = { ...next.cells, ...entry.data.cells };
      } else if (entry.type === "bulk") {
        next.columns = entry.data.columns;
        next.rows = entry.data.rows;
        next.cells = entry.data.cells;
        next.reports = entry.data.reports;
      } else if (entry.type === "bulkRows") {
        const items = [...entry.data.items].sort((a, b) => a.index - b.index);
        const rows = [...next.rows];
        let cells = { ...next.cells };
        let reports = { ...(next.reports || {}) };
        items.forEach((item) => {
          const idx = Math.min(item.index, rows.length);
          rows.splice(idx, 0, item.row);
          cells = { ...cells, ...item.cells };
          if (item.reports) reports = { ...reports, [item.row.id]: item.reports };
        });
        next.rows = rows;
        next.cells = cells;
        next.reports = reports;
      }
      return next;
    });
  };

  const restoreLatest = () => {
    if (cls.trash && cls.trash.length > 0) restoreEntry(cls.trash[0].id);
  };
  const [showTrash, setShowTrash] = useState(false);

  // Per-student report-entry trash: deleting a single record from the
  // student's report keeps it recoverable, same pattern as the class-level trash.
  const removeReportEntry = (rowId, entryId) => {
    updateClass((c) => {
      const list = c.reports?.[rowId] || [];
      const entry = list.find((e) => e.id === entryId);
      if (!entry) return c;
      const reports = { ...(c.reports || {}), [rowId]: list.filter((e) => e.id !== entryId) };
      const trashItem = { id: uid(), type: "reportEntry", when: nowMeta(), data: { entry } };
      const rowTrash = [trashItem, ...((c.reportTrash || {})[rowId] || [])].slice(0, 30);
      const reportTrash = { ...(c.reportTrash || {}), [rowId]: rowTrash };
      return { ...c, reports, reportTrash };
    });
  };
  const restoreReportEntryFromTrash = (rowId, trashId) => {
    updateClass((c) => {
      const rowTrash = (c.reportTrash || {})[rowId] || [];
      const item = rowTrash.find((t) => t.id === trashId);
      if (!item) return c;
      const toAdd = item.type === "reportBulk" ? item.data.entries : [item.data.entry];
      const reports = { ...(c.reports || {}), [rowId]: [...(c.reports?.[rowId] || []), ...toAdd] };
      const reportTrash = { ...(c.reportTrash || {}), [rowId]: rowTrash.filter((t) => t.id !== trashId) };
      return { ...c, reports, reportTrash };
    });
  };
  const restoreLatestReportEntry = (rowId) => {
    const list = cls.reportTrash?.[rowId] || [];
    if (list.length > 0) restoreReportEntryFromTrash(rowId, list[0].id);
  };
  const clearReportTrash = (rowId) => {
    updateClass((c) => ({ ...c, reportTrash: { ...(c.reportTrash || {}), [rowId]: [] } }));
  };
  const updateReportEntryValue = (rowId, entryId, value) => {
    updateClass((c) => ({ ...c, reports: { ...(c.reports || {}), [rowId]: (c.reports?.[rowId] || []).map((e) => (e.id === entryId ? { ...e, value } : e)) } }));
  };
  const addManualReportNote = (rowId, entry) => {
    updateClass((c) => ({ ...c, reports: { ...(c.reports || {}), [rowId]: [...(c.reports?.[rowId] || []), entry] } }));
  };
  // Delete every entry in ONE category (e.g. all "المشاركات" records) for a
  // student, or the student's entire report — both recoverable as one trash item.
  const removeReportCategory = (rowId, colId) => {
    updateClass((c) => {
      const list = c.reports?.[rowId] || [];
      const toRemove = list.filter((e) => (e.colId || "note") === colId);
      if (toRemove.length === 0) return c;
      const reports = { ...(c.reports || {}), [rowId]: list.filter((e) => (e.colId || "note") !== colId) };
      const trashItem = { id: uid(), type: "reportBulk", when: nowMeta(), data: { entries: toRemove } };
      const rowTrash = [trashItem, ...((c.reportTrash || {})[rowId] || [])].slice(0, 30);
      const reportTrash = { ...(c.reportTrash || {}), [rowId]: rowTrash };
      return { ...c, reports, reportTrash };
    });
  };
  const removeAllReportEntries = (rowId) => {
    updateClass((c) => {
      const list = c.reports?.[rowId] || [];
      if (list.length === 0) return c;
      const reports = { ...(c.reports || {}), [rowId]: [] };
      const trashItem = { id: uid(), type: "reportBulk", when: nowMeta(), data: { entries: list } };
      const rowTrash = [trashItem, ...((c.reportTrash || {})[rowId] || [])].slice(0, 30);
      const reportTrash = { ...(c.reportTrash || {}), [rowId]: rowTrash };
      return { ...c, reports, reportTrash };
    });
  };

  const handleChooseFormat = (key) => {
    if (!printChoice) return;
    if (key === "pdf") exportPdfShare(printChoice);
    else if (key === "png") exportPng(printChoice);
    else if (key === "excel") exportExcel(printChoice);
    setPrintChoice(null);
  };

  const visibleRows = cls.rows.filter((r) => r.name.toLowerCase().includes(search.toLowerCase()) && rowMatchesFilter(cls, r, activeFilter));

  const moveColumn = (colId, direction) => {
    updateClass((c) => {
      const idx = c.columns.findIndex((col) => col.id === colId);
      const newIdx = idx + direction;
      if (idx === -1 || newIdx < 0 || newIdx >= c.columns.length) return c;
      const columns = [...c.columns];
      [columns[idx], columns[newIdx]] = [columns[newIdx], columns[idx]];
      return { ...c, columns };
    });
  };
  const moveRow = (rowId, direction) => {
    updateClass((c) => {
      const idx = c.rows.findIndex((r) => r.id === rowId);
      const newIdx = idx + direction;
      if (idx === -1 || newIdx < 0 || newIdx >= c.rows.length) return c;
      const rows = [...c.rows];
      [rows[idx], rows[newIdx]] = [rows[newIdx], rows[idx]];
      return { ...c, rows };
    });
  };

  const togglePinned = (colId) => {
    updateClass((c) => ({ ...c, columns: c.columns.map((col) => (col.id === colId ? { ...col, pinned: !col.pinned } : col)) }));
  };

  const reportRow = reportRowId ? cls.rows.find((r) => r.id === reportRowId) : null;
  const reportEntries = reportRowId ? cls.reports?.[reportRowId] || [] : [];

  const NUM_W = 48;
  const NAME_W = 180;
  const PIN_W = 160;
  const leadWidth = (cls.showRowNumbers ? NUM_W : 0) + NAME_W;
  let pinCounter = 0;
  const columnMeta = cls.columns.map((col) => {
    if (col.pinned) {
      const offset = leadWidth + PIN_W * pinCounter;
      pinCounter += 1;
      return { ...col, pinOffset: offset };
    }
    return { ...col, pinOffset: null };
  });

  return (
    <div className="max-w-[1800px] mx-auto px-4 py-6 page-fade-in">
      <div className="sticky top-0 z-20 pb-2" style={{ background: PAPER }}>
        <button onClick={onBack} className="flex items-center gap-1.5 text-sm font-semibold mb-3 hover:opacity-70" style={{ color: MUTED }}>
          <ArrowRight size={16} /> رجوع للفصول
        </button>

        <div className="rounded-2xl p-4 mb-3 flex flex-wrap items-center gap-x-8 gap-y-2" style={{ background: "#fff", border: `1px solid ${LINE}`, borderInlineStart: `6px solid ${cls.color}` }}>
          <div><p className="text-xs" style={{ color: MUTED }}>المادة</p><p className="font-bold flex items-center gap-1.5" style={{ color: INK }}>{cls.emoji && <span>{cls.emoji}</span>}{cls.subject}</p></div>
          <div><p className="text-xs" style={{ color: MUTED }}>الصف</p><p className="font-bold" style={{ color: INK }}>{cls.grade}</p></div>
          <div><p className="text-xs" style={{ color: MUTED }}>المعلم</p><p className="font-bold" style={{ color: INK }}>{cls.teacher}</p></div>
          <div><p className="text-xs" style={{ color: MUTED }}>العام الدراسي</p><p className="font-bold" style={{ color: INK }}>{cls.yearHijri} هـ / {cls.yearGregorian} م</p></div>
        </div>

        <div className="rounded-xl p-2.5 mb-2.5 flex flex-wrap items-center gap-2" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>العرض والطباعة</span>
          <IconBtn icon={LayoutGrid} label="لوحة العرض" onClick={() => setShowBoard(true)} />
          <IconBtn icon={Printer} label="طباعة" onClick={() => setPrintChoice({ type: "class", cls })} />
          <IconBtn icon={FileOutput} label="طباعة الجدول مفرغ" onClick={() => setPrintChoice({ type: "blank", cls })} />
        </div>

        <div className="rounded-xl p-2.5 mb-3 flex flex-wrap items-center gap-2" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>أدوات الحصة</span>
          <IconBtn icon={CalendarCheck} label="متابعة الحضور" onClick={() => setShowAttendance(true)} />
          <IconBtn icon={Newspaper} label="الأحداث" onClick={() => setShowEvents(true)} />
          <IconBtn icon={Shuffle} label="اختر لي طالبًا" onClick={() => setShowRandomPicker(true)} />
          <IconBtn icon={Users} label="مجموعات عشوائية" onClick={() => setShowRandomGroups(true)} />
        </div>

        <EventsTicker events={cls.events} speed={cls.tickerSpeed || 14} />

        <div className="rounded-xl p-2.5 flex flex-wrap items-center gap-2" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>إدارة الجدول</span>
          <IconBtn icon={Plus} label="إضافة عمود" tone="primary" onClick={() => setColModal({ mode: "add" })} />
          <IconBtn icon={Plus} label="إضافة صف" tone="primary" onClick={() => setRowModal({ mode: "add" })} />
          <IconBtn icon={RotateCcw} label="تراجع" onClick={restoreLatest} />
          <IconBtn icon={FolderOpen} label="استعادة" onClick={() => setShowTrash(true)} />
          <IconBtn icon={Trash2} label="حذف الكل" tone="danger" onClick={deleteAll} />
          <IconBtn icon={Search} label="بحث" onClick={() => setShowSearch((s) => !s)} />
          {showSearch && (
            <div className="relative" style={{ width: "200px" }}>
              <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ابحث عن اسم الطالب..." style={{ ...inputStyle, width: "100%", paddingInlineEnd: "28px" }} />
              {search && (
                <button
                  onClick={() => setSearch("")}
                  title="مسح البحث"
                  className="absolute top-1/2 -translate-y-1/2 p-0.5 rounded-full hover:bg-black/5"
                  style={{ insetInlineEnd: "6px" }}
                >
                  <X size={14} color={MUTED} />
                </button>
              )}
            </div>
          )}
          {cls.columns.length > 0 && (
            activeFilter ? (
              <div className="flex items-center gap-1.5 rounded-lg overflow-hidden" style={{ border: "1px solid #C9E2DB" }}>
                <button onClick={() => setShowFilterModal(true)} className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium" style={{ background: "#EAF3F0", color: "#0F6B5C" }}>
                  <Filter size={15} />
                  {cls.columns.find((c) => c.id === activeFilter.colId)?.name}
                </button>
                <button onClick={() => setActiveFilter(null)} title="إزالة التصفية" className="px-2 py-2 hover:bg-black/5" style={{ background: "#EAF3F0" }}>
                  <X size={14} color="#0F6B5C" />
                </button>
              </div>
            ) : (
              <IconBtn icon={Filter} label="تصفية" onClick={() => setShowFilterModal(true)} />
            )
          )}
        </div>
      </div>

      <div className="mt-3">
        {selectedRowIds.length > 0 && (
          <div className="flex flex-wrap items-center gap-2 mb-3 p-2.5 rounded-xl" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
            <span className="text-sm font-bold px-2" style={{ color: "#0F6B5C" }}>{selectedRowIds.length} طالب محدد</span>
            <button onClick={selectAllRows} className="text-xs font-semibold px-2 py-1 rounded-lg hover:bg-black/5" style={{ color: INK }}>تحديد الكل</button>
            <button onClick={clearSelection} className="text-xs font-semibold px-2 py-1 rounded-lg hover:bg-black/5" style={{ color: INK }}>إلغاء التحديد</button>
            <div className="mr-auto flex gap-2">
              <IconBtn icon={Send} label="نقل إلى فصل آخر" onClick={() => setShowMoveModal(true)} />
              <IconBtn icon={Award} label="شهادات جماعية" onClick={() => setShowBulkCertificateModal(true)} />
              <IconBtn icon={Users} label="رصد جماعي" onClick={() => setShowBulkRecordModal(true)} />
              <IconBtn icon={Trash2} label="حذف المحددين" tone="danger" onClick={() => setConfirmBulkDelete(true)} />
            </div>
          </div>
        )}

      <div className="rounded-2xl" style={{ background: "#fff", border: `1px solid ${LINE}`, maxHeight: "70vh", overflowX: "auto", overflowY: "auto" }}>
        {cls.columns.length === 0 || cls.rows.length === 0 ? (
          <div className="text-center py-16" style={{ color: MUTED }}>
            <p className="font-semibold mb-1">لا توجد بيانات بعد</p>
            <p className="text-sm">ابدأ بإضافة عمود (مثل: الواجب) ثم صف (مثل: اسم الطالب)</p>
          </div>
        ) : (
          <table className={`w-full text-sm ${tableFading ? "trash-toss" : ""}`} style={{ borderCollapse: "separate", borderSpacing: 0 }}>
            <thead>
              <tr>
                {cls.showRowNumbers && (
                  <th
                    className="p-2 text-center"
                    style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, position: "sticky", top: 0, insetInlineStart: 0, zIndex: 9, width: NUM_W, minWidth: NUM_W }}
                  >#</th>
                )}
                <th
                  className="p-2 text-right"
                  style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, position: "sticky", top: 0, insetInlineStart: cls.showRowNumbers ? NUM_W : 0, zIndex: 9, width: NAME_W, minWidth: NAME_W }}
                >الاسم</th>
                {columnMeta.map((col, i) => (
                  <th
                    key={col.id}
                    className={`p-2 text-center relative ${col.id === animatingColId ? "trash-toss" : ""}`}
                    style={{
                      background: colorLight(col.color),
                      border: `1px solid ${LINE}`,
                      color: INK,
                      minWidth: col.pinned ? PIN_W : "150px",
                      width: col.pinned ? PIN_W : undefined,
                      position: "sticky",
                      top: 0,
                      insetInlineStart: col.pinned ? col.pinOffset : undefined,
                      zIndex: col.pinned ? 9 : 7,
                    }}
                  >
                    <div className="flex items-center justify-center gap-1">
                      <MiniIconBtn icon={ChevronRight} title="نقل لليمين" disabled={i === 0} onClick={() => moveColumn(col.id, -1)} />
                      <MiniIconBtn icon={ChevronLeft} title="نقل لليسار" disabled={i === columnMeta.length - 1} onClick={() => moveColumn(col.id, 1)} />
                      <span className="font-semibold">{col.name}</span>
                      {col.autoRenew && <RefreshCw size={11} color="#0F6B5C" title="تفريغ تلقائي مفعّل" />}
                      <MiniIconBtn icon={col.pinned ? Pin : PinOff} title={col.pinned ? "إلغاء التثبيت" : "تثبيت العمود"} color={col.pinned ? "#0F6B5C" : MUTED} onClick={() => togglePinned(col.id)} />
                      <MiniIconBtn icon={Users} title="رصد نفس القيمة لجميع الطلاب" onClick={() => setBulkSetColId(bulkSetColId === col.id ? null : col.id)} />
                      <MiniIconBtn icon={Pencil} title="تعديل العمود" onClick={() => setColModal({ mode: "edit", data: col })} />
                    </div>
                    {bulkSetColId === col.id && (
                      <BulkSetPopover column={col} onApply={(v) => bulkSetColumnValue(col, v)} onClose={() => setBulkSetColId(null)} />
                    )}
                  </th>
                ))}
                <th className="p-1.5 text-center" style={{ background: "#FBEDEA", border: `1px solid ${LINE}`, color: "#9A3B2E", width: 60, minWidth: 60, position: "sticky", top: 0, insetInlineEnd: 125, zIndex: 9 }}>الغياب</th>
                <th className="p-1.5 text-center" style={{ background: "#EEEEE7", border: `1px solid ${LINE}`, color: INK, width: 70, minWidth: 70, position: "sticky", top: 0, insetInlineEnd: 55, zIndex: 9 }}>
                  <div className="flex items-center justify-center gap-1"><ClipboardList size={13} /><span className="font-semibold text-sm">تقرير</span></div>
                </th>
                <th className="p-1.5 text-center" title="حذف" style={{ background: "#FBEDEA", border: `1px solid ${LINE}`, color: "#9A3B2E", width: 55, minWidth: 55, position: "sticky", top: 0, insetInlineEnd: 0, zIndex: 9 }}>
                  <div className="flex items-center justify-center gap-1"><Trash2 size={13} /></div>
                </th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.map((row) => {
                const i = cls.rows.findIndex((r) => r.id === row.id);
                return (
                  <tr key={row.id} className={row.id === animatingRowId ? "trash-toss" : ""}>
                    {cls.showRowNumbers && (
                      <td
                        className="p-2 text-center text-xs font-semibold"
                        style={{ border: `1px solid ${LINE}`, color: MUTED, background: "#fff", position: "sticky", insetInlineStart: 0, zIndex: 3, width: NUM_W, minWidth: NUM_W }}
                      >{i + 1}</td>
                    )}
                    <td
                      className="p-2 font-medium"
                      style={{ border: `1px solid ${LINE}`, color: INK, borderInlineStart: `4px solid ${row.color}`, background: "#fff", position: "sticky", insetInlineStart: cls.showRowNumbers ? NUM_W : 0, zIndex: 3, width: NAME_W, minWidth: NAME_W }}
                    >
                      <div className="flex items-center gap-1.5 mb-1">
                        <input
                          type="checkbox"
                          checked={selectedRowIds.includes(row.id)}
                          onChange={() => toggleRowSelection(row.id)}
                          className="shrink-0"
                        />
                        <span className="font-semibold truncate flex-1">{row.name}</span>
                        {row.medicalNote && row.medicalNote.trim() && (
                          <span title={`تنبيه: ${row.medicalNote}`} className="shrink-0"><AlertTriangle size={12} color="#C97A2B" /></span>
                        )}
                        {row.autoRenew && <RefreshCw size={11} color="#0F6B5C" title="تجديد تلقائي مفعّل" className="shrink-0" />}
                      </div>
                      <div className="flex items-center gap-0.5" style={{ paddingInlineStart: 20 }}>
                        <MiniIconBtn icon={ChevronUp} title="نقل لأعلى" disabled={i === 0} onClick={() => moveRow(row.id, -1)} />
                        <MiniIconBtn icon={ChevronDown} title="نقل لأسفل" disabled={i === cls.rows.length - 1} onClick={() => moveRow(row.id, 1)} />
                        <span className="mx-1" style={{ width: 1, height: 14, background: LINE, display: "inline-block" }} />
                        <MiniIconBtn icon={Copy} title="نسخ الصف" onClick={() => duplicateRowById(row.id)} />
                        <MiniIconBtn icon={Pencil} title="تعديل الصف" onClick={() => setRowModal({ mode: "edit", data: row })} />
                      </div>
                    </td>
                    {columnMeta.map((col) => (
                      <td
                        key={col.id}
                        className={`p-1 text-center ${col.id === animatingColId ? "trash-toss" : ""}`}
                        style={{
                          border: `1px solid ${LINE}`,
                          background: col.pinned ? "#FBFAF6" : "#fff",
                          position: col.pinned ? "sticky" : "static",
                          insetInlineStart: col.pinned ? col.pinOffset : undefined,
                          zIndex: col.pinned ? 2 : 0,
                        }}
                      >
                        <Cell column={col} value={cls.cells[`${row.id}:${col.id}`]} onChange={(v) => setCell(row, col, v)} />
                      </td>
                    ))}
                    <td className="p-1 text-center" style={{ border: `1px solid ${LINE}`, background: "#fff", position: "sticky", insetInlineEnd: 125, zIndex: 2 }}>
                      {blinkRowId === row.id ? (
                        <span className="inline-block w-10 h-4" />
                      ) : (
                        <button onClick={() => markAbsentToday(row.id)} className="text-xs font-bold hover:opacity-70" style={{ color: "#C0392B" }}>غياب</button>
                      )}
                    </td>
                    <td className="p-1 text-center" style={{ border: `1px solid ${LINE}`, background: "#fff", position: "sticky", insetInlineEnd: 55, zIndex: 2 }}>
                      <button onClick={() => setReportRowId(row.id)} title="عرض التقرير" className="flex items-center justify-center mx-auto p-1.5 rounded-md hover:bg-black/5" style={{ border: `1px solid ${LINE}`, color: INK }}>
                        <FileText size={13} />
                      </button>
                    </td>
                    <td className="p-1 text-center" style={{ border: `1px solid ${LINE}`, background: "#fff", position: "sticky", insetInlineEnd: 0, zIndex: 2 }}>
                      <button onClick={() => quickDeleteRow(row)} title="حذف الصف بالكامل" className="p-1 rounded-lg hover:bg-black/5 mx-auto">
                        <Trash2 size={14} color="#C0392B" />
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
      </div>

      {colModal && <ColumnModal initial={colModal.mode === "edit" ? colModal.data : null} onClose={() => setColModal(null)} onSaveMany={saveColumnsMany} onSaveOne={saveColumnOne} onDelete={deleteColumn} />}
      {rowModal && (
        <RowModal
          initial={rowModal.mode === "edit" ? rowModal.data : null}
          onClose={() => setRowModal(null)}
          onSaveMany={saveRowsMany}
          onSaveOne={saveRowOne}
          onDelete={deleteRow}
          showRowNumbers={cls.showRowNumbers}
          onToggleShowRowNumbers={() => updateClass((c) => ({ ...c, showRowNumbers: !c.showRowNumbers }))}
        />
      )}
      {showBoard && <DisplayBoard cls={cls} onClose={() => setShowBoard(false)} onPrint={(dateKey) => setPrintChoice({ type: "class", cls, dateKey })} />}
      {reportRow && (
        <ReportModal
          cls={cls}
          row={reportRow}
          entries={reportEntries}
          reportTrash={cls.reportTrash?.[reportRow.id] || []}
          schoolName={schoolName}
          principalName={principalName}
          countryName={countryName}
          ministryName={ministryName}
          logoImage={logoImage}
          onClose={() => setReportRowId(null)}
          onEditEntry={(entryId, value) => updateReportEntryValue(reportRow.id, entryId, value)}
          onDeleteEntry={(entryId) => removeReportEntry(reportRow.id, entryId)}
          onDeleteCategory={(colId) => removeReportCategory(reportRow.id, colId)}
          onDeleteAllEntries={() => removeAllReportEntries(reportRow.id)}
          onAddNote={(entry) => addManualReportNote(reportRow.id, entry)}
          onRestoreLatest={() => restoreLatestReportEntry(reportRow.id)}
          onRestoreEntry={(trashId) => restoreReportEntryFromTrash(reportRow.id, trashId)}
          onClearTrash={() => clearReportTrash(reportRow.id)}
          onPrint={() => setPrintChoice({ type: "report", cls, row: reportRow, entries: reportEntries })}
        />
      )}
      {confirmAction && (
        <ConfirmDialog
          title={
            confirmAction.type === "deleteAll" ? "حذف كل بيانات الفصل"
            : confirmAction.type === "deleteColumn" ? "حذف العمود"
            : "حذف الصف"
          }
          message={
            confirmAction.type === "deleteAll"
              ? "سيتم حذف جميع الأعمدة والصفوف وبيانات الرصد في هذا الفصل. متابعة؟"
              : confirmAction.type === "deleteColumn"
              ? `سيتم حذف العمود "${confirmAction.column.name}" وكل ما رُصد فيه. يمكن استعادته لاحقًا من سلة المحذوفات. متابعة؟`
              : `سيتم حذف الصف "${confirmAction.row.name}" بكل بياناته نهائيًا. يمكن استعادته لاحقًا من سلة المحذوفات. متابعة؟`
          }
          onCancel={() => setConfirmAction(null)}
          onConfirm={runConfirmedAction}
        />
      )}
      {showTrash && (
        <TrashModal
          trash={cls.trash}
          onClose={() => setShowTrash(false)}
          onRestore={(id) => restoreEntry(id)}
          onClearAll={() => updateClass((c) => ({ ...c, trash: [] }))}
        />
      )}
      {printChoice && (
        <PrintFormatModal onClose={() => setPrintChoice(null)} onChoose={handleChooseFormat} />
      )}
      {showAttendance && (
        <AttendanceModal
          cls={cls}
          updateClass={updateClass}
          onClose={() => setShowAttendance(false)}
          onPrint={(dateKey) => setPrintChoice({ type: "attendance", cls, dateKey })}
          onShare={(dateKey) => exportPdfShare({ type: "attendance", cls, dateKey })}
        />
      )}
      {showEvents && (
        <EventsModal
          events={cls.events}
          speed={cls.tickerSpeed || 14}
          onChangeSpeed={(s) => updateClass((c) => ({ ...c, tickerSpeed: s }))}
          onClose={() => setShowEvents(false)}
          onAdd={addEvent}
          onUpdate={updateEvent}
          onDelete={deleteEvent}
          onMove={moveEvent}
        />
      )}
      {showRandomPicker && <RandomPickerModal rows={cls.rows} onClose={() => setShowRandomPicker(false)} />}
      {showRandomGroups && <RandomGroupsModal rows={cls.rows} onClose={() => setShowRandomGroups(false)} />}
      {showFilterModal && (
        <FilterModal
          columns={cls.columns}
          initial={activeFilter}
          onApply={setActiveFilter}
          onClear={() => setActiveFilter(null)}
          onClose={() => setShowFilterModal(false)}
        />
      )}
      {showBulkRecordModal && (
        <BulkRecordModal
          columns={cls.columns}
          onApply={(col, val) => bulkSetColumnValueForSelected(col, val)}
          onClose={() => setShowBulkRecordModal(false)}
        />
      )}
      {showBulkCertificateModal && (
        <BulkCertificateModal
          cls={cls}
          rows={cls.rows.filter((r) => selectedRowIds.includes(r.id))}
          schoolName={schoolName}
          principalName={principalName}
          countryName={countryName}
          ministryName={ministryName}
          logoImage={logoImage}
          onClose={() => setShowBulkCertificateModal(false)}
        />
      )}
      {showMoveModal && (
        <MoveStudentsModal
          currentClassId={cls.id}
          allClasses={allClasses}
          count={selectedRowIds.length}
          onMove={(destId, includeGrades) => {
            onMoveRowsToClass(cls.id, destId, selectedRowIds, includeGrades);
            clearSelection();
          }}
          onClose={() => setShowMoveModal(false)}
        />
      )}
      {confirmBulkDelete && (
        <ConfirmDialog
          title="حذف الطلاب المحددين"
          message={`سيتم حذف ${selectedRowIds.length} طالب وكل بياناتهم. يمكن استعادتهم لاحقًا من سلة المحذوفات.`}
          onCancel={() => setConfirmBulkDelete(false)}
          onConfirm={() => { bulkDeleteSelected(); setConfirmBulkDelete(false); }}
        />
      )}
    </div>
  );
}

function AuthScreen() {
  const [mode, setMode] = useState("login"); // 'login' | 'signup'
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  const submit = async () => {
    setError(""); setMessage(""); setLoading(true);
    try {
      if (mode === "signup") {
        const { data, error } = await supabase.auth.signUp({ email: email.trim(), password });
        if (error) throw error;
        if (!data.session) {
          setMessage("تم إنشاء الحساب! تحقق من بريدك الإلكتروني لتفعيله، ثم سجّل دخولك.");
          setMode("login");
        }
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
        if (error) throw error;
      }
      // نجاح الدخول يُحدَّث تلقائيًا عبر مستمع onAuthStateChange في App
    } catch (e) {
      setError(e.message === "Invalid login credentials" ? "البريد الإلكتروني أو كلمة المرور غير صحيحة." : e.message);
    } finally {
      setLoading(false);
    }
  };

  const signInWithGoogle = async () => {
    setError("");
    const { error } = await supabase.auth.signInWithOAuth({
      provider: "google",
      options: { redirectTo: window.location.origin },
    });
    if (error) setError(error.message);
    // عند النجاح، يعيد Google توجيهك تلقائيًا لهذه الصفحة وجلستك تُفعَّل من تلقاء نفسها
  };

  return (
    <div className="min-h-screen flex items-center justify-center px-4" style={{ background: PAPER, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }} dir="rtl">
      <div className="w-full max-w-sm rounded-2xl p-6 modal-panel-in" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
        <h1 className="text-xl font-extrabold text-center mb-1" style={{ color: INK, fontFamily: "'Cairo', sans-serif" }}>دفتر المتابعة</h1>
        <p className="text-sm text-center mb-6" style={{ color: MUTED }}>{mode === "login" ? "سجّل دخولك للمتابعة" : "أنشئ حسابًا جديدًا"}</p>

        <button onClick={signInWithGoogle} className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold mb-4 active:scale-95 transition-transform" style={{ border: `1px solid ${LINE}`, color: INK }}>
          <svg width="18" height="18" viewBox="0 0 48 48"><path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3c-1.6 4.7-6.1 8-11.3 8-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.1 8 3l5.7-5.7C34.6 6 29.6 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.7-.4-3.5z"/><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.6 16 18.9 13 24 13c3.1 0 5.8 1.1 8 3l5.7-5.7C34.6 6 29.6 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.5 0 10.4-1.9 14.3-5.1l-6.6-5.6C29.6 35 26.9 36 24 36c-5.2 0-9.6-3.3-11.3-7.9l-6.5 5C9.6 39.6 16.3 44 24 44z"/><path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.3-2.2 4.3-4.1 5.7l6.6 5.6C41.4 36 44 30.6 44 24c0-1.3-.1-2.7-.4-3.5z"/></svg>
          الدخول عبر Google
        </button>
        <div className="flex items-center gap-2 mb-4">
          <div className="flex-1 h-px" style={{ background: LINE }} />
          <span className="text-xs" style={{ color: MUTED }}>أو</span>
          <div className="flex-1 h-px" style={{ background: LINE }} />
        </div>

        <Field label="البريد الإلكتروني">
          <input type="email" style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@example.com" autoComplete="email" />
        </Field>
        <Field label="كلمة المرور">
          <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" autoComplete={mode === "login" ? "current-password" : "new-password"} onKeyDown={(e) => { if (e.key === "Enter") submit(); }} />
        </Field>

        {error && <p className="text-xs mb-3" style={{ color: "#C0392B" }}>{error}</p>}
        {message && <p className="text-xs mb-3" style={{ color: "#0F6B5C" }}>{message}</p>}

        <button disabled={loading || !email.trim() || !password} onClick={submit} className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 active:scale-95 transition-transform" style={{ background: "#0F6B5C" }}>
          {loading ? "جارٍ..." : mode === "login" ? "تسجيل الدخول" : "إنشاء حساب"}
        </button>

        <button onClick={() => { setMode((m) => (m === "login" ? "signup" : "login")); setError(""); setMessage(""); }} className="w-full text-center text-xs font-semibold mt-4" style={{ color: "#0F6B5C" }}>
          {mode === "login" ? "ما عندك حساب؟ أنشئ واحدًا" : "عندك حساب؟ سجّل دخولك"}
        </button>
      </div>
    </div>
  );
}

// ---------- App root ----------

export default function App() {
  useFonts();
  const [session, setSession] = useState(undefined); // undefined = still checking, null = logged out
  const [data, setData] = useState({ classes: [], trash: [], schedule: {}, scheduleImage: null, settings: { feedback: true } });
  const [loaded, setLoaded] = useState(false);
  const [view, setView] = useState({ page: "home" });
  const [printJob, requestPrint] = usePrint();

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => setSession(session));
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) return;
    setLoaded(false);
    (async () => {
      try {
        const { data: row, error } = await supabase.from("user_data").select("data").eq("user_id", session.user.id).maybeSingle();
        if (error) throw error;
        if (row && row.data) setData(row.data);
      } catch (e) {
        console.error("تعذر تحميل البيانات", e);
      } finally {
        setLoaded(true);
      }
    })();
  }, [session]);

  useEffect(() => {
    if (!loaded || !session) return;
    const t = setTimeout(() => {
      supabase.from("user_data")
        .upsert({ user_id: session.user.id, data, updated_at: new Date().toISOString() }, { onConflict: "user_id" })
        .then(({ error }) => { if (error) console.error("تعذر حفظ البيانات", error); });
    }, 600);
    return () => clearTimeout(t);
  }, [data, loaded, session]);

  const openClass = (id) => setView({ page: "class", id });
  const backHome = () => setView({ page: "home" });
  const updateClass = (fn) => setData((d) => ({ ...d, classes: d.classes.map((c) => (c.id === view.id ? fn(c) : c)) }));
  const currentClass = view.page === "class" ? data.classes.find((c) => c.id === view.id) : null;

  // Moves selected students from the currently open class to another class of
  // theirs. Columns are matched by name, so recorded values only transfer for
  // columns that exist (by name) in both classes; anything else is skipped.
  const moveRowsToClass = (sourceClassId, destClassId, rowIds, includeGrades) => {
    setData((d) => {
      const source = d.classes.find((c) => c.id === sourceClassId);
      const dest = d.classes.find((c) => c.id === destClassId);
      if (!source || !dest) return d;
      const movingRows = source.rows.filter((r) => rowIds.includes(r.id));
      const newRows = [];
      const newCellsForDest = {};
      const newReportsForDest = {};
      movingRows.forEach((row) => {
        const newId = uid();
        newRows.push({ ...row, id: newId });
        if (includeGrades) {
          source.columns.forEach((srcCol) => {
            const destCol = dest.columns.find((dc) => dc.name === srcCol.name);
            if (!destCol) return;
            const val = source.cells[`${row.id}:${srcCol.id}`];
            if (val) newCellsForDest[`${newId}:${destCol.id}`] = val;
          });
          const srcReports = source.reports?.[row.id] || [];
          const remapped = srcReports
            .map((e) => {
              const srcCol = source.columns.find((c) => c.id === e.colId);
              const destCol = srcCol ? dest.columns.find((dc) => dc.name === srcCol.name) : null;
              if (!destCol) return null;
              return { ...e, id: uid(), colId: destCol.id, colName: destCol.name, colColor: destCol.color };
            })
            .filter(Boolean);
          if (remapped.length > 0) newReportsForDest[newId] = remapped;
        }
      });
      const updatedDest = {
        ...dest,
        rows: [...dest.rows, ...newRows],
        cells: { ...dest.cells, ...newCellsForDest },
        reports: { ...(dest.reports || {}), ...newReportsForDest },
      };
      const remainingReports = { ...(source.reports || {}) };
      rowIds.forEach((id) => delete remainingReports[id]);
      const updatedSource = {
        ...source,
        rows: source.rows.filter((r) => !rowIds.includes(r.id)),
        cells: Object.fromEntries(Object.entries(source.cells).filter(([k]) => !rowIds.some((id) => k.startsWith(`${id}:`)))),
        reports: remainingReports,
      };
      return { ...d, classes: d.classes.map((c) => (c.id === destClassId ? updatedDest : c.id === sourceClassId ? updatedSource : c)) };
    });
  };

  if (session === undefined) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: PAPER }}>
        <p style={{ color: MUTED, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }}>...جارٍ التحميل</p>
      </div>
    );
  }

  if (!session) {
    return <AuthScreen />;
  }

  if (!loaded) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: PAPER }}>
        <p style={{ color: MUTED, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }}>...جارٍ التحميل</p>
      </div>
    );
  }

  const darkMode = !!data.settings?.darkMode;
  const appContent = (
    <>
      <PrintStyles />
      {view.page === "home" && <HomePage data={data} setData={setData} onOpen={openClass} userEmail={session.user.email} onSignOut={() => supabase.auth.signOut()} />}
      {view.page === "class" && currentClass && <ClassPage cls={currentClass} updateClass={updateClass} onBack={backHome} requestPrint={requestPrint} feedbackEnabled={data.settings?.feedback !== false} schoolName={data.settings?.schoolName} principalName={data.settings?.principalName} countryName={data.settings?.countryName} ministryName={data.settings?.ministryName} logoImage={data.settings?.logoImage} allClasses={data.classes} onMoveRowsToClass={moveRowsToClass} />}
      {view.page === "class" && !currentClass && (
        <div className="max-w-md mx-auto py-20 text-center">
          <p style={{ color: MUTED }}>لم يتم العثور على هذا الفصل</p>
          <button onClick={backHome} className="mt-3 text-sm font-semibold" style={{ color: "#0F6B5C" }}>العودة للرئيسية</button>
        </div>
      )}
      <div className="app-print-root">
        <PrintContent job={printJob} />
      </div>
    </>
  );

  if (darkMode) {
    // filter establishes a new containing block for position:fixed descendants
    // (all our modals), so this wrapper is made fixed/inset-0 itself — that
    // keeps its box identical to the viewport, so fixed modals still render
    // full-screen correctly instead of being clipped to a scrolled document.
    return (
      <div
        dir="rtl"
        style={{
          position: "fixed", inset: 0, overflowY: "auto", WebkitOverflowScrolling: "touch",
          background: PAPER, fontFamily: "'IBM Plex Sans Arabic', sans-serif",
          filter: "invert(1) hue-rotate(180deg)",
        }}
      >
        <style>{`.dark-mode-img-fix { filter: invert(1) hue-rotate(180deg); }`}</style>
        {appContent}
      </div>
    );
  }

  return (
    <div dir="rtl" className="min-h-screen" style={{ background: PAPER, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }}>
      {appContent}
    </div>
  );
}
