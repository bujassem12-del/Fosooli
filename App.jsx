import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import { LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from "recharts";
import {
  Plus, Pencil, Palette, Trash2, Archive, ArchiveRestore, LayoutGrid,
  Printer, Search, ArrowRight, X, Check, Minus, Hash, Type,
  ListChecks, FolderClock, BookOpen, FileText, RefreshCw, ClipboardList,
  Pin, PinOff, Copy, RotateCcw, FolderOpen, FileImage, FileSpreadsheet, ListOrdered,
  Share2, Calendar, CalendarCheck, Newspaper, Eraser, CalendarRange, UserX, Paperclip, Link2,
  Lock, Unlock, ChevronUp, ChevronDown, ChevronLeft, ChevronRight, ImageDown, FileOutput,
  Camera, ImageOff, Settings, Volume2, VolumeX, BarChart3, Users,
  Shuffle, AlertTriangle, MessageSquareWarning, ClipboardCopy, Eye, EyeOff, Award, Download, Target, BookMarked, WifiOff, QrCode, Layers, Gamepad2,
  CalendarPlus, Moon, Sun, Filter, ListTodo, HelpCircle, Send, Activity, Info, ShieldCheck, Pipette, Bell, Move, User, ListPlus, LogOut, MoreHorizontal, Home, MoreVertical, Sparkles, ExternalLink, FileCheck
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
  const preset = COLORS.find((c) => c.hex === hex);
  if (preset) return preset.light;
  if (!hex || hex[0] !== "#" || hex.length !== 7) return "#F3F1E9";
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return "#F3F1E9";
  const mix = (c) => Math.round(c + (255 - c) * 0.82);
  return `rgb(${mix(r)}, ${mix(g)}, ${mix(b)})`;
}

// يستخدم لتلوين خانة عمود "عداد" تدريجيًا حسب قيمتها الحالية نسبةً للحد
// الأقصى — عند ratio=0 يرجع تدرّجًا فاتحًا جدًا من اللون، وعند ratio=1
// يرجع اللون كاملًا بلا تخفيف (تدرّج حراري بسيط).
function intensityColor(hex, ratio) {
  if (!hex || hex[0] !== "#" || hex.length !== 7) return "#F3F1E9";
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return "#F3F1E9";
  const clamped = Math.max(0, Math.min(1, Number(ratio) || 0));
  const mixFactor = 0.85 - clamped * 0.85;
  const mix = (c) => Math.round(c + (255 - c) * mixFactor);
  return `rgb(${mix(r)}, ${mix(g)}, ${mix(b)})`;
}

const CLASS_EMOJIS = ["📐", "📖", "🔬", "🎨", "⚽", "🕌", "🌍", "💻", "✏️", "📊", "🎵", "🧮", "🔤", "🌱", "⚗️", "📚"];

const INK = "#232622";
const PAPER = "#FAF8F3";
const LINE = "#E4DFD2";
const MUTED = "#7A7768";

// Dashboard-style theme accents (sidebar, stat cards, shawahed report banners)
// DASH_GREEN/DASH_GREEN_DARK هنا "let" وليس "const" عمدًا — تسمح بتخصيص لون
// هوية شخصي لكل معلم (راجع applyThemeColor) يُطبَّق تلقائيًا بكل مكان
// بالتطبيق يستخدم هذين المتغيرين، دون الحاجة نمرّر اللون كـ prop بمئات
// الأماكن.
let DASH_GREEN = "#26423B";
let DASH_GREEN_DARK = "#1B322C";
const GOLD = "#D9A441";
const GOLD_LIGHT = "#F7EBD2";

// يغمّق لونًا سداسيًا بنسبة معيّنة (0-1) — يُستخدم لتوليد الدرجة الغامقة
// المرافقة لأي لون هوية شخصي يختاره المعلم.
function darkenHex(hex, amount = 0.3) {
  if (!hex || hex[0] !== "#" || hex.length !== 7) return hex;
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return hex;
  const dim = (c) => Math.max(0, Math.round(c * (1 - amount)));
  const toHex = (c) => c.toString(16).padStart(2, "0");
  return `#${toHex(dim(r))}${toHex(dim(g))}${toHex(dim(b))}`;
}

// يطبّق لون الهوية الشخصية للمعلم على كامل التطبيق — يستدعى مرة كل ما
// تتغيّر إعدادات المستخدم (راجع useEffect بمكوّن App).
function applyThemeColor(hex) {
  const safe = hex && /^#[0-9A-Fa-f]{6}$/.test(hex) ? hex : "#26423B";
  DASH_GREEN = safe;
  DASH_GREEN_DARK = darkenHex(safe, 0.32);
}

// يحوّل لونًا سداسيًا لصيغة rgba بشفافية محددة — يُستخدم لمد لون الهوية
// الشخصية على زخارف خلفية الموقع وحوافه، مو بس الأزرار.
function hexToRgba(hex, alpha = 1) {
  if (!hex || hex[0] !== "#" || hex.length !== 7) return `rgba(38,66,59,${alpha})`;
  const r = parseInt(hex.slice(1, 3), 16), g = parseInt(hex.slice(3, 5), 16), b = parseInt(hex.slice(5, 7), 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) return `rgba(38,66,59,${alpha})`;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// يكتشف عمود الأسماء تلقائيًا من ملف Excel بغض النظر عن ترتيب الأعمدة —
// يهم خصوصًا لملفات نظام نور التي عادةً تضع الرقم التسلسلي أو رقم الهوية
// بالعمود الأول قبل عمود الاسم، بدل الافتراض إن الاسم دائمًا بالعمود الأول.
const NAME_HEADER_WORDS = ["اسم", "الاسم", "اسم الطالب", "الطالب", "name", "student", "students"];
const NON_NAME_HEADER_WORDS = ["م", "الرقم", "رقم", "الرقم التسلسلي", "رقم الهوية", "الهوية", "تاريخ الميلاد", "الصف", "الفصل", "الجنس", "رقم الجلوس", "الحالة", "no", "id", "number", "class", "grade", "gender", "date"];
function looksLikeName(val) {
  const s = String(val ?? "").trim();
  if (!s || s.length < 2 || s.length > 60) return false;
  if (/^[\d\s\-\/.:]+$/.test(s)) return false; // أرقام/تواريخ/هويات بحتة
  if (NON_NAME_HEADER_WORDS.includes(s.toLowerCase())) return false;
  return /[\u0600-\u06FFa-zA-Z]/.test(s); // يحتوي حروفًا عربية أو إنجليزية فعلية
}
function extractNamesSmart(rows) {
  if (!rows || rows.length === 0) return [];
  const colCount = Math.max(...rows.map((r) => r.length), 0);
  let bestCol = 0, bestScore = -1;
  for (let c = 0; c < Math.min(colCount, 10); c++) {
    let score = 0;
    rows.forEach((r) => { if (looksLikeName(r[c])) score++; });
    if (score > bestScore) { bestScore = score; bestCol = c; }
  }
  if (bestScore <= 0) return [];
  const names = [];
  rows.forEach((r) => {
    const val = String(r[bestCol] ?? "").trim();
    if (!val) return;
    if (NAME_HEADER_WORDS.includes(val.toLowerCase())) return;
    if (!looksLikeName(val)) return;
    names.push(val);
  });
  return names;
}

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
// نسبة الحضور: من بين الأيام اللي فعلاً سُجّل فيها حضور لهذا الفصل (يوم واحد
// على الأقل انحط له تحديد غياب)، كم يوم كان الطالب حاضر فيه. الأيام اللي ما
// انسجل فيها أي حضور أصلاً (ما ضغط المعلم "غياب" لأي طالب) ما تُحتسب، لأننا
// ما نملك تأكيد إنها كانت يوم دراسي فعلي بالفصل.
// الدرجة الكلية: مجموع كل أعمدة "عداد" اللي حُدّدت لها درجة قصوى، من أصل
// مجموع تلك الدرجات القصوى. الأعمدة اللي ما فيها درجة قصوى محددة ما تدخل
// بالحساب أصلًا (نعتبرها أعمدة تتبّع عامة مو جزء من الدرجة النهائية).
function totalGrade(cls, rowId) {
  const gradeCols = cls.columns.filter((c) => c.type === "counter" && Number(c.maxValue) > 0);
  if (gradeCols.length === 0) return null;
  let score = 0, max = 0;
  gradeCols.forEach((c) => {
    score += Number(cls.cells[`${rowId}:${c.id}`]) || 0;
    max += Number(c.maxValue);
  });
  const pct = max > 0 ? (score / max) * 100 : 0;
  let band, bandColor, bandBg;
  if (pct >= 90) { band = "ممتاز"; bandColor = "#26423B"; bandBg = "#E3F1EC"; }
  else if (pct >= 80) { band = "جيد جدًا"; bandColor = "#2E7DA6"; bandBg = "#E3EEF5"; }
  else if (pct >= 65) { band = "جيد"; bandColor = "#C97A2B"; bandBg = "#FBEEE0"; }
  else { band = "ضعيف"; bandColor = "#C0392B"; bandBg = "#FBEAE7"; }
  return { score: Math.round(score * 100) / 100, max, pct: Math.round(pct), band, bandColor, bandBg };
}

function attendancePercent(cls, rowId) {
  const days = Object.keys(cls.attendance || {});
  if (days.length === 0) return null;
  const absentDays = days.filter((d) => cls.attendance[d]?.[rowId] === "absent").length;
  return Math.round(((days.length - absentDays) / days.length) * 100);
}

function attendanceStatus(cls, rowId, dateKey) {
  return cls.attendance?.[dateKey]?.[rowId] === "absent" ? "absent" : "present";
}
// "وضع الاختبار": يقفل أزرار التعديل/الحذف الهيكلية للفصل مؤقتًا (منع لمسة
// خطأ أثناء التجول بالفصل)، بينما رصد الدرجات والغياب يبقى شغّالًا طبيعيًا.
function isExamModeActive(cls) {
  return !!(cls.examModeUntil && new Date(cls.examModeUntil).getTime() > Date.now());
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
function wrapCanvasText(ctx, text, maxWidth, emptyFallback = "—") {
  const raw = String(text ?? "").trim();
  if (!raw) return [emptyFallback];
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
function buildTableCanvas({ title, subtitle, headers, rows, blankTemplate = false }) {
  const cellW = 150, lineH = 18, cellPadV = 12, headerH = 42, pad = 24, titleH = 70;
  const cols = headers.length;
  const width = pad * 2 + cols * cellW;
  const scale = 3;
  const measure = document.createElement("canvas").getContext("2d");
  measure.font = "13px Tahoma, Arial";
  const emptyPlaceholder = blankTemplate ? "" : "—";
  const rowLines = rows.map((r) => r.map((val) => wrapCanvasText(measure, val ? String(val) : emptyPlaceholder, cellW - 16, emptyPlaceholder)));
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
  // canvas coordinates are always left-to-right regardless of ctx.direction
  // (that only affects text shaping) — so for RTL tables we must explicitly
  // position column 0 at the RIGHT edge and lay subsequent columns leftward.
  const colX = (i) => pad + (cols - 1 - i) * cellW;
  let y = titleH;
  ctx.fillStyle = "#F3F1E9";
  ctx.fillRect(pad, y, cols * cellW, headerH);
  headers.forEach((h, i) => {
    ctx.strokeStyle = LINE;
    ctx.strokeRect(colX(i), y, cellW, headerH);
    ctx.fillStyle = INK;
    ctx.font = "bold 14px Tahoma, Arial";
    ctx.fillText(String(h), colX(i) + cellW / 2, y + headerH / 2);
  });
  y += headerH;
  rows.forEach((r, ri) => {
    const rh = rowHeights[ri];
    ctx.fillStyle = ri % 2 ? "#FBFAF6" : "#ffffff";
    ctx.fillRect(pad, y, cols * cellW, rh);
    r.forEach((val, ci) => {
      ctx.strokeStyle = LINE;
      ctx.strokeRect(colX(ci), y, cellW, rh);
      ctx.fillStyle = INK;
      ctx.font = "13px Tahoma, Arial";
      const lines = rowLines[ri][ci];
      let ly = y + rh / 2 - (lines.length * lineH) / 2 + lineH / 2;
      lines.forEach((ln) => { ctx.fillText(ln, colX(ci) + cellW / 2, ly); ly += lineH; });
    });
    y += rh;
  });
  return { canvas, logicalWidth: width, logicalHeight: height };
}

// Builds a well-organized, sectioned canvas for a student report: one
// colored section per category (behavior, homework, participation, notes,
// exams...) each with its own mini date/time/value table — legible enough
// to print and hand to a parent.
function buildReportCanvas({ title, subtitle, groups, photoImageElement }) {
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
  if (photoImageElement) {
    const r = 26;
    ctx.save();
    ctx.beginPath();
    ctx.arc(width - pad - r, pad + r, r, 0, Math.PI * 2);
    ctx.closePath();
    ctx.clip();
    ctx.drawImage(photoImageElement, width - pad - r * 2, pad, r * 2, r * 2);
    ctx.restore();
  }
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
    ctx.fillStyle = g.colColor || "#26423B";
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

// أصغر قيمة أعلى تكرارًا بمصفوفة قيم (لأعمدة القوائم المنسدلة) — تُستخدم
// بتقرير ولي الأمر المبسّط لاختيار "الحالة الأكثر تكرارًا" بدل عرض كل سجل.
function mostFrequentValue(values) {
  const counts = {};
  values.forEach((v) => { counts[v] = (counts[v] || 0) + 1; });
  let best = null, bestCount = 0;
  Object.entries(counts).forEach(([k, c]) => { if (c > bestCount) { best = k; bestCount = c; } });
  return { value: best, count: bestCount };
}

// يرسم نجوم تقييم (ممتلئة/فارغة) بدل جدول أرقام — أسهل بكثير لولي أمر
// يقرأها بسرعة من عرض مباشر.
function drawStars(ctx, cx, cy, filled, total, size, color) {
  const gap = size * 1.3;
  const startX = cx - ((total - 1) * gap) / 2;
  for (let i = 0; i < total; i++) {
    const x = startX + i * gap;
    const isFilled = i < filled;
    ctx.save();
    ctx.translate(x, cy);
    ctx.beginPath();
    for (let p = 0; p < 5; p++) {
      const angle = (p * 4 * Math.PI) / 5 - Math.PI / 2;
      const px = Math.cos(angle) * size, py = Math.sin(angle) * size;
      if (p === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath();
    if (isFilled) { ctx.fillStyle = color; ctx.fill(); }
    else { ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.stroke(); }
    ctx.restore();
  }
}

// تقرير ولي الأمر المبسّط: بطاقات وأيقونات ونجوم بدل جداول تفصيلية — مصمَّم
// ليُفهَم بنظرة سريعة بدون خلفية تربوية، مع إبقاء تقرير المعلم التفصيلي
// (buildReportCanvas) كما هو للاستخدام الداخلي.
async function buildParentReportCanvas({ cls, row, entries, meta = {} }) {
  const { schoolName, teacherName, date } = meta;
  const width = 760;
  const scale = 3;
  const pad = 32;
  const groups = groupEntries(entries);

  const cardW = (width - pad * 2 - 16) / 2;
  const cardH = 92;
  const cardsPerRow = 2;
  const cardRows = Math.ceil(groups.length / cardsPerRow);

  const attPct = attendancePercent(cls, row.id);
  const grade = totalGrade(cls, row.id);

  const headerH = 150;
  const statsH = attPct !== null || grade ? 110 : 0;
  const cardsH = groups.length > 0 ? cardRows * (cardH + 14) + 40 : 60;
  const footerH = 90;
  const height = headerH + statsH + cardsH + footerH;

  let photoImageElement = null;
  if (row.photo) { try { photoImageElement = await loadImage(row.photo); } catch (e) { photoImageElement = null; } }

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = PAPER;
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textBaseline = "middle";

  // ---- header ----
  const grad = ctx.createLinearGradient(0, 0, width, 0);
  grad.addColorStop(0, DASH_GREEN);
  grad.addColorStop(1, DASH_GREEN_DARK);
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, width, headerH);

  if (photoImageElement) {
    const r = 34;
    ctx.save();
    ctx.beginPath();
    ctx.arc(width - pad - r, headerH / 2, r, 0, Math.PI * 2);
    ctx.closePath();
    ctx.clip();
    ctx.drawImage(photoImageElement, width - pad - r * 2, headerH / 2 - r, r * 2, r * 2);
    ctx.restore();
    ctx.strokeStyle = "#fff";
    ctx.lineWidth = 3;
    ctx.beginPath();
    ctx.arc(width - pad - r, headerH / 2, r, 0, Math.PI * 2);
    ctx.stroke();
  }

  ctx.textAlign = "center";
  ctx.fillStyle = GOLD_LIGHT;
  ctx.font = "13px Tahoma, Arial";
  ctx.fillText("📋 تقرير مبسّط لولي الأمر", width / 2, 34);
  ctx.fillStyle = "#fff";
  ctx.font = "bold 26px Tahoma, Arial";
  ctx.fillText(row.name, width / 2, 68);
  ctx.font = "13px Tahoma, Arial";
  ctx.fillStyle = "rgba(255,255,255,0.85)";
  const sub = [cls.subject, cls.grade, schoolName].filter(Boolean).join("  •  ");
  ctx.fillText(sub, width / 2, 96);
  ctx.font = "12px Tahoma, Arial";
  ctx.fillStyle = GOLD_LIGHT;
  ctx.fillText(date || formatDateDisplay(todayKey()), width / 2, 122);

  let y = headerH + 20;

  // ---- attendance + overall grade badges ----
  if (attPct !== null || grade) {
    const badges = [];
    if (attPct !== null) {
      const face = attPct >= 90 ? "😊" : attPct >= 75 ? "🙂" : "😟";
      const color = attPct >= 90 ? "#0F9D58" : attPct >= 75 ? "#C97A2B" : "#C0392B";
      badges.push({ icon: face, label: "نسبة الحضور", value: `${attPct}٪`, color });
    }
    if (grade) {
      const face = grade.pct >= 90 ? "🏆" : grade.pct >= 65 ? "⭐" : "📈";
      badges.push({ icon: face, label: "التقييم العام", value: grade.band, color: grade.bandColor });
    }
    const bW = (width - pad * 2 - 16) / badges.length;
    badges.forEach((b, i) => {
      const x = pad + i * (bW + 16);
      ctx.fillStyle = `${b.color}14`;
      ctx.strokeStyle = b.color;
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.roundRect(x, y, bW, 90, 16);
      ctx.fill();
      ctx.stroke();
      ctx.textAlign = "center";
      ctx.font = "28px Tahoma, Arial";
      ctx.fillText(b.icon, x + bW / 2, y + 32);
      ctx.font = "bold 20px Tahoma, Arial";
      ctx.fillStyle = b.color;
      ctx.fillText(b.value, x + bW / 2, y + 58);
      ctx.font = "12px Tahoma, Arial";
      ctx.fillStyle = MUTED;
      ctx.fillText(b.label, x + bW / 2, y + 78);
    });
    y += 90 + 20;
  }

  // ---- category cards ----
  if (groups.length === 0) {
    ctx.textAlign = "center";
    ctx.font = "14px Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText("لا يوجد رصد لهذا الطالب بعد.", width / 2, y + 20);
  } else {
    groups.forEach((g, gi) => {
      const col = cls.columns.find((c) => c.id === g.colId);
      const colIdx = gi % cardsPerRow;
      const rowIdx = Math.floor(gi / cardsPerRow);
      const x = pad + colIdx * (cardW + 16);
      const cy = y + rowIdx * (cardH + 14);

      ctx.fillStyle = "#fff";
      ctx.strokeStyle = LINE;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.roundRect(x, cy, cardW, cardH, 14);
      ctx.fill();
      ctx.stroke();

      // colored icon chip
      const chipR = 20;
      ctx.beginPath();
      ctx.arc(x + 34, cy + cardH / 2, chipR, 0, Math.PI * 2);
      ctx.fillStyle = g.colColor || "#0F6B5C";
      ctx.fill();
      ctx.fillStyle = "#fff";
      ctx.textAlign = "center";
      ctx.font = "bold 16px Tahoma, Arial";
      ctx.fillText((g.colName || "؟").trim().charAt(0), x + 34, cy + cardH / 2);

      ctx.textAlign = "right";
      ctx.fillStyle = INK;
      ctx.font = "bold 14px Tahoma, Arial";
      ctx.fillText(g.colName, x + cardW - 16, cy + 26);

      if (col && col.type === "counter") {
        const nums = g.items.map((it) => Number(it.value)).filter((n) => !Number.isNaN(n));
        const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        const max = Number(col.maxValue) || 5;
        const starCount = Math.max(0, Math.min(5, Math.round((avg / max) * 5)));
        drawStars(ctx, x + cardW - 60, cy + 58, starCount, 5, 7, g.colColor || "#0F6B5C");
        ctx.textAlign = "left";
        ctx.font = "11px Tahoma, Arial";
        ctx.fillStyle = MUTED;
        ctx.fillText(`المعدل: ${Math.round(avg * 10) / 10}`, x + 16, cy + 58);
      } else if (col && col.type === "dropdown") {
        const { value, count } = mostFrequentValue(g.items.map((it) => it.value));
        const opt = (col.options || []).find((o) => o.label === value);
        ctx.textAlign = "right";
        ctx.font = "bold 13px Tahoma, Arial";
        ctx.fillStyle = opt?.color || g.colColor || INK;
        ctx.fillText(value || "—", x + cardW - 16, cy + 56);
        ctx.textAlign = "left";
        ctx.font = "11px Tahoma, Arial";
        ctx.fillStyle = MUTED;
        ctx.fillText(`${count} من ${g.items.length}`, x + 16, cy + 56);
      } else {
        ctx.textAlign = "right";
        ctx.font = "bold 13px Tahoma, Arial";
        ctx.fillStyle = INK;
        ctx.fillText(`${g.items.length} ملاحظة مسجّلة`, x + cardW - 16, cy + 56);
      }
    });
    y += cardRows * (cardH + 14) + 10;
  }

  // ---- footer ----
  ctx.strokeStyle = LINE;
  ctx.beginPath();
  ctx.moveTo(pad, height - footerH + 10);
  ctx.lineTo(width - pad, height - footerH + 10);
  ctx.stroke();
  ctx.textAlign = "center";
  ctx.font = "13px Tahoma, Arial";
  ctx.fillStyle = DASH_GREEN;
  ctx.fillText("🌟 استمروا بدعم ابنكم/ابنتكم في المنزل، ونرحّب بتواصلكم دائمًا", width / 2, height - footerH + 40);
  if (teacherName) {
    ctx.font = "12px Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText(`المعلم: ${teacherName}`, width / 2, height - footerH + 64);
  }

  return { canvas, logicalWidth: width, logicalHeight: height };
}

// Draws a simple decorative certificate of appreciation onto a canvas.
// تقرير "برنامج" رسمي مفصّل — يطابق نموذج الإدارات التعليمية الشائع
// (اسم البرنامج، مجاله، المنفّذون، الأهداف، خطوات التنفيذ، والشواهد بصور).
// يُستخدم لما يكون للشاهد بيانات "program" مفصّلة (وليس مجرد صورة وملاحظة).
async function buildProgramReportCanvas({ entry, cat, meta = {} }) {
  const { countryName, ministryName, region, office, schoolName, logoImage } = meta;
  const p = entry.program || {};
  const width = 800;
  const pad = 32;
  const scale = 3;

  let logoImageElement = null;
  if (logoImage) { try { logoImageElement = await loadImage(logoImage); } catch (e) { logoImageElement = null; } }

  const photos = (p.photos && p.photos.length ? p.photos : (entry.photo ? [entry.photo] : [])).slice(0, 4);
  const photoImgs = [];
  for (const src of photos) {
    try { photoImgs.push(await loadImage(src)); } catch (e) { /* skip */ }
  }

  const measure = document.createElement("canvas").getContext("2d");
  measure.font = "12px 'IBM Plex Sans Arabic', Tahoma, Arial";

  // ---- sizing ----
  const headerH = 90;
  const bannerH = 44;
  const gridGap = 10;
  const boxH = 64;
  const infoGridH = boxH * 3 + gridGap * 2 + 20;

  const objectives = (p.objectives && p.objectives.length ? p.objectives : ["—"]);
  const steps = (p.steps && p.steps.length ? p.steps : ["—"]);
  const listBoxPad = 14;
  const lineH = 19;
  const objLines = objectives.length * lineH + 40;
  const stepLines = steps.length * lineH + 40;
  const listsH = Math.max(objLines, stepLines) + 16;

  const shawahedHeaderH = photos.length > 0 ? 36 : 0;
  const photoRowH = photos.length > 0 ? 150 : 0;
  const photoRows = photos.length > 0 ? Math.ceil(photos.length / 2) : 0;
  const shawahedH = shawahedHeaderH + photoRows * (photoRowH + gridGap);

  const footerH = 30;
  const height = headerH + bannerH + 20 + infoGridH + 16 + listsH + 16 + shawahedH + footerH + pad;

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = Math.max(1, height) * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textBaseline = "middle";

  // ---- header: نص إداري يمين + الشعار وسط ----
  ctx.textAlign = "right";
  ctx.font = "bold 13px 'IBM Plex Sans Arabic', Tahoma, Arial";
  ctx.fillStyle = INK;
  let hy = 26;
  if (region) { ctx.fillText(`الإدارة العامة للتعليم بمنطقة ${region}`, width - pad, hy); hy += 20; }
  else { ctx.fillText("الإدارة العامة للتعليم", width - pad, hy); hy += 20; }
  ctx.font = "12px 'IBM Plex Sans Arabic', Tahoma, Arial";
  ctx.fillStyle = MUTED;
  ctx.fillText(`مكتب التعليم بـ${office || "....."}`, width - pad, hy);

  if (logoImageElement) {
    const s = 52;
    ctx.save();
    ctx.beginPath();
    ctx.roundRect((width - s) / 2, 18, s, s, 10);
    ctx.clip();
    ctx.drawImage(logoImageElement, (width - s) / 2, 18, s, s);
    ctx.restore();
  }
  ctx.textAlign = "left";
  ctx.font = "bold 11px 'IBM Plex Sans Arabic', Tahoma, Arial";
  ctx.fillStyle = DASH_GREEN;
  if (ministryName) ctx.fillText(ministryName, pad, 26);
  if (countryName) { ctx.font = "10px 'IBM Plex Sans Arabic', Tahoma, Arial"; ctx.fillStyle = MUTED; ctx.fillText(countryName, pad, 44); }

  let y = headerH;

  // ---- banner: اسم المدرسة ----
  ctx.fillStyle = DASH_GREEN;
  ctx.fillRect(pad, y, width - pad * 2, bannerH);
  ctx.textAlign = "center";
  ctx.font = "bold 16px 'IBM Plex Sans Arabic', Tahoma, Arial";
  ctx.fillStyle = "#fff";
  ctx.fillText(schoolName || "اسم المدرسة", width / 2, y + bannerH / 2);
  y += bannerH + 20;

  // ---- info grid: يمين (اسم البرنامج، المنفذون، المستفيدون) — يسار (مجال البرنامج، تاريخ التنفيذ، عدد المستفيدين) ----
  const colW = (width - pad * 2 - gridGap) / 2;
  const rightCol = [
    { label: "اسم البرنامج", value: p.programName || entry.title },
    { label: "المنفّذون", value: p.implementers || meta.teacherName || "" },
    { label: "المستفيدون", value: p.beneficiaries || "جميع الطلاب" },
  ];
  const leftCol = [
    { label: "مجال البرنامج", value: p.programField || cat.title },
    { label: "تاريخ التنفيذ", value: p.executionDate || formatDateDisplay(entry.date) },
    { label: "عدد المستفيدين", value: p.beneficiaryCount || "—" },
  ];

  const drawInfoBox = (x, boxY, w, h, label, value) => {
    ctx.fillStyle = "#F7F3E8";
    ctx.strokeStyle = "#E8DFC5";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.roundRect(x, boxY, w, h, 10);
    ctx.fill();
    ctx.stroke();
    ctx.textAlign = "right";
    ctx.font = "bold 12px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = DASH_GREEN;
    ctx.fillText(`${label}:`, x + w - 12, boxY + 20);
    ctx.font = "12px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = INK;
    const lines = wrapCanvasText(measure, String(value || "—"), w - 24).slice(0, 2);
    let vy = boxY + 42;
    lines.forEach((ln) => { ctx.fillText(ln, x + w - 12, vy); vy += 16; });
  };

  for (let i = 0; i < 3; i++) {
    const rowY = y + i * (boxH + gridGap);
    drawInfoBox(pad + colW + gridGap, rowY, colW, boxH, rightCol[i].label, rightCol[i].value);
    drawInfoBox(pad, rowY, colW, boxH, leftCol[i].label, leftCol[i].value);
  }
  y += infoGridH;

  // ---- الأهداف (يمين) + خطوات التنفيذ (يسار) ----
  const listBoxH = listsH;
  const drawListBox = (x, w, label, items) => {
    ctx.fillStyle = "#F7F3E8";
    ctx.strokeStyle = "#E8DFC5";
    ctx.beginPath();
    ctx.roundRect(x, y, w, listBoxH, 10);
    ctx.fill();
    ctx.stroke();
    ctx.textAlign = "right";
    ctx.font = "bold 12px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = DASH_GREEN;
    ctx.fillText(`${label}:`, x + w - 12, y + listBoxPad + 6);
    ctx.font = "11px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = INK;
    let iy = y + listBoxPad + 30;
    items.forEach((item, idx) => {
      const lines = wrapCanvasText(measure, `${idx + 1}. ${item}`, w - 24);
      lines.forEach((ln) => { ctx.fillText(ln, x + w - 12, iy); iy += lineH; });
    });
  };
  drawListBox(pad + colW + gridGap, colW, "الأهداف", objectives);
  drawListBox(pad, colW, "خطوات التنفيذ", steps);
  y += listBoxH + 16;

  // ---- الشواهد ----
  if (photos.length > 0) {
    ctx.fillStyle = DASH_GREEN;
    ctx.fillRect(pad, y, width - pad * 2, shawahedHeaderH);
    ctx.textAlign = "center";
    ctx.font = "bold 13px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = "#fff";
    ctx.fillText("الشواهد", width / 2, y + shawahedHeaderH / 2);
    y += shawahedHeaderH;

    const photoColW = (width - pad * 2 - gridGap) / 2;
    for (let i = 0; i < photoImgs.length; i++) {
      const ri = Math.floor(i / 2), ci = i % 2;
      const px = pad + ci * (photoColW + gridGap);
      const py = y + ri * (photoRowH + gridGap);
      ctx.save();
      ctx.beginPath();
      ctx.roundRect(px, py, photoColW, photoRowH, 8);
      ctx.clip();
      const img = photoImgs[i];
      const iw = img.width, ih = img.height, ir = iw / ih, tr = photoColW / photoRowH;
      let sx, sy, sw, sh;
      if (ir > tr) { sh = ih; sw = ih * tr; sx = (iw - sw) / 2; sy = 0; }
      else { sw = iw; sh = iw / tr; sx = 0; sy = (ih - sh) / 2; }
      ctx.drawImage(img, sx, sy, sw, sh, px, py, photoColW, photoRowH);
      ctx.restore();
      ctx.strokeStyle = LINE;
      ctx.strokeRect(px, py, photoColW, photoRowH);
    }
    y += photoRows * (photoRowH + gridGap);
  }

  // ---- footer ----
  ctx.textAlign = "center";
  ctx.font = "10px 'IBM Plex Sans Arabic', Tahoma, Arial";
  ctx.fillStyle = MUTED;
  ctx.fillText("عبر تطبيق فصولي", width / 2, y + 18);

  return { canvas, logicalWidth: width, logicalHeight: height };
}

function buildCertificateCanvas({ countryName, ministryName, schoolName, logoImageElement, title, studentName, reason, className, teacherName, principalName, date, accentColor }) {
  const width = 1000, height = 760, scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  const accent = accentColor || "#26423B";

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

  // Official header block: logo to the side (not top-center, to reduce visual
  // clutter in the vertical center) + country/ministry/school stacked beside it.
  const OFFSET = 30;
  const headerTextX = logoImageElement ? width / 2 + 45 : width / 2;
  if (logoImageElement) {
    const logoSize = 72;
    ctx.drawImage(logoImageElement, width - 130, 46, logoSize, logoSize);
  }
  let hy = 66;
  ctx.textAlign = "center";
  if (countryName) {
    ctx.fillStyle = "#232622";
    ctx.font = "bold 15px Tahoma, Arial";
    ctx.fillText(countryName, headerTextX, hy);
    hy += 20;
  }
  if (ministryName) {
    ctx.fillStyle = "#7A7768";
    ctx.font = "14px Tahoma, Arial";
    ctx.fillText(ministryName, headerTextX, hy);
    hy += 20;
  }
  if (schoolName) {
    ctx.fillStyle = "#7A7768";
    ctx.font = "16px Tahoma, Arial";
    ctx.fillText(schoolName, headerTextX, hy);
  }

  ctx.textAlign = "center";
  ctx.fillStyle = accent;
  ctx.font = "bold 58px 'Aref Ruqaa', Tahoma, Arial";
  ctx.fillText(title, width / 2, 168 + OFFSET);

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
// normal download if the Web Share API / file sharing isn't available. HTML
// content specifically falls back to opening in a new tab instead of forcing
// نستخدمه عشان نتجنب navigator.share() بالكمبيوتر — نافذة المشاركة الأصلية
// بالكمبيوتر (خصوصًا ويندوز) غالبًا ما تلقى تطبيقات تقدر تشارك معها، فترجّع
// خطأ نظام ("تعذّر إظهار جميع الطرق...") قبل ما يوصل الكود لنا أصلًا.
// يطبع الشاشة كما هي بالضبط (تشمل الأزرار والأدوات) عبر طباعة المتصفح
// الحقيقية، بدل تصدير جدول منسّق — يُفعَّل فقط لحظة الطباعة عبر كلاس مؤقت
// على body حتى لا يؤثر على أي مكان ثاني بالموقع.
function printCurrentScreen() {
  document.body.classList.add("print-class-only");
  const cleanup = () => document.body.classList.remove("print-class-only");
  window.addEventListener("afterprint", cleanup, { once: true });
  setTimeout(cleanup, 5000); // شبكة أمان لو ما أطلق المتصفح afterprint لأي سبب
  window.print();
}

function isTouchPrimary() {
  try {
    return window.matchMedia && window.matchMedia("(pointer: coarse)").matches;
  } catch (e) {
    return false;
  }
}

// a download, since a forced-download of HTML is unreliable on iOS Safari
// (it's a "displayable" type, so Safari often just tries to show it and the
// download silently does nothing) — opening it always works everywhere, and
// the user can share/save/print from that tab using their browser's own tools.
async function shareOrDownloadFile(blob, filename, mime) {
  try {
    const file = new File([blob], filename, { type: mime });
    if (isTouchPrimary() && navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: filename });
      return;
    }
  } catch (e) {
    // fall through to plain download/open below
  }
  if (mime === "text/html") {
    const url = URL.createObjectURL(blob);
    const opened = window.open(url, "_blank");
    if (!opened) downloadBlob(blob, filename); // popup blocked — last resort
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    return;
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
  if (job.type === "shawahedReport") {
    const entries = job.shawahed.entries || {};
    const headers = ["الفئة", "عنوان الشاهد", "ملاحظات", "التاريخ"];
    const rows = [];
    SHAWAHED_CATEGORIES.forEach((cat) => {
      (entries[cat.key] || []).forEach((e) => rows.push([cat.title, e.title, e.notes || "", formatDateDisplay(e.date)]));
    });
    return { title: "تقرير شواهد الأداء الوظيفي", subtitle: `${rows.length} شاهد`, headers, rows, filename: "تقرير-شواهد-الأداء-الوظيفي" };
  }
  if (job.type === "shawahedReportV2") {
    const entries = job.shawahed.entries || {};
    const cats = SHAWAHED_CATEGORIES.filter((c) => job.selectedKeys.includes(c.key));
    const headers = ["الفئة", "عنوان الشاهد", "ملاحظات", "التاريخ"];
    const rows = [];
    cats.forEach((cat) => {
      (entries[cat.key] || []).forEach((e) => rows.push([cat.title, e.title, e.notes || "", formatDateDisplay(e.date)]));
    });
    const isSingle = cats.length === 1;
    return {
      title: isSingle ? `شاهد — ${cats[0].title}` : "تقرير شواهد الأداء الوظيفي",
      subtitle: `${rows.length} شاهد عبر ${cats.length} فئة`,
      headers, rows,
      filename: isSingle ? `شاهد-${cats[0].title}` : "تقرير-شواهد-الأداء-الوظيفي",
    };
  }
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
    return { title: cls.emoji ? `${cls.emoji} ${cls.subject}` : cls.subject, subtitle: dateKey ? `${cls.grade} • ${formatDateDisplay(dateKey)}` : `${cls.grade} • ${cls.teacher}`, headers, rows, filename: cls.subject || "الفصل", blankTemplate: true };
  }
  if (job.type === "blank") {
    const cls = job.cls;
    const headers = ["الاسم", ...cls.columns.map((c) => c.name)];
    const rows = cls.rows.map((row) => [row.name, ...cls.columns.map(() => "")]);
    return { title: cls.subject, subtitle: `${cls.grade} • ${cls.teacher} — نسخة فارغة`, headers, rows, filename: `${cls.subject || "الفصل"}-فارغ`, blankTemplate: true };
  }
  if (job.type === "attendance") {
    const { cls, dateKey } = job;
    const headers = ["الاسم", "الحالة"];
    const rows = cls.rows.map((row) => [row.name, attendanceStatus(cls, row.id, dateKey) === "absent" ? "غائب" : "حاضر"]);
    return { title: `الحضور والغياب — ${formatDateDisplay(dateKey)}`, subtitle: `${cls.subject} • ${cls.grade}`, headers, rows, filename: `حضور-${cls.subject || "الفصل"}-${dateKey}` };
  }
  if (job.type === "classFullReport") {
    const cls = job.cls;
    return { title: `تقرير شامل — ${cls.subject}`, subtitle: `${cls.grade} • ${cls.teacher}`, headers: [], rows: [], filename: `تقرير-شامل-${cls.subject || "الفصل"}` };
  }
  if (job.type === "gradeSheet") {
    const { cls, shortTestIds, finalExamIds, reviewerName } = job;
    const hasFinal = (finalExamIds || []).length > 0;
    const headers = ["الاسم", "معدل الاختبارات القصيرة", ...(hasFinal ? ["معدل اختبار نهاية الفصل"] : []), "المجموع", "توقيع المراجع"];
    const rows = cls.rows.map((row) => {
      const shortVals = shortTestIds.map((id) => Number(cls.cells[`${row.id}:${id}`]) || 0);
      const shortAvg = shortVals.length ? Math.round((shortVals.reduce((a, b) => a + b, 0) / shortVals.length) * 100) / 100 : 0;
      const finalVals = (finalExamIds || []).map((id) => Number(cls.cells[`${row.id}:${id}`]) || 0);
      const finalAvg = finalVals.length ? Math.round((finalVals.reduce((a, b) => a + b, 0) / finalVals.length) * 100) / 100 : 0;
      const total = Math.round((shortAvg + finalAvg) * 100) / 100;
      return [row.name, shortAvg, ...(hasFinal ? [finalAvg] : []), total, ""];
    });
    return {
      title: `كشف رصد درجات المادة — ${cls.subject}`,
      subtitle: `${cls.grade} • ${cls.teacher}${reviewerName ? ` • المراجع: ${reviewerName}` : ""}`,
      headers, rows, filename: `كشف-رصد-درجات-${cls.subject || "الفصل"}`,
    };
  }
  if (job.type === "periodComparison") {
    const { cls, colName, from1, to1, from2, to2, rows: compRows } = job;
    const headers = ["الاسم", `الفترة الأولى (${from1} إلى ${to1})`, `الفترة الثانية (${from2} إلى ${to2})`, "الفرق"];
    const rows = compRows.map((r) => [
      r.row.name,
      r.avg1 === null ? "—" : Math.round(r.avg1 * 100) / 100,
      r.avg2 === null ? "—" : Math.round(r.avg2 * 100) / 100,
      r.diff === null ? "—" : (r.diff > 0 ? "+" : "") + Math.round(r.diff * 100) / 100,
    ]);
    return {
      title: `مقارنة أداء — ${colName || ""}`,
      subtitle: `${cls.subject} • ${cls.grade}`,
      headers, rows, filename: `مقارنة-أداء-${cls.subject || "الفصل"}`,
    };
  }
  if (job.type === "parentReport") {
    const { cls, row, entries } = job;
    const grouped = groupEntries(entries).flatMap((g) => g.items);
    const headers = ["العمود", "اليوم والتاريخ", "الوقت", "القيمة"];
    const rows = grouped.map((e) => [e.colName, `${e.day ? e.day + "، " : ""}${e.date || ""}`, e.time || "", e.value]);
    return { title: `تقرير ولي الأمر: ${row.name}`, subtitle: `${cls.subject} • ${cls.grade}`, headers, rows, filename: `تقرير-ولي-الأمر-${row.name}` };
  }
  if (job.type === "programReport") {
    const { entry, cat } = job;
    const p = entry.program || {};
    const headers = ["الحقل", "القيمة"];
    const rows = [
      ["اسم البرنامج", p.programName || entry.title],
      ["مجال البرنامج", p.programField || cat.title],
      ["المنفّذون", p.implementers || ""],
      ["تاريخ التنفيذ", p.executionDate || formatDateDisplay(entry.date)],
      ["المستفيدون", p.beneficiaries || ""],
      ["عدد المستفيدين", p.beneficiaryCount || ""],
      ...(p.objectives || []).map((o, i) => [`الهدف ${i + 1}`, o]),
      ...(p.steps || []).map((s, i) => [`خطوة التنفيذ ${i + 1}`, s]),
    ];
    return { title: p.programName || entry.title, subtitle: cat.title, headers, rows, filename: `تقرير-برنامج-${entry.title}` };
  }
  const { cls, row, entries } = job;
  const grouped = groupEntries(entries).flatMap((g) => g.items);
  const headers = ["العمود", "اليوم والتاريخ", "الوقت", "القيمة"];
  const rows = grouped.map((e) => [e.colName, `${e.day ? e.day + "، " : ""}${e.date || ""}`, e.time || "", e.value]);
  return { title: `تقرير الطالب: ${row.name}`, subtitle: `${cls.subject} • ${cls.grade}`, headers, rows, filename: `تقرير-${row.name}` };
}

// يبني تقرير الشواهد الملوّن: يمر على كل فئة من فئات الأداء الوظيفي الاثنتا
// عشرة، ويرسم عنوانها الملوّن، ثم كل شاهد فيها (عنوانه، ملاحظاته، صورته إن
// وُجدت). الفئات الفاضية من الشواهد تُتخطى بدون ما تاخذ مساحة بالتقرير.
async function buildShawahedReportCanvas(shawahed, meta = {}) {
  const { countryName, ministryName, schoolName, logoImage, teacherName, principalName } = meta;
  const entries = shawahed.entries || {};
  const scale = 3;
  const width = 900;
  const pad = 50;
  const rowH = 84;
  const catHeaderH = 34;
  const catGap = 20;

  const activeCats = SHAWAHED_CATEGORIES.filter((c) => (entries[c.key] || []).length > 0);
  let totalEntries = 0;
  activeCats.forEach((c) => { totalEntries += entries[c.key].length; });

  let logoImageElement = null;
  if (logoImage) {
    try { logoImageElement = await loadImage(logoImage); } catch (e) { logoImageElement = null; }
  }

  // ارتفاع الترويسة الرسمية (شعار + دولة + وزارة + مدرسة + عنوان + بيانات المعلم)
  let headerH = 30;
  if (countryName) headerH += 20;
  if (ministryName) headerH += 18;
  if (schoolName) headerH += 24;
  headerH += 44; // العنوان
  headerH += 56; // بيانات المعلم/التاريخ
  headerH += 20;

  const footerH = 110; // مساحة التوقيعات
  const height = headerH + activeCats.length * (catHeaderH + catGap) + totalEntries * rowH + footerH + pad;

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textAlign = "center";
  ctx.textBaseline = "alphabetic";

  let hy = 34;
  if (logoImageElement) ctx.drawImage(logoImageElement, width - pad - 56, 14, 56, 56);
  if (countryName) { ctx.font = "bold 15px Tahoma, Arial"; ctx.fillStyle = INK; ctx.fillText(countryName, width / 2, hy); hy += 20; }
  if (ministryName) { ctx.font = "13px Tahoma, Arial"; ctx.fillStyle = MUTED; ctx.fillText(ministryName, width / 2, hy); hy += 18; }
  if (schoolName) { ctx.font = "bold 14px Tahoma, Arial"; ctx.fillStyle = INK; ctx.fillText(schoolName, width / 2, hy); hy += 24; }

  hy += 12;
  ctx.font = "bold 22px Tahoma, Arial";
  ctx.fillStyle = "#26423B";
  ctx.fillText("سجل توثيق شواهد الأداء الوظيفي", width / 2, hy);
  hy += 40;

  ctx.strokeStyle = LINE;
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(pad, hy); ctx.lineTo(width - pad, hy); ctx.stroke();
  hy += 26;

  ctx.textAlign = "right";
  ctx.font = "14px Tahoma, Arial";
  ctx.fillStyle = INK;
  ctx.fillText(`اسم المعلم/ـة: ${teacherName && teacherName.trim() ? teacherName : "...................................."}`, width - pad, hy);
  ctx.textAlign = "left";
  ctx.fillText(`تاريخ الطباعة: ${formatDateDisplay(todayKey())}`, pad, hy);
  hy += 24;
  ctx.textAlign = "right";
  ctx.fillText(`إجمالي الشواهد الموثّقة: ${totalEntries} شاهدًا عبر ${activeCats.length} من أصل ${SHAWAHED_CATEGORIES.length} معيارًا`, width - pad, hy);
  hy += 30;

  let y = hy;
  for (const cat of activeCats) {
    const officialIndex = SHAWAHED_CATEGORIES.indexOf(cat) + 1;
    ctx.fillStyle = cat.color;
    ctx.fillRect(pad, y, width - pad * 2, catHeaderH);
    ctx.fillStyle = "#fff";
    ctx.font = "bold 14px Tahoma, Arial";
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    ctx.fillText(`${officialIndex}. ${cat.title}`, width - pad - 12, y + catHeaderH / 2);
    y += catHeaderH;

    for (const entry of entries[cat.key]) {
      const rowY = y;
      ctx.strokeStyle = LINE;
      ctx.lineWidth = 1;
      ctx.strokeRect(pad, rowY, width - pad * 2, rowH - 8);

      let photoImg = null;
      if (entry.photo) {
        try { photoImg = await loadImage(entry.photo); } catch (e) { photoImg = null; }
      }
      if (photoImg) {
        const thumbSize = rowH - 24;
        ctx.save();
        ctx.beginPath();
        ctx.roundRect(width - pad - 12 - thumbSize, rowY + 12, thumbSize, thumbSize, 8);
        ctx.clip();
        ctx.drawImage(photoImg, width - pad - 12 - thumbSize, rowY + 12, thumbSize, thumbSize);
        ctx.restore();
      }
      const textRightEdge = photoImg ? width - pad - 24 - (rowH - 24) : width - pad - 16;
      ctx.textAlign = "right";
      ctx.fillStyle = INK;
      ctx.font = "bold 14px Tahoma, Arial";
      ctx.fillText(entry.title, textRightEdge, rowY + 26);
      if (entry.notes) {
        ctx.font = "12px Tahoma, Arial";
        ctx.fillStyle = MUTED;
        const wrapped = wrapCanvasText(ctx, entry.notes, textRightEdge - pad - 10);
        ctx.fillText(wrapped[0] || "", textRightEdge, rowY + 48);
      }
      ctx.font = "11px Tahoma, Arial";
      ctx.fillStyle = MUTED;
      ctx.fillText(formatDateDisplay(entry.date), textRightEdge, rowY + rowH - 26);

      y += rowH;
    }
    y += catGap;
  }

  // خط توقيعات رسمي بالأسفل
  y += 20;
  ctx.strokeStyle = LINE;
  ctx.beginPath(); ctx.moveTo(pad, y); ctx.lineTo(width - pad, y); ctx.stroke();
  y += 40;
  ctx.textAlign = "center";
  ctx.font = "13px Tahoma, Arial";
  ctx.fillStyle = INK;
  const sigColW = (width - pad * 2) / 2;
  ctx.fillText("توقيع المعلم/ـة: ....................................", pad + sigColW * 1.5 - sigColW / 2, y);
  ctx.fillText(`توقيع مدير/ة المدرسة${principalName ? ` (${principalName})` : ""}: ....................................`, pad + sigColW / 2, y);

  return { canvas, logicalWidth: width, logicalHeight: height };
}

// يبني تقرير الشواهد بتصميم لوحة معلومات رسمية: بانر علوي بالشعار، شرائط
// إحصائية، بطاقات فئات ملوّنة، صفوف شواهد بعلامة صح وصورة مصغّرة، وصناديق
// توقيع بالأسفل. يقبل قائمة مفاتيح الفئات المختارة فقط (تقرير فئة واحدة أو
// عدة فئات مدمجة معًا حسب اختيار المستخدم).
async function buildShawahedReportCanvasV2(shawahed, selectedKeys, meta = {}) {
  const { countryName, ministryName, schoolName, logoImage, teacherName, principalName, description } = meta;
  const entries = shawahed.entries || {};
  const cats = getAllShawahedCategories(shawahed).filter((c) => selectedKeys.includes(c.key) && (entries[c.key] || []).length > 0);
  const scale = 3;
  const width = 860;
  const pad = 40;

  let logoImageElement = null;
  if (logoImage) {
    try { logoImageElement = await loadImage(logoImage); } catch (e) { logoImageElement = null; }
  }

  const measure = document.createElement("canvas").getContext("2d");
  measure.font = "13px 'IBM Plex Sans Arabic', Tahoma, Arial";

  // ---- الوصف العام: لازم يظهر دايمًا حتى لو الطلب جا من مسار (زي الطباعة
  // السريعة لفئة وحدة) ما مرّ على نافذة اختيار الوصف — نولّد نص افتراضي
  // هنا لضمان عدم بقاء التقرير بلا وصف إطلاقًا.
  const totalEntriesCount = cats.reduce((s, c) => s + (entries[c.key]?.length || 0), 0);
  const finalDescription = (description && description.trim())
    ? description.trim()
    : (cats.length === 1
      ? `يوثّق هذا التقرير ${totalEntriesCount} شاهدًا على أداء ${teacherName || "المعلم/ـة"} فيما يخص معيار "${cats[0].title}"، ويُظهر التزامه/ـا التطبيقي بهذا الجانب من الأداء الوظيفي.`
      : cats.length > 1
      ? `يوثّق هذا التقرير ${totalEntriesCount} شاهدًا على الأداء الوظيفي لـ ${teacherName || "المعلم/ـة"}، موزّعة على ${cats.length} معايير مختلفة.`
      : "");

  // ---- header sizing (٣ أعمدة: يمين الدولة/الوزارة، وسط الشعار، يسار المدرسة) ----
  const headerRowH = 78;
  let headerH = 26 + headerRowH + 20; // شريط أعلى + صف الترويسة + هامش
  headerH += 44; // عنوان التقرير
  headerH += 12;

  // ---- description sizing ----
  const descLines = finalDescription ? wrapCanvasText(measure, finalDescription, width - pad * 2) : [];
  const descH = descLines.length ? descLines.length * 20 + 30 : 0;

  // ---- cards grid sizing (2 columns) ----
  const gap = 16;
  const cardW = (width - pad * 2 - gap) / 2;
  const cardImgH = 130;
  const cardTextH = 96;
  const cardH = cardImgH + cardTextH;
  const catHeaderH = 38;

  let bodyH = 0;
  const catRowsCount = {};
  cats.forEach((cat) => {
    const n = (entries[cat.key] || []).length;
    const rows = Math.ceil(n / 2);
    catRowsCount[cat.key] = rows;
    bodyH += catHeaderH + 14 + rows * (cardH + gap);
  });

  const footerH = 110;
  const height = headerH + descH + bodyH + footerH + pad;

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = Math.max(1, height) * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textBaseline = "middle";

  // ---- clean header: يمين (الدولة/الوزارة) — وسط (الشعار) — يسار (المدرسة) ----
  ctx.fillStyle = DASH_GREEN;
  ctx.fillRect(0, 0, width, 6);

  const rowTop = 26;
  const rowMid = rowTop + headerRowH / 2;

  ctx.textAlign = "right";
  if (countryName) {
    ctx.font = "bold 14px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = INK;
    ctx.fillText(countryName, width - pad, rowMid - 11);
  }
  if (ministryName) {
    ctx.font = "12px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText(ministryName, width - pad, rowMid + 11);
  }

  ctx.textAlign = "left";
  if (schoolName) {
    ctx.font = "bold 14px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = DASH_GREEN;
    const schoolLines = wrapCanvasText(measure, schoolName, 220).slice(0, 2);
    let sy = rowMid - (schoolLines.length - 1) * 10;
    schoolLines.forEach((ln) => { ctx.fillText(ln, pad, sy); sy += 20; });
  }

  if (logoImageElement) {
    const s = 58;
    ctx.save();
    ctx.beginPath();
    ctx.roundRect((width - s) / 2, rowTop, s, s, 12);
    ctx.clip();
    ctx.drawImage(logoImageElement, (width - s) / 2, rowTop, s, s);
    ctx.restore();
  }

  let hy = rowTop + headerRowH + 20;
  ctx.textAlign = "center";
  ctx.font = "bold 22px 'IBM Plex Sans Arabic', Tahoma, Arial";
  ctx.fillStyle = INK;
  const reportTitle = cats.length === 1 ? cats[0].title : "سجل توثيق شواهد الأداء الوظيفي";
  ctx.fillText(reportTitle, width / 2, hy);
  hy += 18;
  ctx.strokeStyle = GOLD;
  ctx.lineWidth = 2.5;
  ctx.beginPath();
  ctx.moveTo(width / 2 - 34, hy);
  ctx.lineTo(width / 2 + 34, hy);
  ctx.stroke();

  let y = headerH;

  // ---- description paragraph ----
  if (descLines.length) {
    ctx.textAlign = "center";
    ctx.font = "13px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = MUTED;
    let dy = y + 16;
    descLines.forEach((ln) => { ctx.fillText(ln, width / 2, dy); dy += 20; });
    y += descH;
  }

  // ---- category cards grid ----
  for (const cat of cats) {
    const list = entries[cat.key] || [];
    const officialIndex = cats.indexOf(cat) + 1;

    ctx.textAlign = "right";
    ctx.font = "bold 14px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = cat.color;
    ctx.fillText(`${officialIndex}.  ${cat.title}`, width - pad, y + catHeaderH / 2);
    ctx.strokeStyle = `${cat.color}33`;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(pad, y + catHeaderH - 6);
    ctx.lineTo(width - pad, y + catHeaderH - 6);
    ctx.stroke();
    y += catHeaderH + 14;

    for (let ri = 0; ri < list.length; ri += 2) {
      const pairY = y;
      for (let ci = 0; ci < 2; ci++) {
        const entry = list[ri + ci];
        if (!entry) continue;
        const cx = pad + ci * (cardW + gap);

        ctx.fillStyle = "#fff";
        ctx.strokeStyle = LINE;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.roundRect(cx, pairY, cardW, cardH, 12);
        ctx.fill();
        ctx.stroke();

        // image area (or colored placeholder)
        let photoImg = null;
        if (entry.photo) { try { photoImg = await loadImage(entry.photo); } catch (e) { photoImg = null; } }
        ctx.save();
        ctx.beginPath();
        ctx.roundRect(cx, pairY, cardW, cardImgH, [12, 12, 0, 0]);
        ctx.clip();
        if (photoImg) {
          const iw = photoImg.width, ih = photoImg.height;
          const ir = iw / ih, tr = cardW / cardImgH;
          let sx, sy, sw, sh;
          if (ir > tr) { sh = ih; sw = ih * tr; sx = (iw - sw) / 2; sy = 0; }
          else { sw = iw; sh = iw / tr; sx = 0; sy = (ih - sh) / 2; }
          ctx.drawImage(photoImg, sx, sy, sw, sh, cx, pairY, cardW, cardImgH);
        } else {
          ctx.fillStyle = `${cat.color}12`;
          ctx.fillRect(cx, pairY, cardW, cardImgH);
          ctx.fillStyle = cat.color;
          ctx.textAlign = "center";
          ctx.font = "bold 26px 'IBM Plex Sans Arabic', Tahoma, Arial";
          ctx.fillText("✓", cx + cardW / 2, pairY + cardImgH / 2);
        }
        ctx.restore();

        // شريط علوي ملوّن بلون المعيار فوق الصورة — يعطي إحساس أرسمي أوضح
        ctx.save();
        ctx.beginPath();
        ctx.roundRect(cx, pairY, cardW, 4, [12, 12, 0, 0]);
        ctx.clip();
        ctx.fillStyle = cat.color;
        ctx.fillRect(cx, pairY, cardW, 4);
        ctx.restore();

        // text area
        const textY = pairY + cardImgH;
        ctx.textAlign = "right";
        ctx.fillStyle = INK;
        ctx.font = "bold 13px 'IBM Plex Sans Arabic', Tahoma, Arial";
        const titleLines = wrapCanvasText(measure, entry.title, cardW - 24).slice(0, 2);
        let ty = textY + 18;
        titleLines.forEach((ln) => { ctx.fillText(ln, cx + cardW - 12, ty); ty += 17; });

        ty += 4;
        ctx.font = "11px 'IBM Plex Sans Arabic', Tahoma, Arial";
        ctx.fillStyle = MUTED;
        const descText = entry.notes && entry.notes.trim()
          ? entry.notes.trim()
          : `شاهد يوثّق "${entry.title}" ضمن معيار "${cat.title}"، ويعكس تطبيقًا فعليًا لهذا الجانب من الأداء الوظيفي داخل الفصل.`;
        const noteLines = wrapCanvasText(measure, descText, cardW - 24).slice(0, 2);
        noteLines.forEach((ln) => { ctx.fillText(ln, cx + cardW - 12, ty); ty += 15; });

        ctx.textAlign = "left";
        ctx.font = "10px 'IBM Plex Sans Arabic', Tahoma, Arial";
        ctx.fillStyle = MUTED;
        ctx.fillText(formatDateDisplay(entry.date), cx + 12, textY + cardTextH - 12);
      }
      y += cardH + gap;
    }
    y += 6;
  }

  // ---- footer signature cards ----
  y += 14;
  const boxW = (width - pad * 2 - 20) / 2;
  const boxH = footerH - 30;
  ctx.textAlign = "center";
  [
    { label: "توقيع المعلم/ـة", name: teacherName, x: pad },
    { label: "توقيع مدير/ة المدرسة", name: principalName, x: pad + boxW + 20 },
  ].forEach((box) => {
    ctx.fillStyle = "#FAF8F3";
    ctx.strokeStyle = LINE;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.roundRect(box.x, y, boxW, boxH, 12);
    ctx.fill();
    ctx.stroke();
    ctx.fillStyle = DASH_GREEN;
    ctx.font = "bold 12px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillText(box.label, box.x + boxW / 2, y + 24);
    ctx.font = "12px 'IBM Plex Sans Arabic', Tahoma, Arial";
    ctx.fillStyle = INK;
    ctx.fillText(box.name && box.name.trim() ? box.name : "....................................", box.x + boxW / 2, y + 48);
    ctx.strokeStyle = "#D8CDA8";
    ctx.beginPath();
    ctx.moveTo(box.x + 20, y + boxH - 16);
    ctx.lineTo(box.x + boxW - 20, y + boxH - 16);
    ctx.stroke();
  });

  return { canvas, logicalWidth: width, logicalHeight: height };
}

async function jobToCanvas(job) {
  if (job.type === "shawahedReport") {
    return buildShawahedReportCanvas(job.shawahed, job.meta || {});
  }
  if (job.type === "shawahedReportV2") {
    return buildShawahedReportCanvasV2(job.shawahed, job.selectedKeys, job.meta || {});
  }
  if (job.type === "report") {
    const { cls, row, entries } = job;
    const groups = groupEntries(entries);
    let photoImageElement = null;
    if (row.photo) {
      try { photoImageElement = await loadImage(row.photo); } catch (e) { photoImageElement = null; }
    }
    return buildReportCanvas({ title: `تقرير الطالب: ${row.name}`, subtitle: `${cls.subject} • ${cls.grade} • ${cls.teacher}`, groups, photoImageElement });
  }
  if (job.type === "parentReport") {
    const { cls, row, entries, meta } = job;
    return buildParentReportCanvas({ cls, row, entries, meta: meta || {} });
  }
  if (job.type === "programReport") {
    const { entry, cat, meta } = job;
    return buildProgramReportCanvas({ entry, cat, meta: meta || {} });
  }
  if (job.type === "classFullReport") {
    const cls = job.cls;
    const studentCanvases = [];
    for (const row of cls.rows) {
      const entries = cls.reports?.[row.id] || [];
      const groups = groupEntries(entries);
      let photoImageElement = null;
      if (row.photo) {
        try { photoImageElement = await loadImage(row.photo); } catch (e) { photoImageElement = null; }
      }
      const { canvas } = buildReportCanvas({ title: row.name, subtitle: `${cls.subject} • ${cls.grade}`, groups, photoImageElement });
      studentCanvases.push(canvas);
    }
    const scale = 3;
    const pad = 24 * scale;
    const headerH = 90 * scale;
    const gap = 30 * scale;
    const maxW = Math.max(headerH, ...studentCanvases.map((c) => c.width), 0) + pad * 2;
    const totalH = headerH + studentCanvases.reduce((sum, c) => sum + c.height + gap, pad);
    const combined = document.createElement("canvas");
    combined.width = maxW;
    combined.height = totalH;
    const ctx = combined.getContext("2d");
    ctx.fillStyle = PAPER;
    ctx.fillRect(0, 0, maxW, totalH);
    ctx.direction = "rtl";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = INK;
    ctx.font = "bold 34px Tahoma, Arial";
    ctx.fillText(`تقرير شامل — ${cls.subject}`, maxW / 2, pad / 2 + 10 * scale);
    ctx.font = "20px Tahoma, Arial";
    ctx.fillStyle = MUTED;
    ctx.fillText(`${cls.grade} • ${cls.teacher} — ${cls.rows.length} طالب`, maxW / 2, pad / 2 + 44 * scale);

    let y = headerH;
    studentCanvases.forEach((c) => {
      const x = (maxW - c.width) / 2;
      ctx.drawImage(c, x, y);
      y += c.height + gap;
    });
    return { canvas: combined, logicalWidth: maxW / scale, logicalHeight: totalH / scale };
  }
  const table = jobToTable(job);
  return buildTableCanvas(table);
}

async function exportPng(job) {
  const filename = jobToTable(job).filename;
  const { canvas } = await jobToCanvas(job);
  canvas.toBlob((blob) => downloadBlob(blob, `${filename}.png`));
}

// Builds a real PDF (image-wrapped, so Arabic renders correctly) and shares
// it directly via the device share sheet when available, otherwise downloads it.
async function exportPdfShare(job) {
  const filename = jobToTable(job).filename;
  const { canvas, logicalWidth, logicalHeight } = await jobToCanvas(job);
  const dataUrl = canvas.toDataURL("image/jpeg", 0.97);
  const blob = buildPdfFromJpegDataUrl(dataUrl, canvas.width, canvas.height, logicalWidth, logicalHeight);
  await shareOrDownloadFile(blob, `${filename}.pdf`, "application/pdf");
}

function exportExcel(job) {
  if (job.type === "classFullReport") {
    const cls = job.cls;
    const wb = XLSX.utils.book_new();
    cls.rows.forEach((row) => {
      const entries = cls.reports?.[row.id] || [];
      const grouped = groupEntries(entries).flatMap((g) => g.items);
      const headers = ["العمود", "اليوم والتاريخ", "الوقت", "القيمة"];
      const rows = grouped.map((e) => [e.colName, `${e.day ? e.day + "، " : ""}${e.date || ""}`, e.time || "", e.value]);
      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      ws["!cols"] = headers.map((_, i) => ({ wch: i === headers.length - 1 ? 45 : 18 }));
      // أسماء أوراق Excel محدودة بـ٣١ حرفًا وممنوع فيها بعض الرموز
      const safeName = (row.name || "طالب").replace(/[\\/*?:[\]]/g, " ").slice(0, 31) || "طالب";
      XLSX.utils.book_append_sheet(wb, ws, safeName);
    });
    XLSX.writeFile(wb, `تقرير-شامل-${cls.subject || "الفصل"}.xlsx`);
    return;
  }
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
      `<td style="padding:6px;text-align:center;border:1px solid ${LINE};font-weight:700;color:${status === "absent" ? "#C0392B" : "#26423B"};">${status === "absent" ? "غائب" : "حاضر"}</td>`,
      ...cls.columns.map((col) => {
        const val = (dateKey ? valueOnDate(cls, row.id, col.id, dateKey) : (cls.cells[`${row.id}:${col.id}`] || lastReportedValue(cls, row.id, col.id))) || "";
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
  <h1>${escapeHtml(row.name)}</h1>
  <p class="sub">${escapeHtml(cls.subject)} • ${escapeHtml(cls.grade)} • ${escapeHtml(cls.teacher)}</p>
  ${groups.length === 0 ? `<p style="color:${MUTED};">لا يوجد رصد بعد.</p>` : sectionsHtml}
</body></html>`;
}

// نسخة HTML نظيفة للمشاركة للقراءة فقط بتبويب الشواهد — بدون أي علامة أو
// إشارة "للقراءة فقط"، تبدو كصفحة تقرير رسمية عادية عند فتحها.
function buildReadOnlyShawahedHtml(shawahed, selectedKeys, meta = {}) {
  const { countryName, ministryName, schoolName, teacherName, principalName, description } = meta;
  const entries = shawahed.entries || {};
  const cats = getAllShawahedCategories(shawahed).filter((c) => selectedKeys.includes(c.key) && (entries[c.key] || []).length > 0);
  const reportTitle = cats.length === 1 ? cats[0].title : "سجل توثيق شواهد الأداء الوظيفي";
  const totalEntriesCount = cats.reduce((s, c) => s + (entries[c.key]?.length || 0), 0);
  const finalDescription = (description && description.trim())
    ? description.trim()
    : (cats.length === 1
      ? `يوثّق هذا التقرير ${totalEntriesCount} شاهدًا على أداء ${teacherName || "المعلم/ـة"} فيما يخص معيار "${cats[0].title}"، ويُظهر التزامه/ـا التطبيقي بهذا الجانب من الأداء الوظيفي.`
      : cats.length > 1
      ? `يوثّق هذا التقرير ${totalEntriesCount} شاهدًا على الأداء الوظيفي لـ ${teacherName || "المعلم/ـة"}، موزّعة على ${cats.length} معايير مختلفة.`
      : "");

  const catsHtml = cats.map((cat) => {
    const officialIndex = cats.indexOf(cat) + 1;
    const cardsHtml = (entries[cat.key] || []).map((e) => {
      const noteText = (e.notes && e.notes.trim())
        ? e.notes.trim()
        : `شاهد يوثّق "${e.title}" ضمن معيار "${cat.title}"، ويعكس تطبيقًا فعليًا لهذا الجانب من الأداء الوظيفي داخل الفصل.`;
      return `
      <div style="border:1px solid ${LINE};border-radius:14px;overflow:hidden;background:#fff;box-shadow:0 2px 8px rgba(35,38,34,0.06);border-top:3px solid ${cat.color};">
        ${e.photo ? `<img src="${e.photo}" style="width:100%;height:150px;object-fit:cover;display:block;" />` : `<div style="width:100%;height:150px;background:${cat.color}12;display:flex;align-items:center;justify-content:center;color:${cat.color};font-size:26px;font-weight:bold;">✓</div>`}
        <div style="padding:14px;">
          <p style="margin:0 0 6px;font-weight:bold;font-size:13px;color:${INK};">${escapeHtml(e.title)}</p>
          <p style="margin:0 0 8px;font-size:11px;color:${MUTED};line-height:1.6;">${escapeHtml(noteText)}</p>
          <p style="margin:0;font-size:10px;color:${MUTED};text-align:left;">${escapeHtml(formatDateDisplay(e.date))}</p>
        </div>
      </div>`;
    }).join("");
    return `
      <div style="margin-bottom:28px;">
        <p style="font-weight:bold;font-size:15px;color:${cat.color};border-bottom:1px solid ${cat.color}33;padding-bottom:8px;margin-bottom:14px;">${officialIndex}. ${escapeHtml(cat.title)}</p>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">${cardsHtml}</div>
      </div>`;
  }).join("");

  return `<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(reportTitle)}</title>
<style>
  body { font-family: Tahoma, Arial, sans-serif; background: ${PAPER}; color: ${INK}; padding: 0; margin: 0; }
  .wrap { max-width: 780px; margin: 0 auto; padding: 32px 20px 60px; }
  .accent-bar { height: 6px; background: ${DASH_GREEN}; }
  .header { text-align: center; margin-bottom: 24px; }
  .header h1 { font-size: 22px; margin: 14px 0 0; }
  .header .divider { width: 68px; height: 2px; background: ${GOLD}; margin: 10px auto 0; }
  .desc { text-align: center; color: ${MUTED}; font-size: 13px; max-width: 620px; margin: 16px auto 28px; line-height: 1.7; }
  @media (max-width: 500px) { div[style*="grid-template-columns"] { grid-template-columns: 1fr !important; } }
</style></head>
<body>
  <div class="accent-bar"></div>
  <div class="wrap">
    <div class="header">
      ${countryName ? `<p style="margin:0;font-weight:bold;font-size:14px;">${escapeHtml(countryName)}</p>` : ""}
      ${ministryName ? `<p style="margin:2px 0 0;color:${MUTED};font-size:12px;">${escapeHtml(ministryName)}</p>` : ""}
      ${schoolName ? `<p style="margin:2px 0 0;font-weight:bold;color:${DASH_GREEN};font-size:13px;">${escapeHtml(schoolName)}</p>` : ""}
      <h1>${escapeHtml(reportTitle)}</h1>
      <div class="divider"></div>
    </div>
    ${finalDescription ? `<p class="desc">${escapeHtml(finalDescription)}</p>` : ""}
    ${cats.length === 0 ? `<p style="text-align:center;color:${MUTED};">لا يوجد شواهد لعرضها.</p>` : catsHtml}
    <div style="display:flex;gap:16px;margin-top:28px;">
      <div style="flex:1;background:#FAF8F3;border:1px solid ${LINE};border-radius:12px;padding:16px;text-align:center;">
        <p style="margin:0 0 6px;font-weight:bold;font-size:12px;color:${DASH_GREEN};">توقيع المعلم/ـة</p>
        <p style="margin:0;font-size:12px;">${teacherName ? escapeHtml(teacherName) : "...................................."}</p>
      </div>
      <div style="flex:1;background:#FAF8F3;border:1px solid ${LINE};border-radius:12px;padding:16px;text-align:center;">
        <p style="margin:0 0 6px;font-weight:bold;font-size:12px;color:${DASH_GREEN};">توقيع مدير/ة المدرسة</p>
        <p style="margin:0;font-size:12px;">${principalName ? escapeHtml(principalName) : "...................................."}</p>
      </div>
    </div>
  </div>
</body></html>`;
}

function useFonts() {
  useEffect(() => {
    if (document.getElementById("mutabaa-fonts")) return;
    const link = document.createElement("link");
    link.id = "mutabaa-fonts";
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&family=IBM+Plex+Sans+Arabic:wght@400;500;600&family=Aref+Ruqaa:wght@400;700&display=swap";
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
        .app-print-root { display: block !important; position: absolute; inset: 0; padding: 24px; direction: rtl; }
        body.print-class-only * { visibility: hidden !important; }
        body.print-class-only .class-print-area, body.print-class-only .class-print-area * { visibility: visible !important; }
        body.print-class-only .class-print-area { position: absolute; inset: 0; padding: 16px; }
        body.print-class-only * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
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
        from { opacity: 0; transform: translateY(48px) scale(0.88); }
        to { opacity: 1; transform: translateY(0) scale(1); }
      }
      .modal-backdrop-in { animation: modalBackdropIn 0.3s ease-out forwards; }
      .modal-panel-in { animation: modalPanelIn 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
      @keyframes pageFadeIn {
        from { opacity: 0; transform: translateY(24px); }
        to { opacity: 1; transform: translateY(0); }
      }
      .page-fade-in { animation: pageFadeIn 0.5s ease-out forwards; }
      @keyframes cardIn {
        from { opacity: 0; transform: translateY(24px) scale(0.9); }
        to { opacity: 1; transform: translateY(0) scale(1); }
      }
      .card-in { animation: cardIn 0.55s cubic-bezier(0.16, 1, 0.3, 1) both; }
      .modal-max-height { max-height: 85vh; }
      @supports (max-height: 85dvh) { .modal-max-height { max-height: 85dvh; } }
      @keyframes toastPop {
        from { opacity: 0; transform: translateY(-14px) scale(0.95); }
        to { opacity: 1; transform: translateY(0) scale(1); }
      }
      .toast-pop { animation: toastPop 0.3s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
      @keyframes celebrateIn {
        0% { opacity: 0; transform: scale(0.7) translateY(10px); }
        60% { opacity: 1; transform: scale(1.05) translateY(-2px); }
        100% { opacity: 1; transform: scale(1) translateY(0); }
      }
      .celebrate-in { animation: celebrateIn 0.5s cubic-bezier(0.34, 1.56, 0.64, 1) forwards; }
      @keyframes confettiPop {
        0% { opacity: 0; transform: scale(0.3) rotate(0deg); }
        50% { opacity: 1; }
        100% { opacity: 0; transform: translate(var(--tx), var(--ty)) scale(1) rotate(var(--rot)); }
      }
      .confetti-bit { position: absolute; top: 50%; left: 50%; animation: confettiPop 0.9s ease-out forwards; }
      @keyframes bigCheckIn {
        0% { opacity: 0; transform: scale(0.3) rotate(-15deg); }
        55% { opacity: 1; transform: scale(1.25) rotate(6deg); }
        75% { transform: scale(0.92) rotate(-2deg); }
        100% { opacity: 1; transform: scale(1) rotate(0deg); }
      }
      .big-check-in { animation: bigCheckIn 0.6s cubic-bezier(0.34, 1.56, 0.64, 1) forwards; }
      @keyframes sadWobble {
        0% { opacity: 0; transform: scale(0.5) rotate(0deg); }
        30% { opacity: 1; transform: scale(1.1) rotate(-8deg); }
        50% { transform: scale(1) rotate(8deg); }
        70% { transform: scale(1) rotate(-4deg); }
        100% { opacity: 1; transform: scale(1) rotate(0deg); }
      }
      .sad-wobble-in { animation: sadWobble 0.6s ease-out forwards; }
      .magic-shimmer { background-size: 220% 220% !important; animation: magicShimmer 3.5s ease infinite; }
      @keyframes magicShimmer {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
      }
      .data-row > td { transition: background-color 0.12s; }
      .data-row:hover > td { background-color: #F3F1EA !important; }
      @keyframes drift1 { 0%, 100% { transform: translate(0, 0) scale(1); } 50% { transform: translate(-3%, 4%) scale(1.06); } }
      @keyframes drift2 { 0%, 100% { transform: translate(0, 0) scale(1); } 50% { transform: translate(4%, -3%) scale(1.05); } }
      .bg-blob-1 { animation: drift1 22s ease-in-out infinite; }
      .bg-blob-2 { animation: drift2 26s ease-in-out infinite; }
      input[type="text"]:focus, input[type="email"]:focus, input[type="password"]:focus,
      input[type="number"]:focus, input[type="date"]:focus, input[type="time"]:focus,
      input[type="search"]:not([type="color"]):focus, textarea:focus, select:focus {
        border-color: #26423B !important;
        box-shadow: 0 0 0 3px rgba(15,107,92,0.14);
      }
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

function IconBtn({ icon: Icon, label, onClick, tone = "default", magic = false, disabled = false }) {
  const tones = {
    default: { bg: "#fff", fg: INK, border: LINE, shadow: "0 1px 2px rgba(35,38,34,0.05)" },
    danger: { bg: "#FBEDEA", fg: "#9A3B2E", border: "#F5DCD5", shadow: "0 1px 2px rgba(154,59,46,0.06)" },
    primary: { bg: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})`, fg: "#fff", border: "transparent", shadow: "0 2px 8px rgba(38,66,59,0.28)" },
    magic: { bg: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)", fg: "#fff", border: "transparent", shadow: "0 3px 12px rgba(124,92,224,0.38)" },
  };
  const t = tones[magic ? "magic" : tone];
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className={`relative flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold transition-all hover:brightness-110 hover:-translate-y-px active:scale-95 active:translate-y-0 whitespace-nowrap disabled:opacity-35 disabled:pointer-events-none disabled:hover:brightness-100 disabled:hover:translate-y-0 ${magic ? "magic-shimmer" : ""}`}
      style={{ background: t.bg, color: t.fg, border: `1px solid ${t.border}`, boxShadow: t.shadow }}
    >
      {magic && <Sparkles size={12} strokeWidth={2.5} className="shrink-0" />}
      <Icon size={14} strokeWidth={2.25} />
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
                    <button key={i} onClick={() => gSelect(d)} className="text-xs py-1.5 rounded-md hover:bg-black/5" style={{ background: isSel ? "#26423B" : "transparent", color: isSel ? "#fff" : INK, fontWeight: isSel ? 700 : 400 }}>{d}</button>
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
                    <button key={i} onClick={() => hSelect(d)} className="text-xs py-1.5 rounded-md hover:bg-black/5" style={{ background: isSel ? "#26423B" : "transparent", color: isSel ? "#fff" : INK, fontWeight: isSel ? 700 : 400 }}>{d}</button>
                  ) : <div key={i} />;
                })}
              </div>
            </>
          )}
          <button onClick={() => selectIso(todayKey())} className="w-full mt-3 text-xs font-bold py-1.5 rounded-lg" style={{ border: `1px solid ${LINE}`, color: "#26423B" }}>اليوم</button>
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

function PrintPreviewModal({ job, format, onClose, onExport }) {
  const [imgUrl, setImgUrl] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const formatLabels = { pdf: "PDF", png: "PNG", excel: "Excel" };
  const formatIcons = { pdf: FileText, png: FileImage, excel: FileSpreadsheet };
  const PrimaryIcon = formatIcons[format] || FileText;

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError("");
    setImgUrl(null);
    (async () => {
      try {
        const { canvas } = await jobToCanvas(job);
        if (!cancelled) setImgUrl(canvas.toDataURL("image/png"));
      } catch (err) {
        console.error("فشلت معاينة الطباعة:", err);
        if (!cancelled) setError("تعذّرت معاينة الملف. يمكنك المتابعة والتصدير مباشرة بالأزرار تحت رغم ذلك.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [job]);

  return (
    <Modal title="معاينة قبل التصدير" onClose={onClose} lg zIndex={80}>
      <div className="rounded-xl overflow-auto mb-4" style={{ border: `1px solid ${LINE}`, background: "#F3F1E9", maxHeight: "60vh" }}>
        {loading ? (
          <div className="flex items-center justify-center py-20">
            <RefreshCw size={22} color={MUTED} className="animate-spin" />
          </div>
        ) : error ? (
          <div className="flex flex-col items-center justify-center gap-2 py-16 px-4 text-center">
            <AlertTriangle size={22} color="#C97A2B" />
            <p className="text-sm" style={{ color: "#8A4A1E" }}>{error}</p>
          </div>
        ) : (
          <img src={imgUrl} alt="معاينة" className="w-full h-auto" />
        )}
      </div>
      <button
        onClick={() => onExport(format)}
        className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white mb-2 transition-all hover:brightness-110 active:scale-95"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})`, boxShadow: "0 2px 8px rgba(38,66,59,0.28)" }}
      >
        <PrimaryIcon size={16} /> تصدير كـ {formatLabels[format]}
      </button>
      <div className="grid grid-cols-2 gap-2">
        {Object.keys(formatLabels).filter((k) => k !== format).map((k) => {
          const Icon = formatIcons[k];
          return (
            <button key={k} onClick={() => onExport(k)} className="flex items-center justify-center gap-1.5 py-2 rounded-lg text-xs font-semibold" style={{ border: `1px solid ${LINE}`, color: MUTED, background: "#fff" }}>
              <Icon size={13} /> تصدير كـ {formatLabels[k]}
            </button>
          );
        })}
      </div>
    </Modal>
  );
}

function Modal({ title, onClose, onBack, children, wide = false, lg = false, xl = false, zIndex = 50, accent = null }) {
  const widthClass = xl ? "max-w-6xl" : lg ? "max-w-5xl" : wide ? "max-w-3xl" : "max-w-md md:max-w-lg";
  const [pos, setPos] = useState({ x: 0, y: 0 });
  const [entered, setEntered] = useState(false);
  const dragRef = useRef({ dragging: false, startX: 0, startY: 0, startPos: { x: 0, y: 0 } });
  const isMagic = accent === "magic";

  useEffect(() => {
    const t = setTimeout(() => setEntered(true), 520);
    return () => clearTimeout(t);
  }, []);

  useEffect(() => {
    const onMove = (e) => {
      if (!dragRef.current.dragging) return;
      setPos({
        x: dragRef.current.startPos.x + (e.clientX - dragRef.current.startX),
        y: dragRef.current.startPos.y + (e.clientY - dragRef.current.startY),
      });
    };
    const onUp = () => { dragRef.current.dragging = false; };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    window.addEventListener("pointercancel", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      window.removeEventListener("pointercancel", onUp);
    };
  }, []);

  const onHeaderPointerDown = (e) => {
    dragRef.current = { dragging: true, startX: e.clientX, startY: e.clientY, startPos: pos };
  };

  return (
    <div
      className="fixed inset-0 flex items-center justify-center p-4 modal-backdrop-in"
      style={{ background: "rgba(25,28,25,0.55)", zIndex }}
      onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        className={`rounded-2xl w-full ${widthClass} modal-max-height overflow-y-auto ${entered ? "" : "modal-panel-in"}`}
        style={{ background: PAPER, border: `1px solid ${LINE}`, ...(pos.x || pos.y ? { transform: `translate(${pos.x}px, ${pos.y}px)` } : {}), boxShadow: isMagic ? "0 24px 60px rgba(78,111,224,0.28)" : "0 24px 60px rgba(20,22,20,0.28)" }}
      >
        <div
          className={`flex items-center justify-between px-5 py-4 sticky top-0 ${isMagic ? "magic-shimmer" : ""}`}
          style={isMagic
            ? { background: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)", backgroundSize: "220% 220%", cursor: "grab", touchAction: "none", userSelect: "none" }
            : { background: PAPER, borderBottom: `1px solid ${LINE}`, cursor: "grab", touchAction: "none", userSelect: "none" }}
          onPointerDown={onHeaderPointerDown}
        >
          <h3 className="font-bold text-lg flex items-center gap-2" style={{ color: isMagic ? "#fff" : INK }}>
            {onBack && (
              <button
                onClick={onBack}
                onPointerDown={(e) => e.stopPropagation()}
                title="رجوع"
                className={`p-1 rounded-full transition-all shrink-0 ${isMagic ? "hover:bg-white/20" : "hover:bg-black/5"}`}
              >
                <ArrowRight size={18} color={isMagic ? "#fff" : MUTED} />
              </button>
            )}
            {isMagic ? <Sparkles size={16} strokeWidth={2.5} /> : <Move size={14} color={MUTED} />}
            {title}
          </h3>
          <button
            onClick={onClose}
            onPointerDown={(e) => e.stopPropagation()}
            className={`p-1.5 rounded-full active:scale-90 transition-all shrink-0 ${isMagic ? "hover:bg-white/20" : "hover:bg-[#FBEDEA]"}`}
            style={{ touchAction: "auto" }}
          >
            <X size={18} color={isMagic ? "#fff" : MUTED} />
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
  borderRadius: "12px",
  border: `1px solid ${LINE}`,
  background: "#fff",
  fontFamily: "'IBM Plex Sans Arabic', sans-serif",
  fontSize: "14px",
  color: INK,
  outline: "none",
  transition: "border-color 0.15s, box-shadow 0.15s",
};

function ColorSwatches({ value, onChange, size = 8 }) {
  const customInputRef = useRef(null);
  return (
    <div className="flex items-center gap-2">
      <div className="relative shrink-0">
        <button
          type="button"
          title="اختر لونًا"
          onClick={() => customInputRef.current?.click()}
          className="rounded-full flex items-center justify-center transition-transform hover:scale-110"
          style={{ width: size * 5, height: size * 5, background: value || "#26423B", boxShadow: `0 0 0 2px #fff, 0 0 0 4px ${value || "#26423B"}` }}
        >
          <Pipette size={size * 1.6} color="#fff" style={{ filter: "drop-shadow(0 0 1.5px rgba(0,0,0,0.5))" }} />
        </button>
        <input
          ref={customInputRef}
          type="color"
          value={value || "#26423B"}
          onChange={(e) => onChange(e.target.value)}
          style={{ position: "absolute", inset: 0, opacity: 0, width: "100%", height: "100%", cursor: "pointer" }}
        />
      </div>
      <span className="text-xs font-mono" style={{ color: MUTED }}>{value || "#26423B"}</span>
    </div>
  );
}

// ---------- option editor (per-option color) ----------

// يدوّر على أول "نطاق" لوني يطابق قيمة معيّنة بعمود عدّاد — يسمح للمعلم
// يحدد ألوان مختلفة تمامًا حسب مجال القيمة (مثال: ١-٥ أحمر، ٦-١٠ برتقالي)
// بدل تدرّج لون واحد آلي.
function matchColorBand(bands, value) {
  if (!bands || bands.length === 0) return null;
  const n = Number(value);
  if (Number.isNaN(n)) return null;
  const band = bands.find((b) => n >= Number(b.min) && n <= Number(b.max));
  return band ? band.color : null;
}

function ColorBandsEditor({ bands, onChange }) {
  const safeBands = bands || [];
  const update = (id, patch) => onChange(safeBands.map((b) => (b.id === id ? { ...b, ...patch } : b)));
  const remove = (id) => onChange(safeBands.filter((b) => b.id !== id));
  const add = () => {
    const lastMax = safeBands.length ? Number(safeBands[safeBands.length - 1].max) : 0;
    onChange([...safeBands, { id: uid(), min: lastMax + 1, max: lastMax + 5, color: COLORS[safeBands.length % COLORS.length].hex }]);
  };

  return (
    <div className="mb-2">
      <p className="text-xs mb-1.5" style={{ color: MUTED }}>
        مثال: من ١ إلى ٥ أحمر، من ٦ إلى ١٠ برتقالي — كل نطاق بلونه الخاص، والخيار كامل مرن لك.
      </p>
      <div className="space-y-1.5">
        {safeBands.map((b) => (
          <div key={b.id} className="flex items-center gap-1.5 p-2 rounded-lg" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <input
              type="number"
              value={b.min}
              onChange={(e) => update(b.id, { min: Number(e.target.value) })}
              style={{ ...inputStyle, width: 56, padding: "6px 8px" }}
              placeholder="من"
            />
            <span className="text-xs" style={{ color: MUTED }}>إلى</span>
            <input
              type="number"
              value={b.max}
              onChange={(e) => update(b.id, { max: Number(e.target.value) })}
              style={{ ...inputStyle, width: 56, padding: "6px 8px" }}
              placeholder="إلى"
            />
            <div className="relative shrink-0">
              <button type="button" title="لون النطاق" className="w-7 h-7 rounded-full" style={{ background: b.color, boxShadow: `0 0 0 1px ${LINE}` }} />
              <input
                type="color"
                value={b.color}
                onChange={(e) => update(b.id, { color: e.target.value })}
                style={{ position: "absolute", inset: 0, opacity: 0, width: "100%", height: "100%", cursor: "pointer" }}
              />
            </div>
            <span className="flex-1 text-xs" style={{ color: INK }}>{b.min} — {b.max}</span>
            <button type="button" onClick={() => remove(b.id)} className="p-1 rounded hover:bg-black/5 shrink-0">
              <X size={13} color={MUTED} />
            </button>
          </div>
        ))}
      </div>
      <button type="button" onClick={add} className="mt-2 text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}>
        <Plus size={13} /> إضافة نطاق
      </button>
    </div>
  );
}

// المستويات الافتراضية لعمود "مستوى مهارة" جديد — قابلة للتعديل والحذف
// والإضافة بالكامل من المعلم (نفس محرر خيارات القوائم المنسدلة).
function defaultSkillLevels() {
  return [
    { id: uid(), label: "مبتدئ", color: "#C0392B" },
    { id: uid(), label: "ناشئ", color: "#C97A2B" },
    { id: uid(), label: "متمكن", color: "#2E7DA6" },
    { id: uid(), label: "متقن", color: "#0F9D58" },
  ];
}

function OptionsEditor({ options, onChange }) {
  const safeOptions = options || [];
  const update = (id, patch) => onChange(safeOptions.map((o) => (o.id === id ? { ...o, ...patch } : o)));
  const remove = (id) => onChange(safeOptions.filter((o) => o.id !== id));
  const add = () => onChange([...safeOptions, { id: uid(), label: "", color: COLORS[0].hex }]);
  const [showBulk, setShowBulk] = useState(false);
  const [bulkText, setBulkText] = useState("");
  const addBulk = () => {
    const lines = bulkText.split("\n").map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) return;
    const added = lines.map((label, i) => ({ id: uid(), label, color: COLORS[i % COLORS.length].hex }));
    onChange([...safeOptions, ...added]);
    setBulkText("");
    setShowBulk(false);
  };
  return (
    <div>
      <div className="space-y-2">
        {safeOptions.map((o) => (
          <div key={o.id} className="flex items-center gap-2 p-2 rounded-lg" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <span className="w-3.5 h-3.5 rounded-full shrink-0" style={{ background: o.color }} />
            <input
              value={o.label}
              onChange={(e) => update(o.id, { label: e.target.value })}
              placeholder="نص الخيار"
              className="flex-1 text-sm outline-none bg-transparent"
              style={{ color: INK }}
            />
            <div className="relative shrink-0">
              <button type="button" title="لون الخيار" className="w-5 h-5 rounded-full" style={{ background: o.color, boxShadow: `0 0 0 1px ${LINE}` }} />
              <input
                type="color"
                value={o.color}
                onChange={(e) => update(o.id, { color: e.target.value })}
                style={{ position: "absolute", inset: 0, opacity: 0, width: "100%", height: "100%", cursor: "pointer" }}
              />
            </div>
            <button type="button" onClick={() => remove(o.id)} className="p-1 rounded hover:bg-black/5 shrink-0">
              <X size={13} color={MUTED} />
            </button>
          </div>
        ))}
      </div>
      <div className="flex items-center gap-3 mt-2">
        <button type="button" onClick={add} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}>
          <Plus size={13} /> إضافة خيار
        </button>
        <button type="button" onClick={() => setShowBulk((s) => !s)} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}>
          <ListPlus size={13} /> إضافة خيارات دفعة واحدة
        </button>
      </div>
      {showBulk && (
        <div className="mt-2 p-2 rounded-lg" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
          <textarea
            value={bulkText}
            onChange={(e) => setBulkText(e.target.value)}
            placeholder={"اكتب كل خيار بسطر مستقل، مثال:\nممتاز\nجيد جدًا\nجيد\nيحتاج تحسين"}
            style={{ ...inputStyle, minHeight: 90, background: "#fff" }}
          />
          <button type="button" disabled={!bulkText.trim()} onClick={addBulk} className="mt-2 text-xs font-bold px-3 py-1.5 rounded-lg text-white disabled:opacity-40" style={{ background: "#26423B" }}>
            إضافة الخيارات
          </button>
        </div>
      )}
    </div>
  );
}

// ---------- reusable field group for a column/row draft ----------

function TypePicker({ type, setType, labels, showLevel = false }) {
  return (
    <div className="flex gap-2">
      {[
        { v: "text", label: labels?.text || "نص حر", icon: Type },
        { v: "counter", label: "عداد (+/-)", icon: Hash },
        { v: "dropdown", label: "قائمة منسدلة", icon: ListChecks },
        ...(showLevel ? [{ v: "level", label: "مستوى مهارة", icon: Layers }] : []),
      ].map((opt) => (
        <button
          key={opt.v}
          type="button"
          onClick={() => setType(opt.v)}
          className="flex-1 flex flex-col items-center gap-1 py-2.5 rounded-xl text-xs font-medium"
          style={{
            border: `1px solid ${type === opt.v ? "#26423B" : LINE}`,
            background: type === opt.v ? "#EAF3F0" : "#fff",
            color: type === opt.v ? "#26423B" : MUTED,
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
        <TypePicker
          type={draft.type}
          setType={(v) => set({ type: v, levels: v === "level" && !(draft.levels || []).length ? defaultSkillLevels() : draft.levels })}
          showLevel
        />
      </Field>
      {draft.type === "dropdown" && (
        <Field label="خيارات القائمة (لكل خيار لونه الخاص)">
          <OptionsEditor options={draft.options || []} onChange={(options) => set({ options })} />
        </Field>
      )}
      {draft.type === "level" && (
        <Field label="مستويات المهارة (من الأدنى للأعلى)" hint="الترتيب مهم — أول مستوى بالقائمة هو الأدنى، وآخر واحد هو الأعلى/الإتقان.">
          <OptionsEditor options={draft.levels || []} onChange={(levels) => set({ levels })} />
        </Field>
      )}
      {draft.type === "counter" && (
        <>
          <Field label="الدرجة القصوى لهذا العمود (اختياري)" hint="لو حددتها، يدخل هذا العمود بحساب 'الدرجة الكلية' التلقائية بالجدول.">
            <input
              type="number"
              value={draft.maxValue || ""}
              onChange={(e) => set({ maxValue: e.target.value })}
              style={inputStyle}
              placeholder="مثال: 10"
            />
          </Field>
          <label className="flex items-center gap-2 text-sm font-medium mt-1 mb-1" style={{ color: INK }}>
            <input type="checkbox" checked={!!draft.colorScale} onChange={(e) => set({ colorScale: e.target.checked })} />
            تلوين الخانة حسب نطاقات لونية تحددها أنت
          </label>
          {draft.colorScale && (
            <ColorBandsEditor bands={draft.colorBands || []} onChange={(colorBands) => set({ colorBands })} />
          )}
        </>
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
        بعد ٢.٥ ثانية من تسجيل أي قيمة، تُفرَّغ الخانة تلقائيًا — مفيد للرصد السريع المتكرر (مثل المشاركة) بدون الحاجة تمسح يدويًا بين كل طالب وآخر.
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
  return { key: uid(), name: "", type: "text", options: [], levels: [], color: COLORS[2].hex, autoRenew: false, pinned: false, bulkValue: "", behaviorFlag: false, behaviorThreshold: 3, colorScale: false, colorBands: [] };
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
                style={{ color: "#26423B" }}
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
            style={{ background: "#26423B" }}
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
  const photoInputRef = useRef(null);
  const handlePhotoChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => set({ photo: reader.result });
    reader.readAsDataURL(file);
    e.target.value = "";
  };
  return (
    <div className="rounded-xl p-3 mb-3" style={{ border: `1px solid ${LINE}`, background: "#FCFBF7" }}>
      {removable && (
        <div className="flex justify-end mb-1">
          <button onClick={onRemove} className="p-1 rounded hover:bg-black/5"><X size={14} color={MUTED} /></button>
        </div>
      )}
      <Field label="صورة الطالب (اختياري)" hint="تظهر في التقرير، وتظهر بالشهادات عند تفعيل خيار تضمين الصورة.">
        <input ref={photoInputRef} type="file" accept="image/*" onChange={handlePhotoChange} style={{ display: "none" }} />
        <div className="flex items-center gap-2">
          {draft.photo ? (
            <img src={draft.photo} alt="صورة الطالب" className="w-12 h-12 rounded-full object-cover dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />
          ) : (
            <div className="w-12 h-12 rounded-full flex items-center justify-center" style={{ background: "#F3F1E9" }}><User size={20} color={MUTED} /></div>
          )}
          <button type="button" onClick={() => photoInputRef.current?.click()} className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ border: `1px solid ${LINE}`, color: INK }}>
            {draft.photo ? "تغيير الصورة" : "رفع صورة"}
          </button>
          {draft.photo && (
            <button type="button" onClick={() => set({ photo: null })} className="text-xs font-semibold" style={{ color: "#C0392B" }}>إزالة</button>
          )}
        </div>
      </Field>
      <Field label="اسم الصف">
        <input style={inputStyle} value={draft.name} onChange={(e) => set({ name: e.target.value })} placeholder="مثال: اسم الطالب" />
      </Field>
      <Field label="طريقة عمل الصف" hint="نوع خلايا الجدول الفعلي يُحدَّد حسب نوع كل عمود؛ هذا التصنيف تسمية توضيحية للصف.">
        <TypePicker type={draft.type} setType={(v) => set({ type: v })} labels={{ text: "مربع فارغ" }} />
      </Field>
      {draft.type === "dropdown" && (
        <Field label="خيارات القائمة">
          <OptionsEditor options={draft.options || []} onChange={(options) => set({ options })} />
        </Field>
      )}
      <Field label="لون الصف">
        <ColorSwatches value={draft.color} onChange={(color) => set({ color })} />
      </Field>
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

function parsePastedNames(text) {
  return text
    .split("\n")
    .map((line) => {
      const trimmed = line.trim();
      if (!trimmed) return "";
      if (!trimmed.includes("\t")) return trimmed;
      // سطر متعدد الأعمدة (لصق من جدول نور مثلًا) — نلقط الخانة اللي تشبه اسمًا فعليًا
      const parts = trimmed.split("\t").map((p) => p.trim()).filter(Boolean);
      const nameLike = parts.find((p) => looksLikeName(p));
      return nameLike || parts[0] || "";
    })
    .filter(Boolean);
}

// يكتشف خانة "الاسم الرباعي" الفعلية داخل صف ملصوق من جدول نور — يشترط
// وجود أكثر من كلمة (الاسم الرباعي دائمًا عدة كلمات) وعدم احتوائه أرقامًا،
// ويتجاهل نصوص الأزرار/الروابط المتكررة بكل صف مثل "عرض"/"إضافة" وعناوين
// الأعمدة نفسها مثل "الاسم الرباعي"/"رقم الهوية".
const NOOR_TABLE_JUNK = ["الاسم الرباعي", "رقم الهوية", "خيارات", "ملاحظات", "عرض", "إضافة", "الطلاب", "تحديد الكل"];
function looksLikeFullName(s) {
  const t = String(s ?? "").trim();
  if (t.length < 4 || t.length > 80) return false;
  if (!/\s/.test(t)) return false; // اسم رباعي = عدة كلمات دائمًا
  if (/\d/.test(t)) return false;
  if (!/[\u0600-\u06FFa-zA-Z]/.test(t)) return false;
  if (NOOR_TABLE_JUNK.includes(t)) return false;
  return true;
}
// رقم الهوية/الإقامة السعودي: ١٠ أرقام. نتساهل بقبول ٩ أرقام احتياطًا لبعض
// التنسيقات، ونتجاهل أي فواصل أو مسافات داخل الرقم قبل الفحص.
function looksLikeNationalId(s) {
  const digits = String(s ?? "").replace(/[^\d]/g, "");
  return /^\d{9,10}$/.test(digits) ? digits : null;
}
// يحوّل نصًا ملصوقًا من جدول "الطلاب" بنظام نور (بعد نسخه كاملًا بالمتصفح)
// إلى مصفوفة {name, nationalId} — يدعم كلًا من اللصق الحقيقي كجدول (خلايا
// مفصولة بمسافة Tab) ولصق قائمة أسماء بسيطة (اسم واحد بكل سطر).
function parseNoorStudentTable(text) {
  const lines = String(text ?? "").split("\n");
  const results = [];
  const seen = new Set();
  lines.forEach((line) => {
    const raw = line.trim();
    if (!raw) return;
    const parts = raw.includes("\t") ? raw.split("\t") : raw.split(/\s{2,}/);
    let name = null, nationalId = null;
    parts.forEach((p) => {
      const pt = p.trim();
      if (!pt) return;
      if (!nationalId) {
        const id = looksLikeNationalId(pt);
        if (id) { nationalId = id; return; }
      }
      if (!name && looksLikeFullName(pt)) name = pt;
    });
    if (!name && parts.length === 1 && looksLikeFullName(raw)) name = raw;
    if (name && !seen.has(name)) {
      seen.add(name);
      results.push({ name, nationalId: nationalId || "" });
    }
  });
  return results;
}

function RowModal({ initial, onClose, onSaveMany, onSaveOne, onDelete, showRowNumbers, onToggleShowRowNumbers, isOwner }) {
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
  const [showNoorEmbed, setShowNoorEmbed] = useState(false);

  const validSingle = single && single.name.trim();
  const validDrafts = drafts.filter((d) => d.name.trim());
  const bulkNames = parsePastedNames(bulkText);

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
        const names = extractNamesSmart(rows);
        if (names.length === 0) setImportError("لم يتم العثور على أي أسماء بهذا الملف. تأكد إنه يحتوي عمودًا فيه أسماء الطلاب.");
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
            <button onClick={() => setTab("bulk")} className="px-3 py-1.5 rounded-full text-xs font-semibold flex items-center gap-1"
              style={{ background: tab === "bulk" ? INK : "transparent", color: tab === "bulk" ? "#fff" : MUTED, border: `1px solid ${tab === "bulk" ? INK : LINE}` }}>
              <ExternalLink size={13} /> استيراد الأسماء من نور
            </button>
            <button onClick={() => setTab("import")} className="px-3 py-1.5 rounded-full text-xs font-semibold flex items-center gap-1"
              style={{ background: tab === "import" ? INK : "transparent", color: tab === "import" ? "#fff" : MUTED, border: `1px solid ${tab === "import" ? INK : LINE}` }}>
              <FileSpreadsheet size={13} /> استيراد من Excel / نور
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
              <button onClick={() => setDrafts([...drafts, emptyRowDraft()])} className="mb-2 text-sm font-semibold flex items-center gap-1" style={{ color: "#26423B" }}>
                <Plus size={15} /> إضافة صف آخر
              </button>
            </>
          ) : tab === "bulk" ? (
            <>
              <button
                onClick={() => setShowNoorEmbed(true)}
                className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white mb-3 transition-all hover:brightness-110 active:scale-95 magic-shimmer"
                style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})`, backgroundSize: "220% 220%" }}
              >
                <ExternalLink size={16} /> استيراد الأسماء من نور تلقائيًا
              </button>
              <p className="text-xs text-center mb-3" style={{ color: MUTED }}>أو ألصق الأسماء يدويًا بالمربع تحت</p>
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-sm font-medium" style={{ color: INK }}>أسماء الصفوف (كل اسم في سطر)</span>
                {bulkText.trim() && (
                  <button
                    type="button"
                    onClick={() => setBulkText("")}
                    className="flex items-center gap-1 text-xs font-semibold px-2 py-1 rounded-lg hover:bg-black/5"
                    style={{ color: "#9A3B2E" }}
                  >
                    <Eraser size={12} /> تفريغ الحقل
                  </button>
                )}
              </div>
              <div className="mb-4">
                <textarea
                  style={{ ...inputStyle, minHeight: "140px", resize: "vertical" }}
                  value={bulkText}
                  onChange={(e) => setBulkText(e.target.value)}
                  placeholder={"محمد أحمد\nسارة خالد\nعبدالله فهد"}
                />
              </div>
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
              <div className="p-3 rounded-xl mb-3 flex items-start gap-2" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
                <Sparkles size={15} color="#26423B" className="shrink-0 mt-0.5" />
                <div className="text-xs" style={{ color: "#26423B" }}>
                  <p className="font-bold mb-1">تستورد من نظام نور؟</p>
                  <p>من نور: افتح صفحة الطلاب ← اضغط "تصدير" أو "طباعة" واختر Excel ← ارفع الملف هنا مباشرة، حتى لو فيه أعمدة إضافية (الرقم، الهوية...) — نكتشف عمود الاسم تلقائيًا.</p>
                </div>
              </div>
              <Field label="ملف Excel أو CSV بأسماء الطلاب">
                <input ref={importInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} style={{ display: "none" }} />
                <button type="button" onClick={() => importInputRef.current?.click()} className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold" style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}>
                  <FileSpreadsheet size={15} color="#26423B" /> اختر ملفًا
                </button>
                {importFileName && <p className="text-xs mt-1.5" style={{ color: MUTED }}>{importFileName}</p>}
              </Field>
              {importError && <p className="text-xs mb-3 font-medium" style={{ color: "#C0392B" }}>{importError}</p>}
              {importNames.length > 0 && (
                <div className="mb-3 p-2 rounded-lg" style={{ border: `1px solid ${LINE}`, maxHeight: 150, overflowY: "auto", background: "#F8F7F2" }}>
                  <p className="text-xs font-bold mb-1" style={{ color: "#26423B" }}>تم العثور على {importNames.length} اسم:</p>
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
            style={{ background: "#26423B" }}
          >تم</button>
        </div>
      </div>
      {showNoorEmbed && (
        <NoorEmbedModal
          isOwner={isOwner}
          onClose={() => setShowNoorEmbed(false)}
          onImportNames={(rows) => {
            setShowNoorEmbed(false);
            onSaveMany(rows.map((r) => ({
              name: r.name,
              type: "text",
              options: [],
              color: bulkColor,
              autoRenew: bulkAutoRenew,
              medicalNote: bulkMedicalNote.trim(),
              nationalId: r.nationalId || "",
            })));
          }}
        />
      )}
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
          style={{ background: "#26423B" }}
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
            style={{ background: "#26423B" }}
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
              style={{ border: `1px solid ${em === emoji ? "#26423B" : LINE}`, background: em === emoji ? "#EAF3F0" : "#fff" }}
            >{em}</button>
          ))}
        </div>
        <input style={{ ...inputStyle, width: 90 }} value={emoji} onChange={(e) => setEmoji(e.target.value.slice(0, 2))} placeholder="أو اكتب رمزًا" />
      </Field>
      <div className="flex justify-end gap-2 mt-5">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button disabled={!valid} onClick={() => onSave({ subject: subject.trim(), grade: grade.trim(), teacher: teacher.trim(), yearHijri: yearH.trim(), yearGregorian: yearG.trim(), color, emoji })}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40" style={{ background: "#26423B" }}>تم</button>
      </div>
    </Modal>
  );
}

// ---------- Dropdown cell with per-option colors ----------

function emptyTestQuestion() {
  return { id: uid(), text: "", options: [{ id: uid(), text: "" }, { id: uid(), text: "" }], correctOptionId: null };
}

function TestBuilderModal({ onSave, onClose, questionBank = [], onAddToBank }) {
  const [title, setTitle] = useState("");
  const [questions, setQuestions] = useState([emptyTestQuestion()]);
  const [saveToBankIds, setSaveToBankIds] = useState(() => new Set());
  const [showBankPicker, setShowBankPicker] = useState(false);
  const [bankSearch, setBankSearch] = useState("");

  const updateQuestion = (qid, patch) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, ...patch } : q)));
  const updateOption = (qid, oid, text) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, options: q.options.map((o) => (o.id === oid ? { ...o, text } : o)) } : q)));
  const addOption = (qid) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, options: [...q.options, { id: uid(), text: "" }] } : q)));
  const removeOption = (qid, oid) => setQuestions((qs) => qs.map((q) => (q.id === qid ? { ...q, options: q.options.filter((o) => o.id !== oid), correctOptionId: q.correctOptionId === oid ? null : q.correctOptionId } : q)));
  const addQuestion = () => setQuestions((qs) => [...qs, emptyTestQuestion()]);
  const removeQuestion = (qid) => setQuestions((qs) => qs.filter((q) => q.id !== qid));
  const toggleSaveToBank = (qid) => setSaveToBankIds((s) => { const n = new Set(s); if (n.has(qid)) n.delete(qid); else n.add(qid); return n; });

  const insertFromBank = (bankQ) => {
    const newOptions = bankQ.options.map((o) => ({ id: uid(), text: o.text, _origId: o.id }));
    const newCorrect = newOptions.find((o) => o._origId === bankQ.correctOptionId)?.id || null;
    const cleanOptions = newOptions.map(({ id, text }) => ({ id, text }));
    setQuestions((qs) => [...qs, { id: uid(), text: bankQ.text, options: cleanOptions, correctOptionId: newCorrect }]);
  };

  const filteredBank = questionBank.filter((q) => q.text.toLowerCase().includes(bankSearch.trim().toLowerCase()));

  const valid = title.trim() && questions.length > 0 && questions.every((q) => q.text.trim() && q.options.length >= 2 && q.options.every((o) => o.text.trim()) && q.correctOptionId);

  return (
    <Modal title="إنشاء اختبار جديد" onClose={onClose} wide>
      <Field label="عنوان الاختبار"><input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="مثال: اختبار الوحدة الأولى" /></Field>

      {questionBank.length > 0 && (
        <div className="mb-4">
          <button onClick={() => setShowBankPicker((s) => !s)} className="text-sm font-semibold flex items-center gap-1.5 mb-2" style={{ color: "#26423B" }}>
            <BookMarked size={15} /> إضافة من بنك الأسئلة ({questionBank.length})
          </button>
          {showBankPicker && (
            <div className="p-3 rounded-xl mb-2" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
              <input value={bankSearch} onChange={(e) => setBankSearch(e.target.value)} style={{ ...inputStyle, marginBottom: 8 }} placeholder="ابحث بنص السؤال..." />
              <div className="space-y-1.5 max-h-52 overflow-y-auto">
                {filteredBank.length === 0 ? (
                  <p className="text-xs text-center py-4" style={{ color: MUTED }}>لا يوجد أسئلة مطابقة.</p>
                ) : filteredBank.map((q) => (
                  <div key={q.id} className="flex items-center gap-2 p-2 rounded-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                    <span className="flex-1 text-xs truncate" style={{ color: INK }}>{q.text}</span>
                    <button onClick={() => insertFromBank(q)} className="text-xs font-bold px-2.5 py-1 rounded-lg text-white shrink-0" style={{ background: "#26423B" }}>إضافة</button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      )}

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
              <button onClick={() => addOption(q.id)} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}><Plus size={12} /> إضافة خيار</button>
              <label className="flex items-center gap-1.5 text-xs mt-1.5" style={{ color: MUTED }}>
                <input type="checkbox" checked={saveToBankIds.has(q.id)} onChange={() => toggleSaveToBank(q.id)} />
                <BookMarked size={12} /> احفظ هذا السؤال ببنك الأسئلة لإعادة استخدامه لاحقًا
              </label>
            </div>
          </div>
        ))}
      </div>
      <button onClick={addQuestion} className="mb-4 text-sm font-semibold flex items-center gap-1" style={{ color: "#26423B" }}><Plus size={15} /> إضافة سؤال</button>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
        <button
          disabled={!valid}
          onClick={() => {
            const toBank = questions.filter((q) => saveToBankIds.has(q.id));
            if (toBank.length > 0 && onAddToBank) onAddToBank(toBank);
            onSave({ id: uid(), title: title.trim(), questions, results: [], createdAt: todayKey() });
          }}
          className="px-5 py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
          style={{ background: "#26423B" }}
        >حفظ الاختبار</button>
      </div>
    </Modal>
  );
}

// تحليل الأسئلة بعد التصحيح: يوري أي سؤال أخطأ فيه أكبر نسبة من الطلاب —
// يساعد المعلم يعرف بالضبط أي جزء من الدرس يحتاج إعادة شرح.
function TestAnalysisModal({ test, onClose }) {
  const results = test.results || [];
  const analysis = test.questions.map((q, qi) => {
    const total = results.length;
    const wrong = results.filter((r) => r.answers?.[q.id] !== q.correctOptionId).length;
    const pct = total > 0 ? Math.round((wrong / total) * 100) : 0;
    return { q, qi, wrong, total, pct };
  }).sort((a, b) => b.pct - a.pct);

  return (
    <Modal title={`تحليل الأسئلة — ${test.title}`} onClose={onClose} wide>
      <p className="text-xs mb-4" style={{ color: MUTED }}>{results.length} طالب مصحّح — الأسئلة مرتّبة من الأصعب (أعلى نسبة خطأ) للأسهل.</p>
      {results.length === 0 ? (
        <p className="text-sm text-center py-8" style={{ color: MUTED }}>لا يوجد نتائج مصحّحة بعد لهذا الاختبار.</p>
      ) : (
        <div className="space-y-2">
          {analysis.map(({ q, qi, wrong, total, pct }) => (
            <div key={q.id} className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
              <div className="flex items-center justify-between gap-2 mb-1.5">
                <p className="text-sm font-semibold flex-1" style={{ color: INK }}>س{qi + 1}. {q.text}</p>
                <span className="text-xs font-bold px-2 py-1 rounded-full shrink-0" style={{ background: pct >= 50 ? "#FBEAE7" : pct >= 25 ? "#FCEFE2" : "#E3F1EC", color: pct >= 50 ? "#C0392B" : pct >= 25 ? "#C97A2B" : "#0F9D58" }}>
                  {wrong} من {total} أخطأوا ({pct}٪)
                </span>
              </div>
              <div className="w-full h-1.5 rounded-full overflow-hidden" style={{ background: "#F0EFE9" }}>
                <div className="h-full rounded-full" style={{ width: `${pct}%`, background: pct >= 50 ? "#C0392B" : pct >= 25 ? "#C97A2B" : "#0F9D58" }} />
              </div>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function TestsListModal({ tests, onCreateNew, onGrade, onGradeCamera, onPrint, onDelete, onArchive, onAnalyze, onClose, bare = false }) {
  const content = (
    <>
      <button onClick={onCreateNew} className="mb-4 flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold text-white" style={{ background: "#26423B" }}>
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
              {(t.results || []).length > 0 && (
                <button onClick={() => onAnalyze(t.id)} title="تحليل الأسئلة" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><BarChart3 size={15} color={MUTED} /></button>
              )}
              <button onClick={() => onPrint(t.id)} title="طباعة الورقة" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Printer size={15} color={MUTED} /></button>
              <button onClick={() => onGradeCamera(t.id)} title="تصحيح بالكاميرا" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Camera size={15} color={MUTED} /></button>
              <button onClick={() => onGrade(t.id)} className="px-3 py-1.5 rounded-lg text-xs font-bold text-white shrink-0" style={{ background: "#26423B" }}>تصحيح</button>
              <button onClick={() => onArchive(t.id)} title="أرشفة" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Archive size={14} color={MUTED} /></button>
              <button onClick={() => onDelete(t.id)} className="p-1.5 rounded hover:bg-black/5 shrink-0"><Trash2 size={14} color="#C0392B" /></button>
            </div>
          ))}
        </div>
      )}
    </>
  );
  if (bare) return content;
  return (
    <Modal title="الاختبارات" onClose={onClose} wide>
      {content}
    </Modal>
  );
}

function AnswerOption({ label, selected, onSelect, color }) {
  const c = color || "#26423B";
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
    ctx.drawImage(logoImageElement, width - 110, 18, 56, 56);
  }
  if (countryName) { ctx.font = "bold 14px Tahoma, Arial"; ctx.fillStyle = "#232622"; ctx.fillText(countryName, width / 2, hy); hy += 18; }
  if (ministryName) { ctx.font = "13px Tahoma, Arial"; ctx.fillStyle = "#7A7768"; ctx.fillText(ministryName, width / 2, hy); hy += 18; }
  if (schoolName) { ctx.font = "14px Tahoma, Arial"; ctx.fillStyle = "#232622"; ctx.fillText(schoolName, width / 2, hy); hy += 24; }

  ctx.font = "bold 22px Tahoma, Arial";
  ctx.fillStyle = "#26423B";
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
    ctx.fillStyle = "#26423B";
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
  const [shuffleQuestions, setShuffleQuestions] = useState(false);

  const buildSheet = (studentName, index) => {
    const questions = shuffleQuestions && !omrFormat ? shuffleArr(test.questions) : test.questions;
    return omrFormat
      ? buildOMRSheetCanvas({ title: test.title, studentName, questionCount: test.questions.length, optionCount, studentIndex: index })
      : buildTestPaperCanvas({ title: test.title, studentName, questions });
  };

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
        <button onClick={downloadBlank} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white" style={{ background: "#26423B" }}>
          <ImageDown size={16} /> تنزيل نسخة فارغة
        </button>
      ) : (
        <>
          <Field label="الفصل">
            <select value={classId} onChange={(e) => setClassId(e.target.value)} style={inputStyle}>
              {classes.filter((c) => !c.archived).map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade} ({c.rows.length} طالب)</option>)}
            </select>
          </Field>
          {!omrFormat && (
            <label className="flex items-start gap-2 text-sm font-medium mb-3" style={{ color: INK }}>
              <input type="checkbox" checked={shuffleQuestions} onChange={(e) => setShuffleQuestions(e.target.checked)} className="mt-0.5" />
              <span>
                ترتيب عشوائي للأسئلة لكل طالب
                <span className="block text-xs font-normal mt-0.5" style={{ color: MUTED }}>يقلّل النسخ بين الطلاب المتجاورين بالجلوس — كل ورقة يصير ترتيب أسئلتها مختلف.</span>
              </span>
            </label>
          )}
          {generating ? (
            <div className="p-3 rounded-xl flex items-center gap-2" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
              <div className="w-4 h-4 rounded-full border-2 shrink-0" style={{ borderColor: "#26423B transparent #26423B #26423B", animation: "spin 0.8s linear infinite" }} />
              <p className="text-xs font-semibold" style={{ color: INK }}>جارٍ التجهيز... {progress}%</p>
            </div>
          ) : (
            <button onClick={downloadPerStudent} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white" style={{ background: "#26423B" }}>
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
          <button disabled={!colId} onClick={startCamera} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white disabled:opacity-40 mt-2" style={{ background: "#26423B" }}>
            <Camera size={16} /> بدء التصحيح
          </button>
          {cameraError && <p className="text-xs mt-2 text-center" style={{ color: "#C0392B" }}>{cameraError}</p>}
        </>
      )}

      {step === "camera" && (
        <>
          {postedCount > 0 && <p className="text-xs text-center mb-2 font-semibold" style={{ color: "#26423B" }}>تم رصد {postedCount} ورقة حتى الآن</p>}
          <p className="text-sm mb-3 text-center" style={{ color: MUTED }}>
            حاذِ إطار الورقة بالكامل مع حدود المستطيل الظاهر، بإضاءة جيدة، ثم التقط.
          </p>
          <div className="relative mx-auto mb-4 rounded-xl overflow-hidden" style={{ width: "100%", maxWidth: 340, aspectRatio: `${OMR_SHEET_W} / ${OMR_SHEET_H}`, background: "#000" }}>
            <video ref={videoRef} autoPlay playsInline muted style={{ width: "100%", height: "100%", objectFit: "cover" }} />
            <div className="absolute inset-2 pointer-events-none rounded-lg" style={{ border: "3px solid #F2C94C" }} />
          </div>
          <div className="flex gap-2">
            <button onClick={() => { stopCamera(); setStep("setup"); }} className="flex-1 py-2.5 rounded-xl text-sm font-medium" style={{ color: MUTED, border: `1px solid ${LINE}` }}>إنهاء</button>
            <button onClick={capture} className="flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-bold text-white" style={{ background: "#26423B" }}>
              <Camera size={16} /> التقط
            </button>
          </div>
        </>
      )}

      {step === "result" && scanResult && cls && (
        <>
          <div className="text-center p-6 rounded-2xl mb-4" style={{ background: "#EAF3F0" }}>
            <p className="text-3xl font-extrabold" style={{ color: "#26423B" }}>{scanResult.score} / {scanResult.total}</p>
            <p className="text-sm mt-1" style={{ color: "#26423B" }}>{scanResult.percentage}%</p>
          </div>
          <Field label="الطالب المكتشف من الورقة" hint={scanResult.unclear ? "تعذّرت قراءة بعض الإجابات بوضوح — تأكد من الدرجة أعلاه." : "تأكد أن الاسم صحيح قبل الرصد؛ صحّحه من القائمة إن لزم."}>
            <select value={manualRowId} onChange={(e) => setManualRowId(e.target.value)} style={inputStyle}>
              <option value="">— لم يُتعرّف على طالب مطابق —</option>
              {cls.rows.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
            </select>
          </Field>
          <div className="flex gap-2 mt-2">
            <button onClick={() => { setStep("camera"); startCamera(); }} className="flex-1 py-2.5 rounded-xl text-sm font-medium" style={{ color: MUTED, border: `1px solid ${LINE}` }}>إعادة التصوير</button>
            <button disabled={!manualRowId} onClick={recordAndScanNext} className="flex-1 py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40" style={{ background: "#26423B" }}>
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
            style={{ background: "#26423B" }}
          >تصحيح الآن</button>
        </>
      ) : (
        <>
          <div className="text-center p-6 rounded-2xl mb-4" style={{ background: "#EAF3F0" }}>
            <p className="text-3xl font-extrabold" style={{ color: "#26423B" }}>{graded.score} / {graded.total}</p>
            <p className="text-sm mt-1" style={{ color: "#26423B" }}>{graded.percentage}% — {graded.studentName}</p>
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
                style={{ background: "#26423B" }}
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
          style={{ background: "#26423B" }}
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
          style={{ background: "#26423B" }}
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
            style={{ background: "#26423B" }}
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

function GradeSheetClassPicker({ classes, onSelect, onClose }) {
  const activeClasses = classes.filter((c) => !c.archived);
  return (
    <Modal title="كشف رصد درجات — اختر الفصل" onClose={onClose} accent="magic">
      {activeClasses.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد فصول بعد.</p>
      ) : (
        <div className="space-y-2">
          {activeClasses.map((c) => (
            <button
              key={c.id}
              onClick={() => onSelect(c.id)}
              className="w-full flex items-center gap-3 p-3 rounded-xl text-right hover:bg-black/5"
              style={{ border: `1px solid ${LINE}`, background: "#fff" }}
            >
              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: c.color }} />
              <div className="flex-1 min-w-0">
                <p className="text-sm font-semibold truncate" style={{ color: INK }}>{c.subject}</p>
                <p className="text-xs" style={{ color: MUTED }}>{c.grade} • {c.teacher}</p>
              </div>
              <ChevronLeft size={16} color={MUTED} />
            </button>
          ))}
        </div>
      )}
    </Modal>
  );
}

function GradeSheetModal({ cls, onClose, onBack, onGenerate }) {
  const gradeColumns = cls.columns.filter((c) => c.type === "counter");
  const [shortTestIds, setShortTestIds] = useState([]);
  const [finalExamIds, setFinalExamIds] = useState([]);
  const [reviewerName, setReviewerName] = useState("");

  const toggleShortTest = (id) => {
    setShortTestIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  };
  const toggleFinalExam = (id) => {
    setFinalExamIds((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));
  };

  return (
    <Modal title="كشف رصد درجات المادة" onClose={onClose} onBack={onBack} accent="magic">
      <p className="text-xs mb-4" style={{ color: MUTED }}>
        هذا الكشف يتبع نفس التنسيق المعتمد بنظام نور: معدل الاختبارات القصيرة، اختبار نهاية الفصل، المجموع، وخانة توقيع المراجع — جاهز للتسليم مباشرة.
      </p>
      {gradeColumns.length === 0 ? (
        <p className="text-sm text-center py-6" style={{ color: MUTED }}>لا يوجد أعمدة من نوع "عداد" بهذا الفصل بعد — أضف أعمدة الدرجات أولًا.</p>
      ) : (
        <>
          <Field label="أعمدة الاختبارات القصيرة (يُحسب معدلها)" hint="اختر كل الأعمدة اللي تمثّل اختبارات قصيرة متكررة.">
            <div className="space-y-1.5">
              {gradeColumns.map((c) => (
                <label key={c.id} className="flex items-center gap-2 text-sm p-2 rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                  <input type="checkbox" checked={shortTestIds.includes(c.id)} onChange={() => toggleShortTest(c.id)} />
                  <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: c.color }} />
                  {c.name}
                </label>
              ))}
            </div>
          </Field>
          <Field label="أعمدة اختبار نهاية الفصل (يُحسب معدلها)" hint="تقدر تختار أكثر من عمود — يُحسب معدلها مثل الاختبارات القصيرة.">
            <div className="space-y-1.5">
              {gradeColumns.map((c) => (
                <label key={c.id} className="flex items-center gap-2 text-sm p-2 rounded-lg" style={{ border: `1px solid ${LINE}` }}>
                  <input type="checkbox" checked={finalExamIds.includes(c.id)} onChange={() => toggleFinalExam(c.id)} />
                  <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: c.color }} />
                  {c.name}
                </label>
              ))}
            </div>
          </Field>
          <Field label="اسم المراجع (اختياري)">
            <input value={reviewerName} onChange={(e) => setReviewerName(e.target.value)} style={inputStyle} placeholder="مثال: أ. محمد العتيبي" />
          </Field>
          <button
            disabled={shortTestIds.length === 0 && finalExamIds.length === 0}
            onClick={() => onGenerate({ shortTestIds, finalExamIds, reviewerName: reviewerName.trim() })}
            className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
            style={{ background: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)" }}
          >
            إنشاء الكشف
          </button>
        </>
      )}
    </Modal>
  );
}

// أداة تصدير الدرجات إلى نور: نور لا يسمح بتعبئة تلقائية من مواقع خارجية
// لأسباب أمنية (نفس السبب اللي يمنع تضمينه بإطار)، فهذه الأداة تجهّز قائمة
// الدرجات بترتيب الطلاب الحالي (المطابق لترتيبهم لو استُوردوا من نور
// أصلًا) جاهزة للنسخ، وتبقى خطوة اللصق داخل صفحة "رصد الدرجات" بنور يدويًا.
function computeColumnPeriodAverage(cls, rowId, colId, from, to) {
  const list = cls.reports?.[rowId] || [];
  const nums = list
    .filter((e) => e.colId === colId && e.dateKey && (!from || e.dateKey >= from) && (!to || e.dateKey <= to))
    .map((e) => Number(e.value))
    .filter((n) => !Number.isNaN(n));
  if (nums.length === 0) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

// مقارنة أداء الطلاب بعمود درجات (نوع "عداد") بين فترتين زمنيتين — يوري
// مين تحسّن ومين تراجع، مفيد لاجتماعات المتابعة وتقارير أولياء الأمور.
function PeriodComparisonModal({ cls, onClose, onPrint }) {
  const gradeColumns = cls.columns.filter((c) => c.type === "counter");
  const [colId, setColId] = useState(gradeColumns[0]?.id || "");
  const [from1, setFrom1] = useState("");
  const [to1, setTo1] = useState("");
  const [from2, setFrom2] = useState("");
  const [to2, setTo2] = useState("");
  const [computed, setComputed] = useState(null);

  const col = cls.columns.find((c) => c.id === colId);

  const runCompare = () => {
    const rows = cls.rows.map((row) => {
      const avg1 = computeColumnPeriodAverage(cls, row.id, colId, from1, to1);
      const avg2 = computeColumnPeriodAverage(cls, row.id, colId, from2, to2);
      const diff = avg1 !== null && avg2 !== null ? avg2 - avg1 : null;
      return { row, avg1, avg2, diff };
    }).sort((a, b) => {
      if (a.diff === null && b.diff === null) return 0;
      if (a.diff === null) return 1;
      if (b.diff === null) return -1;
      return b.diff - a.diff;
    });
    const improved = rows.filter((r) => r.diff !== null && r.diff > 0).length;
    const declined = rows.filter((r) => r.diff !== null && r.diff < 0).length;
    const same = rows.filter((r) => r.diff === 0).length;
    setComputed({ rows, improved, declined, same });
  };

  return (
    <Modal title="مقارنة أداء بين فترتين" onClose={onClose} accent="magic" wide>
      {gradeColumns.length === 0 ? (
        <p className="text-sm text-center py-8" style={{ color: MUTED }}>لا يوجد أعمدة من نوع "عداد" بهذا الفصل بعد.</p>
      ) : (
        <>
          <Field label="العمود (الدرجة) المراد مقارنته">
            <select value={colId} onChange={(e) => { setColId(e.target.value); setComputed(null); }} style={inputStyle}>
              {gradeColumns.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-2">
            <div className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
              <p className="text-xs font-bold mb-2" style={{ color: INK }}>الفترة الأولى</p>
              <div className="grid grid-cols-2 gap-2">
                <input type="date" value={from1} onChange={(e) => setFrom1(e.target.value)} style={inputStyle} />
                <input type="date" value={to1} onChange={(e) => setTo1(e.target.value)} style={inputStyle} />
              </div>
            </div>
            <div className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
              <p className="text-xs font-bold mb-2" style={{ color: INK }}>الفترة الثانية</p>
              <div className="grid grid-cols-2 gap-2">
                <input type="date" value={from2} onChange={(e) => setFrom2(e.target.value)} style={inputStyle} />
                <input type="date" value={to2} onChange={(e) => setTo2(e.target.value)} style={inputStyle} />
              </div>
            </div>
          </div>
          <button
            disabled={!from1 || !to1 || !from2 || !to2}
            onClick={runCompare}
            className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110 mb-4"
            style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
          >
            قارن الآن
          </button>

          {computed && (
            <>
              <div className="flex flex-wrap gap-2 mb-3">
                <span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: "#E3F1EC", color: "#0F9D58" }}>تحسّن: {computed.improved}</span>
                <span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: "#FBEAE7", color: "#C0392B" }}>تراجع: {computed.declined}</span>
                <span className="text-xs font-bold px-2.5 py-1 rounded-full" style={{ background: "#F3F1E9", color: MUTED }}>ثابت/بلا بيانات: {computed.rows.length - computed.improved - computed.declined}</span>
                <IconBtn icon={Printer} label="طباعة / تصدير" onClick={() => onPrint({ colName: col?.name, from1, to1, from2, to2, rows: computed.rows })} />
              </div>
              <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}`, maxHeight: 360, overflowY: "auto" }}>
                <table className="w-full border-collapse text-sm">
                  <thead>
                    <tr>
                      <th className="p-2 text-right" style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, position: "sticky", top: 0 }}>الاسم</th>
                      <th className="p-2 text-center" style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, position: "sticky", top: 0 }}>الفترة الأولى</th>
                      <th className="p-2 text-center" style={{ background: "#F3F1E9", border: `1px solid ${LINE}`, position: "sticky", top: 0 }}>الفترة الثانية</th>
                      <th className="p-2 text-center" style={{ background: GOLD_LIGHT, border: `1px solid ${LINE}`, position: "sticky", top: 0 }}>الفرق</th>
                    </tr>
                  </thead>
                  <tbody>
                    {computed.rows.map((r, i) => {
                      const color = r.diff === null ? MUTED : r.diff > 0 ? "#0F9D58" : r.diff < 0 ? "#C0392B" : MUTED;
                      const arrow = r.diff === null ? "—" : r.diff > 0 ? "▲" : r.diff < 0 ? "▼" : "＝";
                      return (
                        <tr key={r.row.id} style={{ background: i % 2 ? "#FBFAF6" : "#fff" }}>
                          <td className="p-2" style={{ border: `1px solid ${LINE}`, color: INK }}>{r.row.name}</td>
                          <td className="p-2 text-center" style={{ border: `1px solid ${LINE}`, color: MUTED }}>{r.avg1 === null ? "—" : Math.round(r.avg1 * 100) / 100}</td>
                          <td className="p-2 text-center" style={{ border: `1px solid ${LINE}`, color: MUTED }}>{r.avg2 === null ? "—" : Math.round(r.avg2 * 100) / 100}</td>
                          <td className="p-2 text-center font-bold" style={{ border: `1px solid ${LINE}`, color }}>
                            {arrow} {r.diff === null ? "" : Math.abs(Math.round(r.diff * 100) / 100)}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </>
      )}
    </Modal>
  );
}

// محرر قائمة نصوص بسيط (إضافة/حذف/تعديل سطر) — يُستخدم للأهداف وخطوات
// التنفيذ بتقرير البرنامج التفصيلي.
function TextListEditor({ items, onChange, placeholder }) {
  const update = (i, val) => onChange(items.map((it, ix) => (ix === i ? val : it)));
  const remove = (i) => onChange(items.filter((_, ix) => ix !== i));
  const add = () => onChange([...items, ""]);
  return (
    <div className="space-y-1.5">
      {items.map((it, i) => (
        <div key={i} className="flex items-center gap-1.5">
          <span className="text-xs font-bold w-4 shrink-0" style={{ color: MUTED }}>{i + 1}.</span>
          <input value={it} onChange={(e) => update(i, e.target.value)} style={{ ...inputStyle, flex: 1, padding: "6px 10px" }} placeholder={placeholder} />
          {items.length > 1 && <button type="button" onClick={() => remove(i)} className="p-1 rounded hover:bg-black/5 shrink-0"><X size={12} color={MUTED} /></button>}
        </div>
      ))}
      <button type="button" onClick={add} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}><Plus size={12} /> إضافة سطر</button>
    </div>
  );
}

function ShawahedCategoryModal({ category, entries, onAdd, onEdit, onDelete, onArchive, onMove, onShareReadOnlyEntry, onPrintProgram, onClose }) {
  const [title, setTitle] = useState("");
  const [notes, setNotes] = useState("");
  const [photo, setPhoto] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [expandedEntryId, setExpandedEntryId] = useState(null);
  const attachInputRef = useRef(null);
  const [showProgram, setShowProgram] = useState(false);
  const [programField, setProgramField] = useState("");
  const [implementers, setImplementers] = useState("");
  const [executionDate, setExecutionDate] = useState("");
  const [beneficiaries, setBeneficiaries] = useState("جميع الطلاب");
  const [beneficiaryCount, setBeneficiaryCount] = useState("");
  const [objectives, setObjectives] = useState([""]);
  const [steps, setSteps] = useState([""]);
  const [extraPhotos, setExtraPhotos] = useState([]);
  const photoInputRef = useRef(null);
  const extraPhotosInputRef = useRef(null);

  const handlePhotoChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setPhoto(reader.result);
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const handleExtraPhotosChange = (e) => {
    const files = Array.from(e.target.files || []).slice(0, 4 - extraPhotos.length);
    files.forEach((file) => {
      const reader = new FileReader();
      reader.onload = () => setExtraPhotos((prev) => [...prev, reader.result].slice(0, 4));
      reader.readAsDataURL(file);
    });
    e.target.value = "";
  };

  // إرفاق ملف مباشرة على شاهد موجود (بدون الحاجة لفتح وضع التعديل) — يتصرف
  // مثل مجلد صغير لكل شاهد تقدر تحط فيه أي ملف (وورد، PDF، صور...) وتشاركه
  // لاحقًا زي ملفات Google Drive.
  const handleAttachFile = (entryId, e) => {
    const files = Array.from(e.target.files || []);
    files.forEach((file) => {
      if (file.size > 15 * 1024 * 1024) {
        alert(`الملف "${file.name}" كبير جدًا (أكبر من ١٥ ميجا) — يُفضّل ملفات أصغر.`);
        return;
      }
      const reader = new FileReader();
      reader.onload = () => {
        const entry = entries.find((x) => x.id === entryId);
        if (!entry) return;
        const newAttachment = { id: uid(), name: file.name, mimeType: file.type, dataUrl: reader.result, uploadedAt: todayKey() };
        onEdit(entryId, { attachments: [...(entry.attachments || []), newAttachment] });
      };
      reader.readAsDataURL(file);
    });
    e.target.value = "";
  };

  const handleRemoveAttachment = (entryId, attachmentId) => {
    const entry = entries.find((x) => x.id === entryId);
    if (!entry) return;
    onEdit(entryId, { attachments: (entry.attachments || []).filter((a) => a.id !== attachmentId) });
  };

  // مشاركة "خام" (الملف/الصورة نفسها عبر قائمة المشاركة بالجهاز) — مختلفة
  // عن "مشاركة للقراءة فقط" اللي تولّد صفحة تقرير منسّقة بدل الملف الخام.
  const shareEntryRaw = async (entry) => {
    try {
      if (entry.photo) {
        const blob = await (await fetch(entry.photo)).blob();
        await shareOrDownloadFile(blob, `${entry.title}.png`, blob.type || "image/png");
      } else if (entry.attachments && entry.attachments.length > 0) {
        const att = entry.attachments[0];
        const blob = await (await fetch(att.dataUrl)).blob();
        await shareOrDownloadFile(blob, att.name, att.mimeType || blob.type);
      } else {
        onShareReadOnlyEntry(entry);
      }
    } catch (err) {
      onShareReadOnlyEntry(entry);
    }
  };

  // رفع تقرير/ملف جاهز مباشرة كشاهد جديد بضغطة واحدة — بدل تعبئة النموذج
  // يدويًا، تختار الملف وهو يُنشئ الشاهد ويُرفقه تلقائيًا.
  const quickReportInputRef = useRef(null);
  const handleQuickReportUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 15 * 1024 * 1024) {
      alert("الملف كبير جدًا (أكبر من ١٥ ميجا) — يُفضّل ملف أصغر.");
      e.target.value = "";
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const attachment = { id: uid(), name: file.name, mimeType: file.type, dataUrl: reader.result, uploadedAt: todayKey() };
      onAdd({ id: uid(), title: file.name.replace(/\.[^.]+$/, ""), notes: "", photo: null, date: todayKey(), attachments: [attachment] });
    };
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const resetForm = () => {
    setTitle(""); setNotes(""); setPhoto(null); setEditingId(null);
    setShowProgram(false); setProgramField(""); setImplementers(""); setExecutionDate("");
    setBeneficiaries("جميع الطلاب"); setBeneficiaryCount(""); setObjectives([""]); setSteps([""]); setExtraPhotos([]);
  };

  const startEdit = (entry) => {
    setEditingId(entry.id);
    setTitle(entry.title);
    setNotes(entry.notes || "");
    setPhoto(entry.photo || null);
    if (entry.program) {
      setShowProgram(true);
      setProgramField(entry.program.programField || "");
      setImplementers(entry.program.implementers || "");
      setExecutionDate(entry.program.executionDate || "");
      setBeneficiaries(entry.program.beneficiaries || "جميع الطلاب");
      setBeneficiaryCount(entry.program.beneficiaryCount || "");
      setObjectives(entry.program.objectives?.length ? entry.program.objectives : [""]);
      setSteps(entry.program.steps?.length ? entry.program.steps : [""]);
      setExtraPhotos(entry.program.photos || []);
    } else {
      setShowProgram(false);
    }
  };

  const submit = () => {
    if (!title.trim()) return;
    const program = showProgram ? {
      programName: title.trim(),
      programField: programField.trim(),
      implementers: implementers.trim(),
      executionDate: executionDate.trim(),
      beneficiaries: beneficiaries.trim(),
      beneficiaryCount: beneficiaryCount.trim(),
      objectives: objectives.map((o) => o.trim()).filter(Boolean),
      steps: steps.map((s) => s.trim()).filter(Boolean),
      photos: extraPhotos,
    } : undefined;
    if (editingId) {
      onEdit(editingId, { title: title.trim(), notes: notes.trim(), photo, program });
    } else {
      onAdd({ id: uid(), title: title.trim(), notes: notes.trim(), photo, date: todayKey(), program });
    }
    resetForm();
  };

  return (
    <Modal title={category.title} onClose={onClose} accent="magic" wide>
      <input ref={quickReportInputRef} type="file" onChange={handleQuickReportUpload} style={{ display: "none" }} />
      <button
        onClick={() => quickReportInputRef.current?.click()}
        className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white mb-3 transition-all hover:brightness-110"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
      >
        <Paperclip size={16} /> رفع تقرير أو ملف جاهز كشاهد جديد مباشرة
      </button>
      <p className="text-xs text-center mb-4" style={{ color: MUTED }}>— أو أضف شاهدًا يدويًا بالتفصيل بالأسفل —</p>
      <div className="rounded-xl p-3 mb-4" style={{ background: `${category.color}10`, border: `1px solid ${category.color}40` }}>
        <div className="flex items-center justify-between mb-2">
          <p className="text-xs font-bold" style={{ color: category.color }}>{editingId ? "تعديل الشاهد" : "إضافة شاهد جديد"}</p>
          {editingId && (
            <button onClick={resetForm} className="text-xs font-semibold" style={{ color: MUTED }}>إلغاء التعديل</button>
          )}
        </div>
        <div className="flex flex-wrap gap-1.5 mb-2">
          {category.suggestions.map((s) => (
            <button key={s} onClick={() => setTitle(s)} className="text-xs px-2.5 py-1 rounded-full hover:opacity-80" style={{ background: "#fff", border: `1px solid ${LINE}`, color: INK }}>
              {s}
            </button>
          ))}
        </div>
        <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="عنوان الشاهد" style={{ ...inputStyle, marginBottom: 8 }} />
        <textarea value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="ملاحظات إضافية (اختياري)" style={{ ...inputStyle, minHeight: 60, resize: "vertical", marginBottom: 8 }} />
        <div className="flex items-center gap-2 mb-2">
          <input ref={photoInputRef} type="file" accept="image/*" onChange={handlePhotoChange} style={{ display: "none" }} />
          {photo ? (
            <div className="relative">
              <img src={photo} alt="" className="w-14 h-14 rounded-lg object-cover dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />
              <button onClick={() => setPhoto(null)} className="absolute -top-1.5 -left-1.5 p-0.5 rounded-full" style={{ background: "#C0392B" }}><X size={10} color="#fff" /></button>
            </div>
          ) : (
            <button onClick={() => photoInputRef.current?.click()} className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold" style={{ border: `1px solid ${LINE}`, background: "#fff", color: INK }}>
              <Camera size={14} color={category.color} /> إرفاق صورة (اختياري)
            </button>
          )}
        </div>

        <label className="flex items-center gap-2 text-xs font-semibold mb-2 mt-1" style={{ color: INK }}>
          <input type="checkbox" checked={showProgram} onChange={(e) => setShowProgram(e.target.checked)} />
          <ClipboardList size={13} color={category.color} /> إضافة كتقرير برنامج تفصيلي (نموذج إداري رسمي)
        </label>

        {showProgram && (
          <div className="p-3 rounded-xl mb-2 space-y-2" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <div className="grid grid-cols-2 gap-2">
              <input value={programField} onChange={(e) => setProgramField(e.target.value)} placeholder="مجال البرنامج (مثال: النشاط الثقافي)" style={{ ...inputStyle, padding: "8px 10px", fontSize: 12 }} />
              <input value={executionDate} onChange={(e) => setExecutionDate(e.target.value)} placeholder="تاريخ التنفيذ" style={{ ...inputStyle, padding: "8px 10px", fontSize: 12 }} />
              <input value={implementers} onChange={(e) => setImplementers(e.target.value)} placeholder="المنفّذون (أسماء المعلمين)" style={{ ...inputStyle, padding: "8px 10px", fontSize: 12 }} />
              <input value={beneficiaryCount} onChange={(e) => setBeneficiaryCount(e.target.value)} placeholder="عدد المستفيدين" style={{ ...inputStyle, padding: "8px 10px", fontSize: 12 }} />
            </div>
            <input value={beneficiaries} onChange={(e) => setBeneficiaries(e.target.value)} placeholder="المستفيدون" style={{ ...inputStyle, padding: "8px 10px", fontSize: 12 }} />
            <div>
              <p className="text-xs font-bold mb-1" style={{ color: DASH_GREEN }}>الأهداف</p>
              <TextListEditor items={objectives} onChange={setObjectives} placeholder="هدف..." />
            </div>
            <div>
              <p className="text-xs font-bold mb-1" style={{ color: DASH_GREEN }}>خطوات التنفيذ</p>
              <TextListEditor items={steps} onChange={setSteps} placeholder="خطوة..." />
            </div>
            <div>
              <p className="text-xs font-bold mb-1" style={{ color: DASH_GREEN }}>صور الشواهد (حتى ٤ صور)</p>
              <div className="flex flex-wrap gap-2">
                {extraPhotos.map((ph, i) => (
                  <div key={i} className="relative">
                    <img src={ph} alt="" className="w-14 h-14 rounded-lg object-cover" style={{ border: `1px solid ${LINE}` }} />
                    <button onClick={() => setExtraPhotos(extraPhotos.filter((_, ix) => ix !== i))} className="absolute -top-1.5 -left-1.5 p-0.5 rounded-full" style={{ background: "#C0392B" }}><X size={10} color="#fff" /></button>
                  </div>
                ))}
                {extraPhotos.length < 4 && (
                  <>
                    <input ref={extraPhotosInputRef} type="file" accept="image/*" multiple onChange={handleExtraPhotosChange} style={{ display: "none" }} />
                    <button onClick={() => extraPhotosInputRef.current?.click()} className="w-14 h-14 rounded-lg flex items-center justify-center" style={{ border: `1px dashed ${LINE}` }}>
                      <Plus size={16} color={MUTED} />
                    </button>
                  </>
                )}
              </div>
            </div>
          </div>
        )}
        <button
          disabled={!title.trim()}
          onClick={submit}
          className="w-full py-2 rounded-lg text-sm font-bold text-white disabled:opacity-40"
          style={{ background: category.color }}
        >
          {editingId ? "حفظ التعديل" : "إضافة"}
        </button>
      </div>

      <p className="text-xs font-bold mb-2" style={{ color: MUTED }}>الشواهد المضافة ({entries.length})</p>
      {entries.length === 0 ? (
        <p className="text-sm text-center py-8" style={{ color: MUTED }}>لا يوجد شواهد بهذي الفئة بعد.</p>
      ) : (
        <div className="space-y-2">
          {entries.map((e, i) => {
            const isExpanded = expandedEntryId === e.id;
            const attachments = e.attachments || [];
            return (
              <div key={e.id} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${editingId === e.id ? category.color : LINE}`, background: "#fff" }}>
                <div className="flex items-start gap-3 p-3">
                  {e.photo && <img src={e.photo} alt="" className="w-12 h-12 rounded-lg object-cover shrink-0 dark-mode-img-fix" />}
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-semibold" style={{ color: INK }}>{e.title}</p>
                    {e.notes && <p className="text-xs mt-0.5" style={{ color: MUTED }}>{e.notes}</p>}
                    <div className="flex items-center gap-2 mt-1">
                      <p className="text-[11px]" style={{ color: MUTED }}>{formatDateDisplay(e.date)}</p>
                      <button
                        onClick={() => setExpandedEntryId(isExpanded ? null : e.id)}
                        className="flex items-center gap-1 text-[11px] font-semibold px-2 py-0.5 rounded-full"
                        style={{ background: attachments.length ? `${category.color}15` : "#F3F1E9", color: attachments.length ? category.color : MUTED }}
                      >
                        <Paperclip size={10} /> {attachments.length > 0 ? `${attachments.length} ملف مرفق` : "إضافة ملفات"}
                      </button>
                    </div>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button onClick={() => onMove(e.id, -1)} disabled={i === 0} title="نقل لأعلى" className="p-1.5 rounded-lg hover:bg-black/5 disabled:opacity-30"><ChevronUp size={14} color={MUTED} /></button>
                    <button onClick={() => onMove(e.id, 1)} disabled={i === entries.length - 1} title="نقل لأسفل" className="p-1.5 rounded-lg hover:bg-black/5 disabled:opacity-30"><ChevronDown size={14} color={MUTED} /></button>
                    <button onClick={() => shareEntryRaw(e)} title="مشاركة" className="p-1.5 rounded-lg hover:bg-black/5"><Share2 size={14} color={MUTED} /></button>
                    <button onClick={() => onShareReadOnlyEntry(e)} title="مشاركة للقراءة فقط" className="p-1.5 rounded-lg hover:bg-black/5"><Link2 size={14} color={MUTED} /></button>
                    {e.program && (
                      <button onClick={() => onPrintProgram(e)} title="طباعة كتقرير برنامج" className="p-1.5 rounded-lg hover:bg-black/5"><Printer size={14} color={category.color} /></button>
                    )}
                    <button onClick={() => startEdit(e)} title="تعديل" className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={14} color={MUTED} /></button>
                    <button onClick={() => onArchive(e.id)} title="أرشفة" className="p-1.5 rounded-lg hover:bg-black/5"><Archive size={14} color={MUTED} /></button>
                    <button onClick={() => onDelete(e.id)} title="حذف" className="p-1.5 rounded-lg hover:bg-black/5"><Trash2 size={14} color="#C0392B" /></button>
                  </div>
                </div>
                {isExpanded && (
                  <div className="px-3 pb-3 pt-1" style={{ borderTop: `1px solid ${LINE}`, background: "#FAF8F3" }}>
                    <p className="text-xs font-bold mt-2 mb-2" style={{ color: INK }}>الملفات المرفقة — مثل مجلد Google Drive صغير لهذا الشاهد</p>
                    {attachments.length > 0 && (
                      <div className="space-y-1.5 mb-2">
                        {attachments.map((a) => {
                          const Icon = libraryFileIcon(a.mimeType);
                          return (
                            <div key={a.id} className="flex items-center gap-2 p-2 rounded-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
                              <Icon size={16} color={category.color} className="shrink-0" />
                              <span className="flex-1 text-xs truncate" style={{ color: INK }}>{a.name}</span>
                              <a href={a.dataUrl} download={a.name} title="تنزيل/مشاركة" className="p-1 rounded hover:bg-black/5 shrink-0"><ImageDown size={13} color={MUTED} /></a>
                              <button onClick={() => handleRemoveAttachment(e.id, a.id)} title="حذف الملف" className="p-1 rounded hover:bg-black/5 shrink-0"><X size={13} color="#C0392B" /></button>
                            </div>
                          );
                        })}
                      </div>
                    )}
                    <input ref={attachInputRef} type="file" multiple onChange={(ev) => handleAttachFile(e.id, ev)} style={{ display: "none" }} />
                    <button
                      onClick={() => attachInputRef.current?.click()}
                      className="flex items-center gap-1.5 text-xs font-semibold px-3 py-1.5 rounded-lg"
                      style={{ border: `1px solid ${LINE}`, background: "#fff", color: INK }}
                    >
                      <Plus size={13} /> إضافة ملف (أي نوع)
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </Modal>
  );
}

function UnifiedArchiveModal({ archivedClasses, shawahed, archivedTests, onRestoreClass, onRestoreShawahed, onRestoreTest, onDeleteTestForever, onClose, bare = false }) {
  const [tab, setTab] = useState("classes");
  const archivedEntries = shawahed.archivedEntries || {};
  const shawahedCount = getAllShawahedCategories(shawahed).reduce((sum, c) => sum + (archivedEntries[c.key]?.length || 0), 0);
  const tabs = [
    { key: "classes", label: `الفصول (${archivedClasses.length})` },
    { key: "shawahed", label: `الشواهد (${shawahedCount})` },
    { key: "tests", label: `الاختبارات (${archivedTests.length})` },
  ];

  const content = (
    <>
      <div className="flex items-center gap-2 mb-4">
        {tabs.map((t) => (
          <button
            key={t.key}
            onClick={() => setTab(t.key)}
            className="px-3 py-1.5 rounded-full text-xs font-semibold"
            style={{ background: tab === t.key ? INK : "transparent", color: tab === t.key ? "#fff" : MUTED, border: `1px solid ${tab === t.key ? INK : LINE}` }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === "classes" && (
        archivedClasses.length === 0 ? (
          <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد فصول مؤرشفة.</p>
        ) : (
          <div className="space-y-2">
            {archivedClasses.map((c) => (
              <div key={c.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: c.color }} />
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold truncate" style={{ color: INK }}>{c.subject}</p>
                  <p className="text-xs" style={{ color: MUTED }}>{c.grade} • {c.teacher}</p>
                </div>
                <button onClick={() => onRestoreClass(c.id)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold shrink-0" style={{ border: `1px solid ${LINE}`, color: "#26423B" }}>
                  <RotateCcw size={12} /> استعادة
                </button>
              </div>
            ))}
          </div>
        )
      )}

      {tab === "shawahed" && (
        (() => {
          const allArchived = getAllShawahedCategories(shawahed).flatMap((cat) => (archivedEntries[cat.key] || []).map((e) => ({ ...e, cat })));
          if (allArchived.length === 0) return <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد شواهد مؤرشفة.</p>;
          return (
            <div className="space-y-2">
              {allArchived.map((e) => (
                <div key={e.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                  {e.photo && <img src={e.photo} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0 dark-mode-img-fix" />}
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-semibold" style={{ color: INK }}>{e.title}</p>
                    <p className="text-xs" style={{ color: MUTED }}>{e.cat.title}</p>
                  </div>
                  <button onClick={() => onRestoreShawahed(e.cat.key, e.id)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold shrink-0" style={{ border: `1px solid ${LINE}`, color: "#26423B" }}>
                    <RotateCcw size={12} /> استعادة
                  </button>
                </div>
              ))}
            </div>
          );
        })()
      )}

      {tab === "tests" && (
        archivedTests.length === 0 ? (
          <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد اختبارات مؤرشفة.</p>
        ) : (
          <div className="space-y-2">
            {archivedTests.map((t) => (
              <div key={t.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold truncate" style={{ color: INK }}>{t.title}</p>
                  <p className="text-xs" style={{ color: MUTED }}>{t.questions.length} سؤال</p>
                </div>
                <button onClick={() => onRestoreTest(t.id)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold shrink-0" style={{ border: `1px solid ${LINE}`, color: "#26423B" }}>
                  <RotateCcw size={12} /> استعادة
                </button>
                <button onClick={() => onDeleteTestForever(t.id)} title="حذف نهائي" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0"><Trash2 size={14} color="#C0392B" /></button>
              </div>
            ))}
          </div>
        )
      )}
    </>
  );

  if (bare) return content;

  return (
    <Modal title="الأرشيف" onClose={onClose} wide>
      {content}
    </Modal>
  );
}

function ShawahedArchiveModal({ shawahed, archivedEntries, onRestore, onClose }) {
  const allArchived = getAllShawahedCategories(shawahed || {}).flatMap((cat) => (archivedEntries[cat.key] || []).map((e) => ({ ...e, cat })));
  return (
    <Modal title="أرشيف الشواهد" onClose={onClose} wide>
      {allArchived.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد شواهد مؤرشفة.</p>
      ) : (
        <div className="space-y-2">
          {allArchived.map((e) => (
            <div key={e.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
              {e.photo && <img src={e.photo} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0 dark-mode-img-fix" />}
              <div className="flex-1 min-w-0">
                <p className="text-sm font-semibold" style={{ color: INK }}>{e.title}</p>
                <p className="text-xs" style={{ color: MUTED }}>{e.cat.title}</p>
              </div>
              <button onClick={() => onRestore(e.cat.key, e.id)} className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold shrink-0" style={{ border: `1px solid ${LINE}`, color: "#26423B" }}>
                <RotateCcw size={12} /> استعادة
              </button>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

// نافذة تحديد "هدف" (عدد شواهد مطلوب) لكل معيار من المعايير الاثني عشر —
// يُستخدم بعدها لعرض شريط تقدّم بصري بدل الاعتماد على تذكّر أي معيار ناقص.
function ShawahedGoalsModal({ shawahed, goals, onSave, onClose }) {
  const [draft, setDraft] = useState(() => ({ ...(goals || {}) }));
  const setGoal = (key, val) => setDraft((d) => ({ ...d, [key]: val }));
  const allCats = getAllShawahedCategories(shawahed || {});

  return (
    <Modal title="تحديد أهداف الشواهد" onClose={onClose} accent="magic" wide>
      <p className="text-xs mb-4" style={{ color: MUTED }}>
        حدّد عدد الشواهد المستهدف لكل معيار (اتركه فارغًا لعدم عرض شريط تقدّم لذلك المعيار). بعد الحفظ، تظهر أشرطة تقدّم بصرية بكل بطاقة تبيّن مدى اكتمالها.
      </p>
      <div className="space-y-2 mb-4 max-h-96 overflow-y-auto">
        {allCats.map((cat, i) => (
          <div key={cat.key} className="flex items-center gap-3 p-2.5 rounded-xl" style={{ border: `1px solid ${LINE}` }}>
            <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: cat.color }} />
            <span className="flex-1 text-sm" style={{ color: INK }}>{i + 1}. {cat.title}</span>
            <input
              type="number"
              min={0}
              value={draft[cat.key] ?? ""}
              onChange={(e) => setGoal(cat.key, e.target.value ? Math.max(0, Number(e.target.value)) : "")}
              style={{ ...inputStyle, width: 80, textAlign: "center" }}
              placeholder="—"
            />
          </div>
        ))}
      </div>
      <button
        onClick={() => onSave(draft)}
        className="w-full py-2.5 rounded-xl text-sm font-bold text-white transition-all hover:brightness-110"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
      >
        حفظ الأهداف
      </button>
    </Modal>
  );
}

// مودال يسمح باختيار فئة واحدة (تقرير مستقل) أو عدة فئات (تُدمج بتقرير
// واحد) قبل المعاينة/الطباعة/المشاركة، مع وصف عام جاهز وقابل للتعديل.
function ShawahedExportPickerModal({ shawahed, initialKeys, teacherName, onClose, onConfirm, onShareReadOnly }) {
  const entries = shawahed.entries || {};
  const available = getAllShawahedCategories(shawahed).filter((c) => (entries[c.key] || []).length > 0);
  const [selected, setSelected] = useState(initialKeys && initialKeys.length ? initialKeys : available.map((c) => c.key));
  const [description, setDescription] = useState("");
  const [descTouched, setDescTouched] = useState(false);

  const toggle = (key) => setSelected((s) => (s.includes(key) ? s.filter((k) => k !== key) : [...s, key]));
  const selectAll = () => setSelected(available.map((c) => c.key));
  const selectNone = () => setSelected([]);

  useEffect(() => {
    if (descTouched) return;
    const chosen = available.filter((c) => selected.includes(c.key));
    const total = chosen.reduce((s, c) => s + (entries[c.key]?.length || 0), 0);
    const namesList = chosen.map((c) => c.title).join("، ");
    const auto = chosen.length === 0
      ? ""
      : chosen.length === 1
      ? `يوثّق هذا التقرير ${total} شاهدًا على أداء ${teacherName || "المعلم/ـة"} فيما يخص معيار "${chosen[0].title}"، ويُظهر التزامه/ـا التطبيقي بهذا الجانب من الأداء الوظيفي.`
      : `يوثّق هذا التقرير ${total} شاهدًا على الأداء الوظيفي لـ ${teacherName || "المعلم/ـة"}، موزّعة على المعايير التالية: ${namesList}.`;
    setDescription(auto);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected.join(","), available.length]);

  return (
    <Modal title="اختر الفئات للتقرير" onClose={onClose} accent="magic">
      {available.length === 0 ? (
        <p className="text-sm text-center py-8" style={{ color: MUTED }}>لا يوجد شواهد مضافة بعد.</p>
      ) : (
        <>
          <div className="flex items-center gap-2 mb-3">
            <button onClick={selectAll} className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ border: `1px solid ${LINE}`, color: INK }}>تحديد الكل</button>
            <button onClick={selectNone} className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ border: `1px solid ${LINE}`, color: INK }}>إلغاء التحديد</button>
            <span className="text-xs mr-auto" style={{ color: MUTED }}>
              {selected.length === 1 ? "تقرير فئة واحدة" : selected.length > 1 ? `دمج ${selected.length} فئة بتقرير واحد` : "لم يُحدَّد شيء"}
            </span>
          </div>
          <div className="space-y-1.5 mb-4 max-h-56 overflow-y-auto">
            {available.map((c) => (
              <label
                key={c.key}
                className="flex items-center gap-2.5 p-2.5 rounded-xl cursor-pointer"
                style={{ border: `1px solid ${selected.includes(c.key) ? c.color : LINE}`, background: selected.includes(c.key) ? `${c.color}0F` : "#fff" }}
              >
                <input type="checkbox" checked={selected.includes(c.key)} onChange={() => toggle(c.key)} />
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: c.color }} />
                <span className="flex-1 text-sm font-medium" style={{ color: INK }}>{c.title}</span>
                <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: "#F3F1E9", color: MUTED }}>{entries[c.key].length}</span>
              </label>
            ))}
          </div>
          <Field label="وصف عام للتقرير (جاهز — تقدر تعدّله)">
            <textarea
              value={description}
              onChange={(e) => { setDescription(e.target.value); setDescTouched(true); }}
              style={{ ...inputStyle, minHeight: 80, resize: "vertical" }}
              placeholder="وصف مختصر يوضّح فحوى الشواهد المرفقة..."
            />
          </Field>
          <button
            disabled={selected.length === 0}
            onClick={() => onConfirm(selected, description)}
            className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
            style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
          >
            معاينة التقرير {selected.length > 0 ? `(${selected.length})` : ""}
          </button>
          {onShareReadOnly && (
            <button
              disabled={selected.length === 0}
              onClick={() => onShareReadOnly(selected, description)}
              className="w-full mt-2 py-2.5 rounded-xl text-sm font-bold disabled:opacity-40 transition-all hover:bg-black/5"
              style={{ border: `1px solid ${LINE}`, color: INK }}
            >
              مشاركة رابط للقراءة فقط (للإدارة مثلًا)
            </button>
          )}
        </>
      )}
    </Modal>
  );
}

// مكتبة الملفات: مكان لرفع وحفظ مصادر تعليمية (أوراق عمل، عروض...)
// ومشاركتها مع فصل معيّن أو كل الفصول — بدل التنقل بين تطبيقات ثانية.
// ⚠️ الملفات تُخزَّن كـ base64 داخل بيانات حسابك (نفس أسلوب صور الطلاب
// والشعارات بالتطبيق) — مناسب لملفات صغيرة إلى متوسطة، وننصح بعدم رفع
// ملفات كبيرة جدًا (فيديوهات مثلًا) حتى لا يبطئ تحميل بياناتك.
function libraryFileIcon(mimeType) {
  if (!mimeType) return FileText;
  if (mimeType.startsWith("image/")) return FileImage;
  if (mimeType.includes("pdf")) return FileText;
  if (mimeType.includes("sheet") || mimeType.includes("excel")) return FileSpreadsheet;
  return FileText;
}

// شارة صغيرة تبين هل التطبيق متصل بالإنترنت فعليًا وهل آخر تعديل انحفظ
// بالسحابة أو لسه بانتظار الاتصال — تطمينة سريعة خصوصًا بشبكة مدرسة ضعيفة.
function SyncStatusBadge({ isOnline, syncStatus }) {
  let icon = Check, color = "#0F9D58", bg = "#E3F1EC", label = "محفوظ";
  if (!isOnline) { icon = WifiOff; color = "#C0392B"; bg = "#FBEAE7"; label = "غير متصل"; }
  else if (syncStatus === "saving") { icon = RefreshCw; color = "#C97A2B"; bg = "#FCEFE2"; label = "جارٍ الحفظ..."; }
  else if (syncStatus === "error") { icon = AlertTriangle; color = "#C0392B"; bg = "#FBEAE7"; label = "تعذّر الحفظ"; }
  const Icon = icon;
  return (
    <div
      title={isOnline ? (syncStatus === "saving" ? "جارٍ حفظ آخر تعديل بالسحابة..." : syncStatus === "error" ? "تعذّر حفظ آخر تعديل — تحقق من الاتصال" : "كل تعديلاتك محفوظة بالسحابة") : "لا يوجد اتصال بالإنترنت — تعديلاتك محفوظة بجهازك بس حاليًا وبتُرفع تلقائيًا لما يرجع الاتصال"}
      className="flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold shrink-0"
      style={{ background: bg, color }}
    >
      <Icon size={12} className={syncStatus === "saving" && isOnline ? "animate-spin" : ""} />
      {label}
    </div>
  );
}

function LibraryHub({ library, classes, onUpload, onDelete, onAssign, bare = false }) {
  const fileInputRef = useRef(null);
  const [filterClassId, setFilterClassId] = useState("all");
  const [uploading, setUploading] = useState(false);

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 25 * 1024 * 1024) {
      alert("الملف كبير جدًا (أكبر من ٢٥ ميجا) — يُفضّل رفع ملفات أصغر حتى لا يبطئ تحميل بياناتك.");
      e.target.value = "";
      return;
    }
    setUploading(true);
    const reader = new FileReader();
    reader.onload = () => {
      onUpload({
        id: uid(),
        name: file.name,
        mimeType: file.type,
        size: file.size,
        dataUrl: reader.result,
        classId: null,
        uploadedAt: todayKey(),
      });
      setUploading(false);
    };
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const filtered = (library || []).filter((item) => filterClassId === "all" || item.classId === filterClassId || (filterClassId === "shared" && !item.classId));

  const content = (
    <>
      <div className="p-3 rounded-xl mb-4 flex items-start gap-2" style={{ background: "#FCEFE2", border: "1px solid #F0D2CB" }}>
        <Info size={15} color="#C97A2B" className="shrink-0 mt-0.5" />
        <p className="text-xs" style={{ color: "#8A4A1E" }}>
          الملفات تُحفظ ضمن بيانات حسابك مباشرة — تجنّب رفع ملفات كبيرة جدًا (الحد الأقصى هنا ٢٥ ميجابايت لكل ملف).
        </p>
      </div>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <input ref={fileInputRef} type="file" onChange={handleFile} style={{ display: "none" }} />
        <button
          onClick={() => fileInputRef.current?.click()}
          disabled={uploading}
          className="flex items-center gap-2 px-4 py-2 rounded-xl text-sm font-bold text-white disabled:opacity-60"
          style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
        >
          <Plus size={16} /> {uploading ? "جارٍ الرفع..." : "رفع ملف جديد"}
        </button>
        <select value={filterClassId} onChange={(e) => setFilterClassId(e.target.value)} style={{ ...inputStyle, width: "auto" }}>
          <option value="all">كل الملفات</option>
          <option value="shared">مشتركة (كل الفصول)</option>
          {classes.filter((c) => !c.archived).map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade}</option>)}
        </select>
      </div>

      {filtered.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد ملفات بعد.</p>
      ) : (
        <div className="space-y-2">
          {filtered.map((item) => {
            const Icon = libraryFileIcon(item.mimeType);
            const cls = classes.find((c) => c.id === item.classId);
            return (
              <div key={item.id} className="flex items-center gap-3 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ background: "#F3F1E9" }}>
                  <Icon size={18} color="#26423B" />
                </div>
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-semibold truncate" style={{ color: INK }}>{item.name}</p>
                  <p className="text-xs" style={{ color: MUTED }}>{formatDateDisplay(item.uploadedAt)} • {cls ? `${cls.subject} — ${cls.grade}` : "مشترك (كل الفصول)"}</p>
                </div>
                <select
                  value={item.classId || ""}
                  onChange={(e) => onAssign(item.id, e.target.value || null)}
                  className="text-xs rounded-lg px-2 py-1.5 shrink-0"
                  style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff", maxWidth: 130 }}
                >
                  <option value="">مشترك (الكل)</option>
                  {classes.filter((c) => !c.archived).map((c) => <option key={c.id} value={c.id}>{c.subject}</option>)}
                </select>
                <a href={item.dataUrl} download={item.name} title="تنزيل" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0">
                  <ImageDown size={15} color={MUTED} />
                </a>
                <button onClick={() => onDelete(item.id)} title="حذف" className="p-1.5 rounded-lg hover:bg-black/5 shrink-0">
                  <Trash2 size={15} color="#C0392B" />
                </button>
              </div>
            );
          })}
        </div>
      )}
    </>
  );

  if (bare) return content;
  return (
    <Modal title="المكتبة" onClose={() => {}} wide>
      {content}
    </Modal>
  );
}

// نافذة إضافة/تعديل فئة شواهد خاصة بالمعلم (فوق المعايير الرسمية الـ١٢).
function ShawahedCategoryEditModal({ initial, onClose, onSave }) {
  const [title, setTitle] = useState(initial?.title || "");
  const [color, setColor] = useState(initial?.color || COLORS[0].hex);
  return (
    <Modal title={initial ? "تعديل الفئة" : "إضافة فئة جديدة"} onClose={onClose} accent="magic">
      <Field label="اسم الفئة">
        <input value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle} placeholder="مثال: التميز والابتكار" />
      </Field>
      <Field label="لون الفئة">
        <ColorSwatches value={color} onChange={setColor} />
      </Field>
      <button
        disabled={!title.trim()}
        onClick={() => onSave(title.trim(), color)}
        className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
      >
        {initial ? "حفظ التعديل" : "إضافة الفئة"}
      </button>
    </Modal>
  );
}

function ShawahedHub({ shawahed, onUpdate, onClose, onExport, onQuickPrint, onShareReadOnly, onShareReadOnlyOne, onPrintProgram, teacherName, bare = false }) {
  const [openCat, setOpenCat] = useState(null);
  const [showArchive, setShowArchive] = useState(false);
  const [showExportPicker, setShowExportPicker] = useState(false);
  const [showGoals, setShowGoals] = useState(false);
  const [showAddCategory, setShowAddCategory] = useState(false);
  const [editingCategory, setEditingCategory] = useState(null);
  const [confirmDeleteCategory, setConfirmDeleteCategory] = useState(null);
  const entries = shawahed.entries || {};
  const goals = shawahed.goals || {};
  const archivedEntries = shawahed.archivedEntries || {};
  const allCategories = getAllShawahedCategories(shawahed);
  const totalCount = allCategories.reduce((sum, c) => sum + (entries[c.key]?.length || 0), 0);

  const addEntry = (catKey, entry) => {
    onUpdate({ ...shawahed, entries: { ...entries, [catKey]: [...(entries[catKey] || []), entry] } });
  };
  const editEntry = (catKey, entryId, patch) => {
    onUpdate({ ...shawahed, entries: { ...entries, [catKey]: (entries[catKey] || []).map((e) => (e.id === entryId ? { ...e, ...patch } : e)) } });
  };
  const deleteEntry = (catKey, entryId) => {
    onUpdate({ ...shawahed, entries: { ...entries, [catKey]: (entries[catKey] || []).filter((e) => e.id !== entryId) } });
  };
  const archiveEntry = (catKey, entryId) => {
    const entry = (entries[catKey] || []).find((e) => e.id === entryId);
    if (!entry) return;
    onUpdate({
      ...shawahed,
      entries: { ...entries, [catKey]: entries[catKey].filter((e) => e.id !== entryId) },
      archivedEntries: { ...archivedEntries, [catKey]: [...(archivedEntries[catKey] || []), entry] },
    });
  };
  const restoreEntry = (catKey, entryId) => {
    const entry = (archivedEntries[catKey] || []).find((e) => e.id === entryId);
    if (!entry) return;
    onUpdate({
      ...shawahed,
      entries: { ...entries, [catKey]: [...(entries[catKey] || []), entry] },
      archivedEntries: { ...archivedEntries, [catKey]: archivedEntries[catKey].filter((e) => e.id !== entryId) },
    });
  };
  const moveEntry = (catKey, entryId, direction) => {
    const list = entries[catKey] || [];
    const idx = list.findIndex((e) => e.id === entryId);
    const swapIdx = idx + direction;
    if (idx < 0 || swapIdx < 0 || swapIdx >= list.length) return;
    const next = [...list];
    [next[idx], next[swapIdx]] = [next[swapIdx], next[idx]];
    onUpdate({ ...shawahed, entries: { ...entries, [catKey]: next } });
  };

  const moveCategory = (key, direction) => {
    const keys = allCategories.map((c) => c.key);
    const idx = keys.indexOf(key);
    const swapIdx = idx + direction;
    if (idx < 0 || swapIdx < 0 || swapIdx >= keys.length) return;
    const next = [...keys];
    [next[idx], next[swapIdx]] = [next[swapIdx], next[idx]];
    onUpdate({ ...shawahed, categoryOrder: next });
  };
  const addCustomCategory = (title, color) => {
    const key = `custom-${uid()}`;
    const newCat = { key, title, color, suggestions: [] };
    const keys = allCategories.map((c) => c.key);
    onUpdate({
      ...shawahed,
      customCategories: [...(shawahed.customCategories || []), newCat],
      categoryOrder: [...keys, key],
    });
    setShowAddCategory(false);
  };
  const editCustomCategory = (key, patch) => {
    onUpdate({ ...shawahed, customCategories: (shawahed.customCategories || []).map((c) => (c.key === key ? { ...c, ...patch } : c)) });
    setEditingCategory(null);
  };
  const deleteCustomCategory = (key) => {
    const { [key]: _removed, ...restEntries } = entries;
    onUpdate({
      ...shawahed,
      customCategories: (shawahed.customCategories || []).filter((c) => c.key !== key),
      entries: restEntries,
      categoryOrder: (shawahed.categoryOrder || []).filter((k) => k !== key),
    });
    setConfirmDeleteCategory(null);
  };

  const content = (
    <>
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <p className="text-xs font-semibold flex-1" style={{ color: MUTED }}>{totalCount} شاهد بكل الفئات</p>
        <IconBtn icon={Plus} label="إضافة فئة" onClick={() => setShowAddCategory(true)} />
        <IconBtn icon={Target} label="تحديد أهداف" onClick={() => setShowGoals(true)} />
        <IconBtn icon={Archive} label="الأرشيف" onClick={() => setShowArchive(true)} />
        <IconBtn icon={FileText} label="طباعة / مشاركة" magic onClick={() => setShowExportPicker(true)} />
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
        {allCategories.map((cat, i) => {
          const count = entries[cat.key]?.length || 0;
          const goal = Number(goals[cat.key]) || 0;
          const pct = goal > 0 ? Math.min(100, Math.round((count / goal) * 100)) : null;
          const isCustom = isCustomShawahedCategory(shawahed, cat.key);
          return (
            <div
              key={cat.key}
              className="text-right rounded-2xl p-4 hover:-translate-y-0.5 transition-all"
              style={{ background: "#fff", border: `1px solid ${LINE}`, boxShadow: "0 1px 3px rgba(35,38,34,0.06)" }}
            >
              <button onClick={() => setOpenCat(cat)} className="w-full text-right">
                <div className="flex items-center justify-between mb-2">
                  <span className="w-9 h-9 rounded-full flex items-center justify-center font-bold text-white" style={{ background: cat.color }}>{count}</span>
                  <FileCheck size={16} color={cat.color} />
                </div>
                <p className="text-sm font-semibold leading-snug mb-2" style={{ color: INK }}>{cat.title}</p>
                {pct !== null && (
                  <div className="mb-2">
                    <div className="w-full h-1.5 rounded-full overflow-hidden mb-1" style={{ background: "#F0EFE9" }}>
                      <div className="h-full rounded-full transition-all" style={{ width: `${pct}%`, background: pct >= 100 ? "#0F9D58" : cat.color }} />
                    </div>
                    <p className="text-[11px]" style={{ color: pct >= 100 ? "#0F9D58" : MUTED }}>
                      {pct >= 100 ? "✓ اكتمل الهدف" : `${count} من ${goal} (${pct}٪)`}
                    </p>
                  </div>
                )}
              </button>
              <div className="flex items-center gap-1 pt-2 mt-1" style={{ borderTop: `1px solid ${LINE}` }}>
                <button onClick={() => moveCategory(cat.key, -1)} disabled={i === 0} title="نقل لأعلى" className="p-1 rounded hover:bg-black/5 disabled:opacity-30"><ChevronUp size={13} color={MUTED} /></button>
                <button onClick={() => moveCategory(cat.key, 1)} disabled={i === allCategories.length - 1} title="نقل لأسفل" className="p-1 rounded hover:bg-black/5 disabled:opacity-30"><ChevronDown size={13} color={MUTED} /></button>
                {count > 0 && (
                  <button onClick={() => onQuickPrint && onQuickPrint(cat.key)} title="مشاركة" className="p-1 rounded hover:bg-black/5"><Share2 size={13} color={MUTED} /></button>
                )}
                <button onClick={() => onShareReadOnly([cat.key], "")} title="مشاركة للقراءة فقط" className="p-1 rounded hover:bg-black/5"><Link2 size={13} color={MUTED} /></button>
                {isCustom && (
                  <>
                    <button onClick={() => setEditingCategory(cat)} title="تعديل" className="p-1 rounded hover:bg-black/5"><Pencil size={13} color={MUTED} /></button>
                    <button onClick={() => setConfirmDeleteCategory(cat)} title="حذف الفئة" className="p-1 rounded hover:bg-black/5"><Trash2 size={13} color="#C0392B" /></button>
                  </>
                )}
              </div>
            </div>
          );
        })}
      </div>

      {showGoals && (
        <ShawahedGoalsModal
          shawahed={shawahed}
          goals={goals}
          onClose={() => setShowGoals(false)}
          onSave={(newGoals) => { onUpdate({ ...shawahed, goals: newGoals }); setShowGoals(false); }}
        />
      )}

      {(showAddCategory || editingCategory) && (
        <ShawahedCategoryEditModal
          initial={editingCategory}
          onClose={() => { setShowAddCategory(false); setEditingCategory(null); }}
          onSave={(title, color) => (editingCategory ? editCustomCategory(editingCategory.key, { title, color }) : addCustomCategory(title, color))}
        />
      )}

      {confirmDeleteCategory && (
        <ConfirmDialog
          title="حذف الفئة"
          message={`سيتم حذف فئة "${confirmDeleteCategory.title}" وكل الشواهد بداخلها (${(entries[confirmDeleteCategory.key] || []).length}) نهائيًا. متابعة؟`}
          confirmLabel="حذف"
          onCancel={() => setConfirmDeleteCategory(null)}
          onConfirm={() => deleteCustomCategory(confirmDeleteCategory.key)}
        />
      )}

      {showExportPicker && (
        <ShawahedExportPickerModal
          shawahed={shawahed}
          teacherName={teacherName}
          onClose={() => setShowExportPicker(false)}
          onConfirm={(keys, description) => { setShowExportPicker(false); onExport(keys, description); }}
          onShareReadOnly={(keys, description) => { setShowExportPicker(false); onShareReadOnly(keys, description); }}
        />
      )}

      {openCat && (
        <ShawahedCategoryModal
          category={openCat}
          entries={entries[openCat.key] || []}
          onAdd={(entry) => addEntry(openCat.key, entry)}
          onEdit={(entryId, patch) => editEntry(openCat.key, entryId, patch)}
          onDelete={(id) => deleteEntry(openCat.key, id)}
          onArchive={(id) => archiveEntry(openCat.key, id)}
          onMove={(entryId, direction) => moveEntry(openCat.key, entryId, direction)}
          onShareReadOnlyEntry={(entry) => onShareReadOnlyOne(entry, openCat)}
          onPrintProgram={(entry) => onPrintProgram(entry, openCat)}
          onClose={() => setOpenCat(null)}
        />
      )}
      {showArchive && (
        <ShawahedArchiveModal
          shawahed={shawahed}
          archivedEntries={archivedEntries}
          onRestore={restoreEntry}
          onClose={() => setShowArchive(false)}
        />
      )}
    </>
  );

  if (bare) return content;

  return (
    <Modal title="شواهد الأداء الوظيفي" onClose={onClose} accent="magic" xl>
      {content}
    </Modal>
  );
}

function NoorEmbedModal({ onImportNames, onClose, isOwner }) {
  const [mode, setMode] = useState("paste"); // "paste" | "auto"
  const [pasteText, setPasteText] = useState("");
  const [selectedIdx, setSelectedIdx] = useState(null); // Set of selected indices, null = not yet computed

  // إعدادات "الوضع الآلي" (خادم وسيط منشور من طرفك) — تُحفظ محليًا بجهازك
  // فقط (localStorage)، ما تُرفَع لقاعدة البيانات، لأن الكوكي بيانات حساسة.
  const [serverUrl, setServerUrl] = useState(() => localStorage.getItem("fosooli-noor-server-url") || "");
  const [apiKey, setApiKey] = useState(() => localStorage.getItem("fosooli-noor-api-key") || "");
  const [cookie, setCookie] = useState(() => localStorage.getItem("fosooli-noor-cookie") || "");
  const [pagePath, setPagePath] = useState(() => localStorage.getItem("fosooli-noor-page-path") || "");
  const [stage, setStage] = useState("");
  const [classCode, setClassCode] = useState("");
  const [sectionCode, setSectionCode] = useState("");
  const [semester, setSemester] = useState("");
  const [autoResults, setAutoResults] = useState([]);
  const [autoLoading, setAutoLoading] = useState(false);
  const [autoError, setAutoError] = useState("");

  useEffect(() => { localStorage.setItem("fosooli-noor-server-url", serverUrl); }, [serverUrl]);
  useEffect(() => { localStorage.setItem("fosooli-noor-api-key", apiKey); }, [apiKey]);
  useEffect(() => { localStorage.setItem("fosooli-noor-cookie", cookie); }, [cookie]);
  useEffect(() => { localStorage.setItem("fosooli-noor-page-path", pagePath); }, [pagePath]);

  const pastedResults = parseNoorStudentTable(pasteText);
  const parsed = mode === "auto" ? autoResults : pastedResults;
  const selected = selectedIdx || new Set(parsed.map((_, i) => i));
  useEffect(() => {
    // كل مرة يتغيّر الاستخراج (نص جديد لُصق أو نتيجة آلية جديدة) نُعيد تحديد الجميع افتراضيًا
    setSelectedIdx(new Set(parsed.map((_, i) => i)));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pasteText, autoResults]);

  const toggleOne = (i) => {
    setSelectedIdx((prev) => {
      const s = new Set(prev || parsed.map((_, idx) => idx));
      if (s.has(i)) s.delete(i); else s.add(i);
      return s;
    });
  };

  const confirmedCount = parsed.filter((_, i) => selected.has(i)).length;

  const fetchAutomatically = async () => {
    setAutoError("");
    setAutoResults([]);
    if (!serverUrl.trim() || !apiKey.trim() || !cookie.trim() || !pagePath.trim()) {
      setAutoError("عبّي رابط الخادم ومفتاح API والكوكي ورابط الصفحة أولًا.");
      return;
    }
    setAutoLoading(true);
    try {
      const res = await fetch(`${serverUrl.replace(/\/$/, "")}/api/noor/students`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey },
        body: JSON.stringify({ cookie, pagePath, stage, classCode, sectionCode, semester }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error([data.error, data.detail].filter(Boolean).join(" — ") || "فشل الجلب");
      setAutoResults(data.students || []);
    } catch (err) {
      setAutoError(err.message === "Failed to fetch" ? "تعذّر الوصول للخادم — تأكد من الرابط ومن أن الخادم يعمل." : err.message);
    } finally {
      setAutoLoading(false);
    }
  };

  const STEPS = [
    "افتح نور بنافذة مستقلة من الزر تحت، وسجّل دخولك بحسابك (عبر نفاذ أو بوابة الدخول الموحد).",
    "من القائمة الجانبية اختر «الطلاب» ← «الطلاب حسب العلاقات التدريسية».",
    "اختر نظام الدراسة / الصف / القسم / الفصل، ثم اضغط «عرض».",
    "اضغط كلمة «الطلاب» أمام المادة المطلوبة لعرض قائمة الطلاب الكاملة.",
    "حدّد جدول الطلاب بالكامل (اسحب فوقه بالماوس، أو Ctrl+A) وانسخه (Ctrl+C).",
    "ارجع لهذه النافذة، والصق الجدول بالمربع تحت — سيُستخرَج الاسم ورقم الهوية تلقائيًا.",
  ];

  return (
    <Modal title="استيراد الأسماء من نور" onClose={onClose} accent="magic" xl>
      {isOwner && (
        <div className="flex gap-2 mb-4">
          <button onClick={() => setMode("paste")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: mode === "paste" ? INK : "transparent", color: mode === "paste" ? "#fff" : MUTED, border: `1px solid ${mode === "paste" ? INK : LINE}` }}>
            لصق يدوي (يعمل فورًا)
          </button>
          <button onClick={() => setMode("auto")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: mode === "auto" ? INK : "transparent", color: mode === "auto" ? "#fff" : MUTED, border: `1px solid ${mode === "auto" ? INK : LINE}` }}>
            الوضع الآلي (يحتاج خادم متصل)
          </button>
        </div>
      )}

      {(mode === "paste" || !isOwner) ? (
        <>
          <div className="p-3 rounded-xl mb-3 flex items-start gap-2" style={{ background: "#FCEFE2", border: "1px solid #F0D2CB" }}>
            <Info size={15} color="#C97A2B" className="shrink-0 mt-0.5" />
            <p className="text-xs" style={{ color: "#8A4A1E" }}>
              نور يمنع صراحة فتحه داخل تطبيقات أخرى (إجراء أمني من طرفهم)، لذلك الخطوة تُنجَز بفتح نور بنافذة مستقلة ثم نسخ الجدول ولصقه هنا — الاستخراج والإضافة بعدها تلقائي بالكامل.
            </p>
          </div>

          <button
            onClick={() => window.open("https://noor.moe.gov.sa/", "_blank")}
            className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white mb-4 transition-all hover:brightness-110 active:scale-95"
            style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
          >
            <ExternalLink size={16} /> فتح نظام نور
          </button>

          <div className="rounded-xl p-3 mb-4" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
            <p className="text-xs font-bold mb-2" style={{ color: INK }}>خطوات الاستيراد</p>
            <ol className="space-y-1.5">
              {STEPS.map((s, i) => (
                <li key={i} className="flex items-start gap-2 text-xs" style={{ color: MUTED }}>
                  <span className="w-4 h-4 rounded-full flex items-center justify-center shrink-0 font-bold text-white" style={{ background: DASH_GREEN, fontSize: 9 }}>{i + 1}</span>
                  {s}
                </li>
              ))}
            </ol>
          </div>

          <div className="p-3 rounded-xl mb-4" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-xs font-bold mb-2" style={{ color: INK }}>الصق جدول الطلاب من نور هنا</p>
            <textarea
              style={{ ...inputStyle, minHeight: "110px", resize: "vertical", fontFamily: "monospace", fontSize: 12 }}
              value={pasteText}
              onChange={(e) => setPasteText(e.target.value)}
              placeholder={"الصق الجدول كاملًا هنا بعد نسخه من نور..."}
            />
          </div>
        </>
      ) : (
        <>
          <div className="p-3 rounded-xl mb-3 flex items-start gap-2" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
            <Info size={15} color={DASH_GREEN} className="shrink-0 mt-0.5" />
            <p className="text-xs" style={{ color: DASH_GREEN }}>
              هذا الوضع يحتاج خادمًا وسيطًا منشورًا (راجع ملف <b>noor-proxy-server/README.md</b> اللي أرسلته لك لخطوات الإعداد الكاملة — مرة وحدة فقط). بعد التجهيز، عبّي الحقول تحت وستُحفَظ بجهازك.
            </p>
          </div>
          <Field label="رابط الخادم"><input value={serverUrl} onChange={(e) => setServerUrl(e.target.value)} style={inputStyle} placeholder="https://noor-proxy-xxxx.onrender.com" /></Field>
          <Field label="مفتاح API"><input value={apiKey} onChange={(e) => setApiKey(e.target.value)} style={inputStyle} placeholder="نفس SERVER_API_KEY" /></Field>
          <Field label="الكوكي (من تبويب Network بنور)"><textarea value={cookie} onChange={(e) => setCookie(e.target.value)} style={{ ...inputStyle, minHeight: 70, fontFamily: "monospace", fontSize: 11 }} placeholder="الصق قيمة هيدر Cookie كاملة" /></Field>
          <Field label="رابط الصفحة كامل (من شريط العنوان)" hint="انسخه من شريط عنوان المتصفح وأنت واقف بصفحة قائمة الطلاب — يتغيّر كل جلسة دخول، انسخه بنفس وقت الكوكي.">
            <input value={pagePath} onChange={(e) => setPagePath(e.target.value)} style={{ ...inputStyle, fontFamily: "monospace", fontSize: 11 }} placeholder="https://noor.moe.gov.sa/Noor/UsersManagement/TeacherStudents.aspx?EKME-..." />
          </Field>
          <div className="grid grid-cols-2 gap-2 mb-2">
            <input value={stage} onChange={(e) => setStage(e.target.value)} style={inputStyle} placeholder="الصف (اختياري)" />
            <input value={classCode} onChange={(e) => setClassCode(e.target.value)} style={inputStyle} placeholder="القسم (اختياري)" />
            <input value={sectionCode} onChange={(e) => setSectionCode(e.target.value)} style={inputStyle} placeholder="الفصل الدراسي (اختياري)" />
            <input value={semester} onChange={(e) => setSemester(e.target.value)} style={inputStyle} placeholder="فصل آخر (اختياري)" />
          </div>
          <button
            onClick={fetchAutomatically}
            disabled={autoLoading}
            className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white mb-2 disabled:opacity-60 transition-all hover:brightness-110"
            style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
          >
            {autoLoading ? <RefreshCw size={16} className="animate-spin" /> : <Download size={16} />}
            {autoLoading ? "جاري الجلب..." : "جلب الطلاب تلقائيًا"}
          </button>
          {autoError && <p className="text-xs text-center mb-2" style={{ color: "#C0392B" }}>{autoError}</p>}
        </>
      )}

      {parsed.length > 0 && (
        <div className="rounded-2xl overflow-hidden mb-4" style={{ border: `1px solid ${LINE}` }}>
          <div className="p-4" style={{ background: `linear-gradient(135deg, ${DASH_GREEN}, ${DASH_GREEN_DARK})`, color: "#fff" }}>
            <div className="flex items-center gap-3">
              <div className="w-11 h-11 rounded-full flex items-center justify-center shrink-0" style={{ background: "rgba(255,255,255,0.18)" }}>
                <Users size={20} color="#fff" />
              </div>
              <div>
                <p className="font-bold text-base">بيانات الطلاب</p>
                <p className="text-xs opacity-85">تم استخراج {parsed.length} طالب</p>
              </div>
            </div>
          </div>
          <div className="max-h-64 overflow-y-auto">
            {parsed.map((r, i) => (
              <label
                key={i}
                className="flex items-center gap-3 px-4 py-2.5"
                style={{ borderTop: i > 0 ? `1px solid ${LINE}` : "none", background: selected.has(i) ? "#fff" : "#FAF8F3" }}
              >
                <input type="checkbox" checked={selected.has(i)} onChange={() => toggleOne(i)} />
                <span className="flex-1 min-w-0">
                  <span className="block text-sm font-semibold truncate" style={{ color: INK }}>{r.name}</span>
                  {r.nationalId && <span className="block text-xs" style={{ color: MUTED }}>رقم الهوية: {r.nationalId}</span>}
                </span>
                <span className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold shrink-0" style={{ background: GOLD_LIGHT, color: DASH_GREEN }}>{i + 1}</span>
              </label>
            ))}
          </div>
          <div className="flex gap-2 p-3" style={{ borderTop: `1px solid ${LINE}` }}>
            <button
              disabled={confirmedCount === 0}
              onClick={() => onImportNames(parsed.filter((_, i) => selected.has(i)))}
              className="flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
              style={{ background: "#0F9D58" }}
            >
              <Download size={15} /> استيراد البيانات {confirmedCount > 0 ? `(${confirmedCount})` : ""}
            </button>
            <button onClick={() => { setPasteText(""); setAutoResults([]); }} className="px-5 py-2.5 rounded-xl text-sm font-semibold" style={{ border: `1px solid ${LINE}`, color: MUTED }}>
              <X size={14} className="inline" /> إلغاء
            </button>
          </div>
        </div>
      )}
    </Modal>
  );
}

function StudentPickerModal({ rows, onSelect, onClose }) {
  const [selected, setSelected] = useState(rows[0]?.id || "");
  return (
    <Modal title="عرض تقرير طالب" onClose={onClose} accent="magic">
      {rows.length === 0 ? (
        <p className="text-sm text-center py-6" style={{ color: MUTED }}>لا يوجد طلاب بهذا الفصل بعد.</p>
      ) : (
        <>
          <Field label="اختر الطالب">
            <select value={selected} onChange={(e) => setSelected(e.target.value)} style={inputStyle}>
              {rows.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}
            </select>
          </Field>
          <button onClick={() => onSelect(selected)} className="w-full py-2.5 rounded-xl text-sm font-bold text-white transition-all hover:brightness-110 active:scale-95" style={{ background: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)" }}>
            عرض التقرير
          </button>
        </>
      )}
    </Modal>
  );
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
          style={{ background: "#26423B" }}
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
        className="flex items-center gap-1.5 text-sm px-3 py-1.5 rounded-full font-semibold transition-transform active:scale-95"
        style={selected
          ? { background: selected.color, color: "#fff", minWidth: "80px", justifyContent: "center", boxShadow: `0 2px 6px ${selected.color}55` }
          : { border: `1px solid ${LINE}`, background: "#fff", color: MUTED, minWidth: "80px", justifyContent: "center" }}
      >
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

// خانة "مستوى مهارة": مثل القائمة المنسدلة، لكن مع شريط شرائح صغير يوري
// موضع المستوى المختار ضمن الترتيب الكامل (مبتدئ ← متقن) بنظرة واحدة —
// أنسب من رقم مجرّد لمواد المهارات (قراءة، خط...).
function LevelCell({ column, value, onChange }) {
  const [open, setOpen] = useState(false);
  const levels = column.levels || [];
  const idx = levels.findIndex((l) => l.label === value);
  const activeColor = idx >= 0 ? levels[idx].color : LINE;
  return (
    <div className="relative flex flex-col items-center gap-1">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full font-semibold transition-transform active:scale-95"
        style={idx >= 0
          ? { background: activeColor, color: "#fff" }
          : { border: `1px solid ${LINE}`, background: "#fff", color: MUTED }}
      >
        {idx >= 0 ? levels[idx].label : "—"}
      </button>
      {levels.length > 0 && (
        <div className="flex gap-0.5">
          {levels.map((l, i) => (
            <span key={l.id} className="w-4 h-1.5 rounded-full" style={{ background: i <= idx ? activeColor : "#E4DFD2" }} />
          ))}
        </div>
      )}
      {open && (
        <div className="absolute z-30 top-full mt-1 rounded-lg shadow-lg py-1 min-w-[140px]" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
          <button onClick={() => { onChange(""); setOpen(false); }} className="w-full text-right px-3 py-1.5 text-xs hover:bg-black/5" style={{ color: MUTED }}>بدون</button>
          {levels.map((l) => (
            <button key={l.id} onClick={() => { onChange(l.label); setOpen(false); }} className="w-full flex items-center gap-2 px-3 py-1.5 text-sm hover:bg-black/5">
              <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: l.color }} />
              <span style={{ color: INK }}>{l.label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// يبني معرّف DOM ثابت لكل خلية إدخال بالجدول، يُستخدم للتنقل بلوحة
// المفاتيح (Enter) للخلية المقابلة بالصف التالي مباشرة — مفيد لرصد سريع
// لعمود كامل (درجات اختبار مثلًا) بدون ما تلمس كل خلية بالماوس.
function cellDomId(rowId, colId) {
  return `cell-input-${rowId}-${colId}`;
}
function focusNextCell(nextRowId, colId) {
  if (!nextRowId) return;
  requestAnimationFrame(() => {
    const el = document.getElementById(cellDomId(nextRowId, colId));
    if (el) { el.focus(); if (el.select) el.select(); }
  });
}

function Cell({ column, value, onChange, rowId, nextRowId }) {
  if (column.type === "counter") return <CounterCell value={value} onChange={onChange} rowId={rowId} colId={column.id} nextRowId={nextRowId} />;
  if (column.type === "dropdown") return <DropdownCell column={column} value={value} onChange={onChange} />;
  if (column.type === "level") return <LevelCell column={column} value={value} onChange={onChange} />;
  return <TextCell value={value} onChange={onChange} rowId={rowId} colId={column.id} nextRowId={nextRowId} />;
}

// Counter cell: +/- buttons and direct typing both only update a local
// number, and only log ONE report entry once the value settles (debounced)
// or when the input loses focus — so reaching "4" via four clicks (or typing
// "5" directly) logs a single record, not one per click.
function CounterCell({ value, onChange, rowId, colId, nextRowId }) {
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
      <button onClick={() => bump(-1)} className="w-6 h-6 rounded-md flex items-center justify-center hover:bg-black/5 shrink-0" style={{ border: `1px solid ${LINE}` }}><Minus size={12} /></button>
      <input
        id={rowId && colId ? cellDomId(rowId, colId) : undefined}
        type="number"
        value={n}
        onChange={(e) => { const v = Number(e.target.value) || 0; setN(v); scheduleCommit(v); }}
        onBlur={() => commitNow(n)}
        onKeyDown={(e) => {
          if (e.key === "Enter") { commitNow(n); e.currentTarget.blur(); focusNextCell(nextRowId, colId); }
        }}
        className="w-9 text-center font-semibold text-sm bg-transparent"
        style={{ outline: "none", color: INK }}
      />
      <button onClick={() => bump(1)} className="w-6 h-6 rounded-md flex items-center justify-center hover:bg-black/5 shrink-0" style={{ border: `1px solid ${LINE}` }}><Plus size={12} /></button>
    </div>
  );
}

// Free-text cell: typing only updates local state (so it feels normal to type
// in), and only commits — i.e. only logs ONE report entry with the finished
// sentence — on blur or Enter. This is what fixes notes being logged letter
// by letter as you type.
function TextCell({ value, onChange, rowId, colId, nextRowId }) {
  const [local, setLocal] = useState(value || "");
  useEffect(() => { setLocal(value || ""); }, [value]);
  const commit = () => {
    if (local !== (value || "")) onChange(local);
  };
  return (
    <input
      id={rowId && colId ? cellDomId(rowId, colId) : undefined}
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => {
        if (e.key === "Enter") { commit(); e.currentTarget.blur(); focusNextCell(nextRowId, colId); }
      }}
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
  if (entry.type === "clearStudents") return `تفريغ كل الطلاب (${entry.data.rows.length} طالب، الأعمدة بقيت كما هي)`;
  if (entry.type === "bulkRows") return `حذف جماعي (${entry.data.items.length} طالب)`;
  if (entry.type === "class") return `فصل: ${entry.data.cls.subject} — ${entry.data.cls.grade}`;
  if (entry.type === "classesBulk") return `حذف جماعي (${entry.data.classes.length} فصل)`;
  if (entry.type === "reportEntry") return `${entry.data.entry.colName}: ${entry.data.entry.value}`;
  if (entry.type === "reportBulk") return `حذف جماعي من التقرير (${entry.data.entries.length} رصد)`;
  return "عنصر محذوف";
}

// شواهد الأداء الوظيفي — الفئات الاثنتا عشرة الرسمية، وقائمة اقتراحات لكل
// فئة تسهّل إضافة شاهد جديد بضغطة (يقدر المعلم يعدّل النص أو يكتب غيره تمامًا).
const SHAWAHED_CATEGORIES = [
  { key: "duties", title: "أداء الواجبات الوظيفية", color: "#26423B", suggestions: ["سجل الحضور والانصراف", "محضر اجتماع", "خطاب شكر وتقدير"] },
  { key: "community", title: "التفاعل مع المجتمع المهني", color: "#3B4C8C", suggestions: ["شهادة مشاركة بدورة تدريبية", "محضر مجتمع تعلم مهني", "شهادة تقديم ورشة عمل"] },
  { key: "parents", title: "التفاعل مع أولياء الأمور", color: "#C97A2B", suggestions: ["سجل تواصل مع ولي أمر", "دعوة اجتماع أولياء أمور", "استبيان رضا أولياء الأمور"] },
  { key: "strategies", title: "التنوع في استراتيجيات التدريس", color: "#7A4E9E", suggestions: ["تحضير درس باستراتيجية التعلم التعاوني", "صورة من نشاط تعلم نشط", "خطة درس متنوعة الاستراتيجيات"] },
  { key: "outcomes", title: "تحسين نتائج المتعلمين", color: "#2E9FD6", suggestions: ["مقارنة نتائج قبل وبعد", "شهادة تفوق طالب", "تقرير تحسن مستوى الفصل"] },
  { key: "plan", title: "إعداد وتنفيذ خطة التعلم", color: "#B4526A", suggestions: ["خطة فصلية", "خطة درس يومية", "توزيع منهج"] },
  { key: "techUse", title: "توظيف تقنيات ووسائل التعلم المناسبة", color: "#4E6FE0", suggestions: ["لقطة من استخدام السبورة الذكية", "رابط عرض تفاعلي", "صورة وسيلة تعليمية"] },
  { key: "environment", title: "تهيئة بيئة تعليمية", color: "#6B9E4E", suggestions: ["صورة تجهيز الفصل", "لوحة تحفيزية بالفصل", "ركن تعليمي"] },
  { key: "classroom", title: "الإدارة الصفية", color: "#9E6B4E", suggestions: ["نظام لوائح الفصل", "سجل تعزيز سلوك", "خطة إدارة صفية"] },
  { key: "analysis", title: "تحليل نتائج المتعلمين وتشخيص مستواهم", color: "#5C7C9E", suggestions: ["تحليل نتائج اختبار", "تقرير تشخيصي لطالب", "رسم بياني لنتائج الفصل"] },
  { key: "assessment", title: "تنوع أساليب التقويم", color: "#9E4E7A", suggestions: ["نموذج تقويم أداء", "بطاقة ملاحظة", "اختبار قصير متنوع"] },
  { key: "activities", title: "تهيئة البيئة المدرسية للبرامج والأنشطة الطلابية", color: "#C9A227", suggestions: ["صورة نشاط طلابي", "برنامج إذاعة مدرسية", "مسابقة صفية"] },
];

// المعايير الرسمية الاثنا عشر ثابتة، لكن نسمح للمعلم يضيف فئات خاصة به
// فوقها، ويرتّب الكل (رسمي + خاص) بالترتيب اللي يناسبه.
function getAllShawahedCategories(shawahed) {
  const custom = shawahed.customCategories || [];
  const all = [...SHAWAHED_CATEGORIES, ...custom];
  const allKeys = all.map((c) => c.key);
  const savedOrder = (shawahed.categoryOrder || []).filter((k) => allKeys.includes(k));
  const orderedKeys = [...savedOrder, ...allKeys.filter((k) => !savedOrder.includes(k))];
  return orderedKeys.map((k) => all.find((c) => c.key === k)).filter(Boolean);
}

function isCustomShawahedCategory(shawahed, key) {
  return (shawahed.customCategories || []).some((c) => c.key === key);
}



const SUBSCRIPTION_LABELS = { free: "مجاني", trial: "تجريبي", active: "مشترك", expired: "منتهي" };

function TransferStudentsModal({ profiles, onClose }) {
  const [sourceId, setSourceId] = useState("");
  const [destId, setDestId] = useState("");
  const [sourceData, setSourceData] = useState(null);
  const [destData, setDestData] = useState(null);
  const [sourceClassId, setSourceClassId] = useState("");
  const [destClassId, setDestClassId] = useState("");
  const [selectedRowIds, setSelectedRowIds] = useState([]);
  const [includeGrades, setIncludeGrades] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [done, setDone] = useState(false);

  const loadUserData = async (id, setter) => {
    const { data, error } = await supabase.from("user_data").select("data").eq("user_id", id).maybeSingle();
    if (error) { setError(error.message); return; }
    setter(data?.data || { classes: [] });
  };

  useEffect(() => {
    setSourceClassId(""); setSelectedRowIds([]); setSourceData(null);
    if (sourceId) loadUserData(sourceId, setSourceData);
  }, [sourceId]);
  useEffect(() => {
    setDestClassId(""); setDestData(null);
    if (destId) loadUserData(destId, setDestData);
  }, [destId]);

  const sourceClass = sourceData?.classes?.find((c) => c.id === sourceClassId);
  const destClass = destData?.classes?.find((c) => c.id === destClassId);

  const toggleRow = (id) => setSelectedRowIds((ids) => (ids.includes(id) ? ids.filter((x) => x !== id) : [...ids, id]));

  const executeTransfer = async () => {
    if (!sourceClass || !destClass || selectedRowIds.length === 0) return;
    setBusy(true); setError("");
    try {
      const newRows = [];
      const newCellsForDest = {};
      const newReportsForDest = {};
      selectedRowIds.forEach((rowId) => {
        const row = sourceClass.rows.find((r) => r.id === rowId);
        if (!row) return;
        const newId = uid();
        newRows.push({ ...row, id: newId });
        if (includeGrades) {
          sourceClass.columns.forEach((srcCol) => {
            const destCol = destClass.columns.find((dc) => dc.name === srcCol.name);
            if (!destCol) return;
            const val = sourceClass.cells[`${rowId}:${srcCol.id}`];
            if (val) newCellsForDest[`${newId}:${destCol.id}`] = val;
          });
          const srcReports = sourceClass.reports?.[rowId] || [];
          const remapped = srcReports
            .map((e) => {
              const srcCol = sourceClass.columns.find((c) => c.id === e.colId);
              const destCol = srcCol ? destClass.columns.find((dc) => dc.name === srcCol.name) : null;
              if (!destCol) return null;
              return { ...e, id: uid(), colId: destCol.id, colName: destCol.name, colColor: destCol.color };
            })
            .filter(Boolean);
          if (remapped.length > 0) newReportsForDest[newId] = remapped;
        }
      });

      const updatedDestClass = {
        ...destClass,
        rows: [...destClass.rows, ...newRows],
        cells: { ...destClass.cells, ...newCellsForDest },
        reports: { ...(destClass.reports || {}), ...newReportsForDest },
      };
      const newDestData = { ...destData, classes: destData.classes.map((c) => (c.id === destClassId ? updatedDestClass : c)) };

      const remainingReports = { ...(sourceClass.reports || {}) };
      selectedRowIds.forEach((id) => delete remainingReports[id]);
      const updatedSourceClass = {
        ...sourceClass,
        rows: sourceClass.rows.filter((r) => !selectedRowIds.includes(r.id)),
        cells: Object.fromEntries(Object.entries(sourceClass.cells).filter(([k]) => !selectedRowIds.some((id) => k.startsWith(`${id}:`)))),
        reports: remainingReports,
      };
      const newSourceData = { ...sourceData, classes: sourceData.classes.map((c) => (c.id === sourceClassId ? updatedSourceClass : c)) };

      const { error: e1 } = await supabase.from("user_data").update({ data: newSourceData, updated_at: new Date().toISOString() }).eq("user_id", sourceId);
      if (e1) throw e1;
      const { error: e2 } = await supabase.from("user_data").update({ data: newDestData, updated_at: new Date().toISOString() }).eq("user_id", destId);
      if (e2) throw e2;
      setDone(true);
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Modal title="نقل طلاب بين حسابين" onClose={onClose} wide>
      <p className="text-sm mb-4" style={{ color: MUTED }}>أداة خاصة بالمالك — تنقل طلابًا من فصل بحساب معلم إلى فصل بحساب معلم آخر مباشرة.</p>
      {error && <p className="text-xs mb-3" style={{ color: "#C0392B" }}>{error}</p>}
      {done ? (
        <div className="text-center py-6">
          <p className="text-sm font-bold mb-4" style={{ color: "#26423B" }}>تم النقل بنجاح ✓</p>
          <button onClick={onClose} className="px-5 py-2 rounded-lg text-sm font-bold text-white" style={{ background: "#26423B" }}>إغلاق</button>
        </div>
      ) : (
        <>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
            <div>
              <Field label="من حساب">
                <select value={sourceId} onChange={(e) => setSourceId(e.target.value)} style={inputStyle}>
                  <option value="">اختر معلمًا...</option>
                  {profiles.map((p) => <option key={p.id} value={p.id}>{p.email}</option>)}
                </select>
              </Field>
              {sourceData && (
                <Field label="الفصل">
                  <select value={sourceClassId} onChange={(e) => { setSourceClassId(e.target.value); setSelectedRowIds([]); }} style={inputStyle}>
                    <option value="">اختر فصلًا...</option>
                    {sourceData.classes.map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade} ({c.rows?.length || 0} طالب)</option>)}
                  </select>
                </Field>
              )}
              {sourceClass && (
                <div className="rounded-xl p-2 max-h-52 overflow-y-auto" style={{ border: `1px solid ${LINE}` }}>
                  {sourceClass.rows.length === 0 ? (
                    <p className="text-xs text-center py-3" style={{ color: MUTED }}>لا يوجد طلاب بهذا الفصل.</p>
                  ) : sourceClass.rows.map((r) => (
                    <label key={r.id} className="flex items-center gap-2 py-1.5 text-sm" style={{ color: INK }}>
                      <input type="checkbox" checked={selectedRowIds.includes(r.id)} onChange={() => toggleRow(r.id)} />
                      {r.name}
                    </label>
                  ))}
                </div>
              )}
            </div>
            <div>
              <Field label="إلى حساب">
                <select value={destId} onChange={(e) => setDestId(e.target.value)} style={inputStyle}>
                  <option value="">اختر معلمًا...</option>
                  {profiles.filter((p) => p.id !== sourceId).map((p) => <option key={p.id} value={p.id}>{p.email}</option>)}
                </select>
              </Field>
              {destData && (
                <Field label="الفصل الوجهة">
                  <select value={destClassId} onChange={(e) => setDestClassId(e.target.value)} style={inputStyle}>
                    <option value="">اختر فصلًا...</option>
                    {destData.classes.map((c) => <option key={c.id} value={c.id}>{c.subject} — {c.grade}</option>)}
                  </select>
                </Field>
              )}
              <label className="flex items-start gap-2 text-sm font-medium mt-2" style={{ color: INK }}>
                <input type="checkbox" checked={includeGrades} onChange={(e) => setIncludeGrades(e.target.checked)} className="mt-0.5" />
                <span>
                  نقل الدرجات والرصد معهم
                  <span className="block text-xs font-normal mt-0.5" style={{ color: MUTED }}>يُنقل فقط ما كان له عمود بنفس الاسم بالفصل الوجهة.</span>
                </span>
              </label>
            </div>
          </div>
          <button
            disabled={busy || !sourceClass || !destClass || selectedRowIds.length === 0}
            onClick={executeTransfer}
            className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40"
            style={{ background: "#26423B" }}
          >
            {busy ? "جارٍ النقل..." : `نقل ${selectedRowIds.length} طالب`}
          </button>
        </>
      )}
    </Modal>
  );
}

function AdminPanelModal({ currentUserId, siteSettings, updateSiteSettings, onClose }) {
  const [profiles, setProfiles] = useState(null);
  const [usage, setUsage] = useState({}); // user_id -> { classes, students }
  const [error, setError] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [confirmPromoteId, setConfirmPromoteId] = useState(null);
  const [busyId, setBusyId] = useState(null);
  const [search, setSearch] = useState("");
  const [announcement, setAnnouncement] = useState(siteSettings.announcement || "");
  const [announcementActive, setAnnouncementActive] = useState(!!siteSettings.announcementActive);
  const [siteTagline, setSiteTagline] = useState(siteSettings.siteTagline || "");
  const [resetSentId, setResetSentId] = useState(null);
  const [showTransfer, setShowTransfer] = useState(false);
  const [sortBy, setSortBy] = useState("newest");
  const [tab, setTab] = useState("overview");
  const logoInputRef = useRef(null);

  const loadAll = async () => {
    setError("");
    const { data: profilesData, error: pErr } = await supabase.from("profiles").select("id, email, is_owner, is_disabled, subscription_status, subscription_expires_at, created_at").order("created_at", { ascending: false });
    if (pErr) { setError(pErr.message); return; }
    setProfiles(profilesData || []);

    const { data: userData, error: uErr } = await supabase.from("user_data").select("user_id, data");
    if (uErr) { setError(uErr.message); return; }
    const map = {};
    (userData || []).forEach((row) => {
      const classes = row.data?.classes || [];
      const students = classes.reduce((sum, c) => sum + (c.rows?.length || 0), 0);
      map[row.user_id] = { classes: classes.length, students };
    });
    setUsage(map);
  };

  useEffect(() => { loadAll(); }, []);

  const toggleDisabled = async (p) => {
    setBusyId(p.id);
    const { error } = await supabase.from("profiles").update({ is_disabled: !p.is_disabled }).eq("id", p.id);
    if (error) setError(error.message); else await loadAll();
    setBusyId(null);
  };

  const promoteToOwner = async (id) => {
    setBusyId(id);
    const { error } = await supabase.from("profiles").update({ is_owner: true }).eq("id", id);
    if (error) setError(error.message); else await loadAll();
    setBusyId(null);
    setConfirmPromoteId(null);
  };

  const deleteUserData = async (id) => {
    setBusyId(id);
    const { error } = await supabase.from("user_data").delete().eq("user_id", id);
    if (error) setError(error.message); else await loadAll();
    setBusyId(null);
    setConfirmDeleteId(null);
  };

  const updateSubscription = async (id, patch) => {
    setBusyId(id);
    const { error } = await supabase.from("profiles").update(patch).eq("id", id);
    if (error) setError(error.message); else await loadAll();
    setBusyId(null);
  };

  const sendPasswordReset = async (email, id) => {
    setBusyId(id);
    const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin });
    if (error) setError(error.message); else setResetSentId(id);
    setBusyId(null);
  };

  const saveAnnouncement = () => {
    updateSiteSettings((s) => ({ ...s, announcement: announcement.trim(), announcementActive }));
  };

  const saveSiteIdentity = () => {
    updateSiteSettings((s) => ({ ...s, siteTagline: siteTagline.trim() }));
  };

  const handleLogoUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateSiteSettings((s) => ({ ...s, siteLogo: reader.result }));
    reader.readAsDataURL(file);
    e.target.value = "";
  };

  const exportUsersList = () => {
    const rows = (profiles || []).map((p) => ({
      "البريد الإلكتروني": p.email,
      "الحالة": p.is_disabled ? "معطّل" : "نشط",
      "مالك": p.is_owner ? "نعم" : "لا",
      "حالة الاشتراك": SUBSCRIPTION_LABELS[p.subscription_status] || p.subscription_status,
      "تاريخ انتهاء الاشتراك": p.subscription_expires_at || "",
      "الفصول": usage[p.id]?.classes || 0,
      "الطلاب": usage[p.id]?.students || 0,
      "تاريخ التسجيل": p.created_at?.slice(0, 10) || "",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "المستخدمون");
    XLSX.writeFile(wb, "مستخدمو-فصولي.xlsx");
  };

  const todayISO = todayKey();
  const isExpiringSoon = (p) => {
    if (!p.subscription_expires_at || (p.subscription_status !== "active" && p.subscription_status !== "trial")) return false;
    const days = (new Date(p.subscription_expires_at) - new Date(todayISO)) / 86400000;
    return days >= 0 && days <= 7;
  };
  const isExpired = (p) => p.subscription_expires_at && p.subscription_expires_at < todayISO && (p.subscription_status === "active" || p.subscription_status === "trial");

  const filtered = (profiles || [])
    .filter((p) => p.email?.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => {
      if (sortBy === "oldest") return new Date(a.created_at) - new Date(b.created_at);
      if (sortBy === "classes") return (usage[b.id]?.classes || 0) - (usage[a.id]?.classes || 0);
      if (sortBy === "expiry") return (a.subscription_expires_at || "9999") < (b.subscription_expires_at || "9999") ? -1 : 1;
      return new Date(b.created_at) - new Date(a.created_at);
    });
  const totalUsers = profiles?.length || 0;
  const activeUsers = (profiles || []).filter((p) => !p.is_disabled).length;
  const disabledUsers = totalUsers - activeUsers;
  const totalClasses = Object.values(usage).reduce((s, u) => s + u.classes, 0);
  const totalStudents = Object.values(usage).reduce((s, u) => s + u.students, 0);
  const paidUsers = (profiles || []).filter((p) => p.subscription_status === "active").length;
  const expiringSoonUsers = (profiles || []).filter(isExpiringSoon).length;

  // بيانات نمو التسجيل لآخر 30 يومًا
  const growthData = (() => {
    const days = [];
    for (let i = 29; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const key = d.toISOString().slice(0, 10);
      days.push({ date: key, label: `${d.getDate()}/${d.getMonth() + 1}`, count: 0 });
    }
    (profiles || []).forEach((p) => {
      const key = p.created_at?.slice(0, 10);
      const day = days.find((d) => d.date === key);
      if (day) day.count++;
    });
    return days;
  })();

  return (
    <Modal title="لوحة التحكم" onClose={onClose} lg>
      {error && <p className="text-xs mb-3" style={{ color: "#C0392B" }}>{error}</p>}

      <div className="flex gap-2 mb-4 overflow-x-auto pb-1">
        {[
          { id: "overview", label: "نظرة عامة" },
          { id: "site", label: "إعدادات الموقع" },
          { id: "users", label: "المستخدمون" },
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className="text-sm font-semibold px-4 py-2 rounded-xl shrink-0"
            style={{ background: tab === t.id ? "#26423B" : "transparent", color: tab === t.id ? "#fff" : MUTED, border: `1px solid ${tab === t.id ? "#26423B" : LINE}` }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === "overview" && (
        <>
          <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 mb-4">
            {[
              { label: "المستخدمون", value: totalUsers },
              { label: "نشطون", value: activeUsers },
              { label: "مشتركون", value: paidUsers },
              { label: "الفصول", value: totalClasses },
              { label: "الطلاب", value: totalStudents },
            ].map((s) => (
              <div key={s.label} className="rounded-xl p-3 text-center" style={{ background: "#F8F7F2", border: `1px solid ${LINE}` }}>
                <p className="text-xl font-extrabold" style={{ color: INK }}>{s.value}</p>
                <p className="text-xs" style={{ color: MUTED }}>{s.label}</p>
              </div>
            ))}
          </div>
          <div className="flex gap-4 flex-wrap mb-4">
            {disabledUsers > 0 && <p className="text-xs" style={{ color: "#9A3B2E" }}>{disabledUsers} حساب معطّل حاليًا</p>}
            {expiringSoonUsers > 0 && <p className="text-xs font-semibold" style={{ color: "#C97A2B" }}>⏳ {expiringSoonUsers} اشتراك سينتهي خلال ٧ أيام</p>}
          </div>

          <div className="p-3 rounded-xl mb-4" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-3" style={{ color: INK }}>تسجيلات آخر ٣٠ يومًا</p>
            <div style={{ width: "100%", height: 140 }}>
              <ResponsiveContainer>
                <LineChart data={growthData}>
                  <CartesianGrid strokeDasharray="3 3" stroke={LINE} />
                  <XAxis dataKey="label" tick={{ fontSize: 10 }} interval={4} stroke={MUTED} />
                  <YAxis tick={{ fontSize: 10 }} allowDecimals={false} width={24} stroke={MUTED} />
                  <Tooltip contentStyle={{ fontSize: 12, direction: "rtl" }} labelFormatter={(l) => `يوم ${l}`} formatter={(v) => [v, "تسجيلات"]} />
                  <Line type="monotone" dataKey="count" stroke="#26423B" strokeWidth={2} dot={false} />
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        </>
      )}

      {tab === "site" && (
        <>
          <div className="p-3 rounded-xl mb-4" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-2" style={{ color: INK }}>هوية الموقع</p>
            <div className="flex items-center gap-3 mb-3">
              {siteSettings.siteLogo ? (
                <img src={siteSettings.siteLogo} alt="الشعار" className="w-12 h-12 rounded-xl object-cover" style={{ border: `1px solid ${LINE}` }} />
              ) : (
                <div className="w-12 h-12 rounded-xl flex items-center justify-center" style={{ background: "#26423B" }}><BookOpen size={20} color="#fff" /></div>
              )}
              <input ref={logoInputRef} type="file" accept="image/*" onChange={handleLogoUpload} style={{ display: "none" }} />
              <button onClick={() => logoInputRef.current?.click()} className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ border: `1px solid ${LINE}`, color: INK }}>تغيير الشعار</button>
              {siteSettings.siteLogo && (
                <button onClick={() => updateSiteSettings((s) => ({ ...s, siteLogo: null }))} className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ color: "#C0392B" }}>إزالة</button>
              )}
            </div>
            <input style={inputStyle} value={siteTagline} onChange={(e) => setSiteTagline(e.target.value)} placeholder="الوصف المختصر تحت اسم الموقع بالصفحة الرئيسية" />
            <div className="flex justify-end mt-2">
              <button onClick={saveSiteIdentity} className="text-xs font-bold px-4 py-2 rounded-lg text-white" style={{ background: "#26423B" }}>حفظ</button>
            </div>
          </div>

          <div className="p-3 rounded-xl mb-4" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-2" style={{ color: INK }}>إعلان عام لكل المستخدمين</p>
            <textarea style={{ ...inputStyle, minHeight: 60 }} value={announcement} onChange={(e) => setAnnouncement(e.target.value)} placeholder="مثال: صيانة مجدولة يوم الخميس، أو ميزة جديدة أُضيفت..." />
            <div className="flex items-center justify-between mt-2">
              <label className="flex items-center gap-2 text-sm font-medium" style={{ color: INK }}>
                <input type="checkbox" checked={announcementActive} onChange={(e) => setAnnouncementActive(e.target.checked)} />
                إظهار الإعلان الآن
              </label>
              <button onClick={saveAnnouncement} className="text-xs font-bold px-4 py-2 rounded-lg text-white" style={{ background: "#26423B" }}>حفظ</button>
            </div>
          </div>
        </>
      )}

      {tab === "users" && (
        <>
          <div className="flex items-center gap-2 mb-3 flex-wrap">
            <input style={{ ...inputStyle, flex: 1, minWidth: 160 }} value={search} onChange={(e) => setSearch(e.target.value)} placeholder="ابحث بالبريد الإلكتروني..." />
            <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="text-sm rounded-xl px-3 py-2" style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}>
              <option value="newest">الأحدث تسجيلًا</option>
              <option value="oldest">الأقدم تسجيلًا</option>
              <option value="classes">الأكثر فصولًا</option>
              <option value="expiry">أقرب انتهاء اشتراك</option>
            </select>
            <IconBtn icon={ImageDown} label="تصدير Excel" onClick={exportUsersList} />
            <IconBtn icon={Send} label="نقل طلاب بين حسابين" onClick={() => setShowTransfer(true)} />
          </div>

          {!profiles ? (
            <p className="text-sm text-center py-8" style={{ color: MUTED }}>...جارٍ التحميل</p>
          ) : filtered.length === 0 ? (
            <p className="text-sm text-center py-8" style={{ color: MUTED }}>لا يوجد نتائج.</p>
          ) : (
            <div className="space-y-2">
              {filtered.map((p) => (
                <div key={p.id} className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: p.is_disabled ? "#FBEDEA" : "#fff" }}>
                  <div className="flex items-center gap-2 flex-wrap mb-2">
                    <div className="flex-1 min-w-[160px]">
                      <p className="text-sm font-semibold flex items-center gap-1.5 flex-wrap" style={{ color: INK }}>
                        {p.email}
                        {p.is_owner && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: "#EAF3F0", color: "#26423B" }}>مالك</span>}
                        {p.is_disabled && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: "#FBEDEA", color: "#9A3B2E" }}>معطّل</span>}
                        {isExpired(p) && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: "#FBEDEA", color: "#9A3B2E" }}>اشتراك منتهي</span>}
                        {!isExpired(p) && isExpiringSoon(p) && <span className="text-xs font-bold px-2 py-0.5 rounded-full" style={{ background: "#FCEFE2", color: "#C97A2B" }}>⏳ ينتهي قريبًا</span>}
                        {p.id === currentUserId && <span className="text-xs" style={{ color: MUTED }}>(أنت)</span>}
                      </p>
                      <p className="text-xs" style={{ color: MUTED }}>
                        سجّل بتاريخ {formatDateDisplay(p.created_at?.slice(0, 10))} • {usage[p.id]?.classes || 0} فصل • {usage[p.id]?.students || 0} طالب
                      </p>
                    </div>
                    {!p.is_owner && p.id !== currentUserId && (
                      <div className="flex gap-2 shrink-0 flex-wrap">
                        <button disabled={busyId === p.id} onClick={() => toggleDisabled(p)} className="text-xs font-bold px-3 py-1.5 rounded-lg disabled:opacity-40" style={{ color: p.is_disabled ? "#26423B" : "#9A3B2E", border: `1px solid ${p.is_disabled ? "#C9E2DB" : "#F0D2CB"}` }}>
                          {p.is_disabled ? "تفعيل" : "تعطيل"}
                        </button>
                        <button disabled={busyId === p.id} onClick={() => setConfirmPromoteId(p.id)} className="text-xs font-bold px-3 py-1.5 rounded-lg disabled:opacity-40" style={{ color: "#26423B", border: "1px solid #C9E2DB" }}>
                          ترقية لمالك
                        </button>
                        <button disabled={busyId === p.id} onClick={() => setConfirmDeleteId(p.id)} className="text-xs font-bold px-3 py-1.5 rounded-lg disabled:opacity-40" style={{ color: "#C0392B", border: "1px solid #F0D2CB" }}>
                          حذف البيانات
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center gap-2 flex-wrap pt-2" style={{ borderTop: `1px solid ${LINE}` }}>
                    <select
                      value={p.subscription_status || "free"}
                      onChange={(e) => updateSubscription(p.id, { subscription_status: e.target.value })}
                      disabled={busyId === p.id}
                      className="text-xs rounded-lg px-2 py-1.5"
                      style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}
                    >
                      {Object.entries(SUBSCRIPTION_LABELS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
                    </select>
                    <input
                      type="date"
                      value={p.subscription_expires_at || ""}
                      onChange={(e) => updateSubscription(p.id, { subscription_expires_at: e.target.value || null })}
                      disabled={busyId === p.id}
                      className="text-xs rounded-lg px-2 py-1.5"
                      style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}
                      title="تاريخ انتهاء الاشتراك"
                    />
                    <button disabled={busyId === p.id} onClick={() => sendPasswordReset(p.email, p.id)} className="text-xs font-semibold px-3 py-1.5 rounded-lg mr-auto" style={{ color: "#26423B", border: "1px solid #C9E2DB" }}>
                      {resetSentId === p.id ? "أُرسل ✓" : "إرسال رابط إعادة تعيين كلمة المرور"}
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}
      {confirmDeleteId && (
        <ConfirmDialog
          title="حذف بيانات الحساب"
          message="سيُحذف كل ما لدى هذا المستخدم من فصول وطلاب ورصد بشكل نهائي، بدون إمكانية استرجاع. حسابه نفسه يبقى موجودًا ويقدر يدخل بحساب فارغ. متابعة؟"
          onCancel={() => setConfirmDeleteId(null)}
          onConfirm={() => deleteUserData(confirmDeleteId)}
        />
      )}
      {confirmPromoteId && (
        <ConfirmDialog
          title="ترقية إلى مالك"
          message="سيحصل هذا الحساب على كل صلاحيات المالك (تعديل تذييل الموقع، تعطيل/حذف أي حساب آخر، ترقية حسابات أخرى). لا يمكن التراجع عن هذا إلا يدويًا من قاعدة البيانات. متابعة؟"
          confirmLabel="ترقية"
          onCancel={() => setConfirmPromoteId(null)}
          onConfirm={() => promoteToOwner(confirmPromoteId)}
        />
      )}
      {showTransfer && <TransferStudentsModal profiles={profiles || []} onClose={() => setShowTransfer(false)} />}
    </Modal>
  );
}

function SettingsModal({ feedback, onToggleFeedback, darkMode, onToggleDarkMode, themeColor, density, fontScale, schoolName, principalName, countryName, ministryName, logoImage, teacherPhoto, onChangeSchoolInfo, footerContacts, footerBadges, onAddContact, onUpdateContact, onRemoveContact, onAddBadge, onRemoveBadge, userEmail, onSignOut, isOwner, onBackupData, onRestoreData, onClose }) {
  const logoInputRef = useRef(null);
  const teacherPhotoInputRef = useRef(null);
  const badgeInputRef = useRef(null);
  const restoreInputRef = useRef(null);
  const [confirmRemoveLogo, setConfirmRemoveLogo] = useState(false);
  const [pendingRestore, setPendingRestore] = useState(null);
  const [tab, setTab] = useState("general");
  const handleRestoreFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        setPendingRestore(parsed);
      } catch (err) {
        alert("تعذّرت قراءة الملف — تأكد إنه نسخة احتياطية صحيحة بصيغة JSON.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };
  const handleLogoChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChangeSchoolInfo({ logoImage: reader.result });
    reader.readAsDataURL(file);
    e.target.value = "";
  };
  const handleTeacherPhotoChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onChangeSchoolInfo({ teacherPhoto: reader.result });
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
      <div className="flex gap-2 mb-4 overflow-x-auto pb-1">
        {[
          { id: "general", label: "عام" },
          { id: "school", label: "بيانات المدرسة" },
          ...(isOwner ? [{ id: "footer", label: "تذييل الصفحة" }] : []),
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className="text-sm font-semibold px-4 py-2 rounded-xl shrink-0"
            style={{ background: tab === t.id ? "#26423B" : "transparent", color: tab === t.id ? "#fff" : MUTED, border: `1px solid ${tab === t.id ? "#26423B" : LINE}` }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === "general" && (
        <>
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
              {darkMode ? <Moon size={18} color="#26423B" /> : <Sun size={18} color={MUTED} />}
              <div>
                <p className="text-sm font-semibold" style={{ color: INK }}>الوضع الليلي</p>
                <p className="text-xs" style={{ color: MUTED }}>ألوان داكنة أريح للعين في الإضاءة الخافتة</p>
              </div>
            </div>
            <button
              onClick={onToggleDarkMode}
              className="w-11 h-6 rounded-full shrink-0 relative transition-colors"
              style={{ background: darkMode ? "#26423B" : LINE }}
            >
              <span className="absolute top-0.5 w-5 h-5 rounded-full bg-white shadow transition-all" style={{ [darkMode ? "right" : "left"]: "2px" }} />
            </button>
          </div>
          <div className="flex items-center justify-between p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <div className="flex items-center gap-2">
              {feedback ? <Volume2 size={18} color="#26423B" /> : <VolumeX size={18} color={MUTED} />}
              <div>
                <p className="text-sm font-semibold" style={{ color: INK }}>تنبيه صوتي واهتزاز عند الرصد</p>
                <p className="text-xs" style={{ color: MUTED }}>تأكيد سريع (صوت + اهتزاز خفيف) كل مرة تسجّل غيابًا أو قيمة</p>
              </div>
            </div>
            <button
              onClick={onToggleFeedback}
              className="w-11 h-6 rounded-full shrink-0 relative transition-colors"
              style={{ background: feedback ? "#26423B" : LINE }}
            >
              <span className="absolute top-0.5 w-5 h-5 rounded-full bg-white shadow transition-all" style={{ [feedback ? "right" : "left"]: "2px" }} />
            </button>
          </div>
          <div className="p-3 rounded-xl mt-3" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <div className="flex items-center gap-2 mb-2">
              <Palette size={18} color={themeColor || "#26423B"} />
              <div>
                <p className="text-sm font-semibold" style={{ color: INK }}>لون الهوية الشخصية</p>
                <p className="text-xs" style={{ color: MUTED }}>يُطبَّق على الأزرار الرئيسية والتبويبات بكل التطبيق</p>
              </div>
            </div>
            <ColorSwatches value={themeColor || "#26423B"} onChange={(c) => onChangeSchoolInfo({ themeColor: c })} />
          </div>
          <div className="p-3 rounded-xl mt-3" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-2" style={{ color: INK }}>كثافة عرض الجدول</p>
            <div className="flex gap-2">
              <button
                onClick={() => onChangeSchoolInfo({ density: "comfortable" })}
                className="flex-1 text-xs font-semibold py-2 rounded-lg"
                style={{ background: density !== "compact" ? "#26423B" : "transparent", color: density !== "compact" ? "#fff" : MUTED, border: `1px solid ${density !== "compact" ? "#26423B" : LINE}` }}
              >مريح</button>
              <button
                onClick={() => onChangeSchoolInfo({ density: "compact" })}
                className="flex-1 text-xs font-semibold py-2 rounded-lg"
                style={{ background: density === "compact" ? "#26423B" : "transparent", color: density === "compact" ? "#fff" : MUTED, border: `1px solid ${density === "compact" ? "#26423B" : LINE}` }}
              >مضغوط (طلاب أكثر بنفس الشاشة)</button>
            </div>
          </div>
          <div className="p-3 rounded-xl mt-3" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-2" style={{ color: INK }}>حجم الخط</p>
            <div className="flex gap-2">
              {[{ v: 1, l: "عادي" }, { v: 1.1, l: "أكبر قليلًا" }, { v: 1.25, l: "أكبر" }].map((opt) => (
                <button
                  key={opt.v}
                  onClick={() => onChangeSchoolInfo({ fontScale: opt.v })}
                  className="flex-1 text-xs font-semibold py-2 rounded-lg"
                  style={{ background: (fontScale || 1) === opt.v ? "#26423B" : "transparent", color: (fontScale || 1) === opt.v ? "#fff" : MUTED, border: `1px solid ${(fontScale || 1) === opt.v ? "#26423B" : LINE}` }}
                >{opt.l}</button>
              ))}
            </div>
          </div>
          <div className="p-3 rounded-xl mt-3" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-1" style={{ color: INK }}>نسخة احتياطية</p>
            <p className="text-xs mb-3" style={{ color: MUTED }}>طمأنينة إضافية غير الاعتماد على السحابة بس — ملف واحد يشمل كل فصولك وطلابك وشواهدك.</p>
            <div className="flex flex-wrap gap-2">
              <IconBtn icon={Download} label="تنزيل نسخة احتياطية كاملة" onClick={onBackupData} />
              <input ref={restoreInputRef} type="file" accept=".json" onChange={handleRestoreFile} style={{ display: "none" }} />
              <IconBtn icon={FolderOpen} label="استعادة من نسخة" onClick={() => restoreInputRef.current?.click()} />
            </div>
          </div>
        </>
      )}

      {tab === "school" && (
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
          <Field label="صورتك الشخصية (اختياري)" hint="تظهر في التقارير والشهادات عند تفعيل خيار تضمينها.">
            <input ref={teacherPhotoInputRef} type="file" accept="image/*" onChange={handleTeacherPhotoChange} style={{ display: "none" }} />
            <div className="flex items-center gap-2">
              {teacherPhoto ? (
                <img src={teacherPhoto} alt="صورتك" className="w-10 h-10 rounded-full object-cover dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />
              ) : (
                <div className="w-10 h-10 rounded-full flex items-center justify-center" style={{ background: "#F3F1E9" }}><User size={16} color={MUTED} /></div>
              )}
              <button type="button" onClick={() => teacherPhotoInputRef.current?.click()} className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold" style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}>
                <Camera size={15} color="#26423B" /> {teacherPhoto ? "استبدال الصورة" : "رفع صورة"}
              </button>
              {teacherPhoto && (
                <button type="button" onClick={() => onChangeSchoolInfo({ teacherPhoto: null })} title="إزالة" className="p-1.5 rounded hover:bg-black/5">
                  <ImageOff size={15} color={MUTED} />
                </button>
              )}
            </div>
          </Field>
          <Field label="شعار الوزارة / المدرسة (اختياري)" hint="ارفع الشعار الرسمي الذي تملكه — لا يمكن للتطبيق توليد الشعارات الحكومية تلقائيًا.">
            <input ref={logoInputRef} type="file" accept="image/*" onChange={handleLogoChange} style={{ display: "none" }} />
            <div className="flex items-center gap-2">
              <button type="button" onClick={() => logoInputRef.current?.click()} className="flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-semibold" style={{ border: `1px solid ${LINE}`, color: INK, background: "#fff" }}>
                <Camera size={15} color="#26423B" /> {logoImage ? "استبدال الشعار" : "رفع شعار"}
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
      )}

      {tab === "footer" && (
        isOwner ? (
          <div className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <p className="text-sm font-semibold mb-1" style={{ color: INK }}>تذييل الصفحة الرئيسية <span className="text-xs font-normal px-2 py-0.5 rounded-full" style={{ background: "#EAF3F0", color: "#26423B" }}>مالك الموقع</span></p>
            <p className="text-xs mb-3" style={{ color: MUTED }}>بيانات التواصل وشهادات الثقة/الاعتماد التي تظهر أسفل الصفحة الرئيسية لجميع المشتركين — تتحكم فيها إضافةً وحذفًا بالكامل.</p>

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
            <button onClick={onAddContact} className="text-xs font-semibold flex items-center gap-1 mb-4" style={{ color: "#26423B" }}>
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
            <button onClick={() => badgeInputRef.current?.click()} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}>
              <Plus size={13} /> إضافة شهادة/شعار ثقة
            </button>
          </div>
        ) : (
          <div className="p-3 rounded-xl flex items-start gap-2" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
            <Info size={15} color={MUTED} className="shrink-0 mt-0.5" />
            <p className="text-xs" style={{ color: MUTED }}>تذييل الصفحة الرئيسية وشهادات الثقة خاصة بمالك الموقع فقط، ولا تظهر لباقي المشتركين.</p>
          </div>
        )
      )}
      {confirmRemoveLogo && (
        <ConfirmDialog
          title="إزالة الشعار"
          message="سيُحذف الشعار المرفوع من الإعدادات، وستحتاج لرفعه مجددًا لاستخدامه في الشهادات. متابعة؟"
          confirmLabel="إزالة"
          onCancel={() => setConfirmRemoveLogo(false)}
          onConfirm={() => { onChangeSchoolInfo({ logoImage: null }); setConfirmRemoveLogo(false); }}
        />
      )}
      {pendingRestore && (
        <ConfirmDialog
          title="استعادة من نسخة احتياطية"
          message="سيتم استبدال كل بياناتك الحالية (الفصول، الطلاب، الشواهد...) بمحتوى هذا الملف بالكامل، ولا يمكن التراجع عن هذا إلا لو عندك نسخة احتياطية ثانية أحدث. متابعة؟"
          confirmLabel="استبدال والاستعادة"
          onCancel={() => setPendingRestore(null)}
          onConfirm={() => { onRestoreData(pendingRestore); setPendingRestore(null); }}
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
      <div className="relative mx-auto mb-5" style={{ width: 340, height: 340 }}>
        <div style={{ position: "absolute", top: -8, left: "50%", transform: "translateX(-50%)", zIndex: 10, filter: "drop-shadow(0 3px 3px rgba(0,0,0,0.25))" }}>
          <svg width="36" height="44" viewBox="0 0 34 42">
            <path d="M17 0C7.6 0 0 7.6 0 17c0 12.75 17 25 17 25s17-12.25 17-25C34 7.6 26.4 0 17 0z" fill="#232622" />
            <circle cx="17" cy="17" r="7" fill="#fff" />
          </svg>
        </div>
        <div
          className="rounded-full relative overflow-hidden"
          style={{
            width: 340, height: 340,
            background: conicGradient,
            border: "5px solid #fff",
            boxShadow: "0 8px 28px rgba(0,0,0,0.2)",
            transform: `rotate(${rotation}deg)`,
            transition: spinning ? "transform 3.5s cubic-bezier(0.17, 0.67, 0.12, 0.99)" : "none",
          }}
        >
          {rows.map((row, i) => {
            const midAngle = i * segmentAngle + segmentAngle / 2;
            const rad = ((midAngle - 90) * Math.PI) / 180;
            const labelRadius = 340 / 2 * 0.66;
            const x = 170 + labelRadius * Math.cos(rad);
            const y = 170 + labelRadius * Math.sin(rad);
            const labelFontSize = rows.length > 20 ? 10 : rows.length > 14 ? 11 : rows.length > 8 ? 13 : 15;
            const chipMaxWidth = rows.length > 14 ? 56 : rows.length > 8 ? 70 : 90;
            return (
              <div
                key={row.id}
                style={{
                  position: "absolute", left: x, top: y,
                  transform: `translate(-50%, -50%) rotate(${-rotation}deg)`,
                  transition: spinning ? "transform 3.5s cubic-bezier(0.17, 0.67, 0.12, 0.99)" : "none",
                }}
              >
                <span style={{
                  display: "block", maxWidth: chipMaxWidth,
                  fontSize: labelFontSize, fontWeight: 800, color: "#fff",
                  textAlign: "center", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                  textShadow: "0 1px 2px rgba(0,0,0,0.45), 0 0 4px rgba(0,0,0,0.25)",
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
        <div className="rounded-2xl p-5 mb-4 text-center celebrate-in" style={{ background: `${wheelColors[rows.findIndex((r) => r.id === picked.id)]}14`, border: `2px solid ${wheelColors[rows.findIndex((r) => r.id === picked.id)]}` }}>
          <p className="text-xs font-bold mb-2" style={{ color: wheelColors[rows.findIndex((r) => r.id === picked.id)] }}>🎉 الطالب المختار</p>
          {picked.photo ? (
            <img src={picked.photo} alt={picked.name} className="w-16 h-16 rounded-full object-cover mx-auto mb-2 dark-mode-img-fix" style={{ border: `3px solid ${wheelColors[rows.findIndex((r) => r.id === picked.id)]}` }} />
          ) : (
            <div className="w-16 h-16 rounded-full mx-auto mb-2 flex items-center justify-center font-extrabold text-2xl text-white" style={{ background: wheelColors[rows.findIndex((r) => r.id === picked.id)] }}>
              {(picked.name || "؟").trim().charAt(0)}
            </div>
          )}
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
            <button onClick={generate} className="flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-bold text-white" style={{ background: "#26423B" }}>
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
    try { await document.fonts.load("bold 58px 'Aref Ruqaa'"); } catch (e) { /* falls back gracefully */ }
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
          <div className="w-4 h-4 rounded-full border-2 shrink-0" style={{ borderColor: "#26423B transparent #26423B #26423B", animation: "spin 0.8s linear infinite" }} />
          <p className="text-xs font-semibold" style={{ color: INK }}>جارٍ توليد الشهادات... {progress}%</p>
        </div>
      ) : (
        <button onClick={generateAndDownloadZip} className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-bold text-white" style={{ background: "#26423B" }}>
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
  {
    id: "behavior",
    name: "السلوك والانضباط الصفي",
    weaknessSigns: "كثرة الحركة والانشغال عن الدرس، مقاطعة المعلم والزملاء، صعوبة الالتزام بتعليمات الصف، وتكرار المشكلات مع الزملاء.",
    causes: "ضعف في مهارات ضبط النفس، رغبة في لفت الانتباه، أو صعوبة في فهم حدود السلوك المقبول داخل الصف.",
    goals: "أن يلتزم الطالب بقوانين الصف الأساسية لفترة زمنية محددة، وأن يقل عدد الملاحظات السلوكية بشكل ملحوظ.",
    methods: "نظام تعزيز فوري وواضح للسلوك الإيجابي، اتفاق مسبق مع الطالب على قواعد بسيطة ومفهومة، تجاهل السلوكيات البسيطة غير المؤذية، وتواصل منتظم مع ولي الأمر لمتابعة نفس النهج بالمنزل.",
  },
  {
    id: "attention",
    name: "التركيز والانتباه",
    weaknessSigns: "شرود ذهني متكرر أثناء الشرح، صعوبة إنهاء المهمة الواحدة، نسيان التعليمات بعد وقت قصير، وتشتت سريع بأقل مؤثر خارجي.",
    causes: "قصر مدى الانتباه الطبيعي للمرحلة العمرية، بيئة صفية مشتتة، أو إرهاق وقلة نوم كافٍ.",
    goals: "زيادة المدة الزمنية التي يستطيع الطالب التركيز فيها على مهمة واحدة تدريجيًا، وإنهاء المهام القصيرة دون تذكير متكرر.",
    methods: "تقسيم المهمة لأجزاء زمنية قصيرة مع فواصل راحة، تقليل المشتتات القريبة من الطالب، تعليمات واضحة ومختصرة خطوة بخطوة، وتعزيز فوري عند إتمام أي جزء من المهمة.",
  },
  {
    id: "handwriting",
    name: "الخط والتنظيم الكتابي",
    weaknessSigns: "خط غير واضح أو غير منتظم، صعوبة محاذاة الكتابة على السطر، بطء شديد في الكتابة اليدوية، وعدم ترتيب الدفاتر والواجبات.",
    causes: "ضعف في المهارات الحركية الدقيقة، قلة التدريب على الكتابة اليدوية، أو إمساك غير صحيح للقلم.",
    goals: "تحسين وضوح الخط وانتظامه على السطر، وزيادة سرعة الكتابة تدريجيًا دون التأثير على وضوحها.",
    methods: "تمارين قصيرة يومية على تتبع الحروف والكلمات، التأكد من طريقة إمساك القلم الصحيحة، استخدام أوراق مسطّرة بمساعدات بصرية، وتعزيز إيجابي لأي تحسن في الترتيب والوضوح.",
  },
  {
    id: "social",
    name: "المهارات الاجتماعية",
    weaknessSigns: "صعوبة في تكوين صداقات، انعزال عن الأنشطة الجماعية، صعوبة في التعبير عن المشاعر بطريقة مناسبة، وتكرار الخلافات مع الزملاء.",
    causes: "قلة الفرص للتفاعل الاجتماعي الإيجابي، خجل زائد، أو ضعف في مهارات التواصل والتعبير.",
    goals: "أن يشارك الطالب في نشاط جماعي واحد على الأقل بانتظام، وأن يعبّر عن احتياجاته بطريقة لفظية مناسبة.",
    methods: "إشراك الطالب تدريجيًا في أنشطة جماعية صغيرة، نمذجة طرق التواصل الإيجابي، تعزيز فوري لأي محاولة تفاعل إيجابية، والتنسيق مع ولي الأمر لتوفير فرص اجتماعية مشابهة خارج المدرسة.",
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
                  style={{ background: templateId === t.id ? "#26423B" : "transparent", color: templateId === t.id ? "#fff" : INK, border: `1px solid ${templateId === t.id ? "#26423B" : LINE}` }}
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
      try { await document.fonts.load("bold 58px 'Aref Ruqaa'"); } catch (e) { /* font not available, canvas falls back gracefully */ }
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
                <button onClick={() => onRestore(entry.id)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-bold text-white shrink-0" style={{ background: "#26423B" }}>
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

// نافذة تفعيل وضع الاختبار: تختار مدة فيقفل خلالها كل أزرار إضافة/تعديل/حذف
// الأعمدة والصفوف بالفصل (رصد الدرجات والغياب يبقى شغّالًا طبيعيًا).
function ExamModeModal({ onClose, onActivate }) {
  const PRESETS = [15, 30, 45, 60, 90];
  const [minutes, setMinutes] = useState(45);
  return (
    <Modal title="تفعيل وضع الاختبار" onClose={onClose} accent="magic">
      <div className="flex items-start gap-2 p-3 rounded-xl mb-4" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
        <Lock size={15} color={DASH_GREEN} className="shrink-0 mt-0.5" />
        <p className="text-xs" style={{ color: DASH_GREEN }}>
          يقفل أزرار إضافة/تعديل/حذف الأعمدة والصفوف والحذف الجماعي لفترة تحددها، لمنع أي لمسة خطأ أثناء تجوّلك بالفصل. رصد الدرجات وتسجيل الغياب يبقيان شغّالين بشكل طبيعي.
        </p>
      </div>
      <Field label="المدة">
        <div className="grid grid-cols-3 gap-2 mb-2">
          {PRESETS.map((m) => (
            <button
              key={m}
              onClick={() => setMinutes(m)}
              className="py-2 rounded-lg text-sm font-semibold"
              style={{ background: minutes === m ? DASH_GREEN : "#fff", color: minutes === m ? "#fff" : INK, border: `1px solid ${minutes === m ? DASH_GREEN : LINE}` }}
            >
              {m} دقيقة
            </button>
          ))}
        </div>
        <input
          type="number"
          min={1}
          value={minutes}
          onChange={(e) => setMinutes(Math.max(1, Number(e.target.value) || 1))}
          style={inputStyle}
          placeholder="مدة مخصّصة بالدقائق"
        />
      </Field>
      <button
        onClick={() => onActivate(minutes)}
        className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-bold text-white mt-2 transition-all hover:brightness-110"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
      >
        <Lock size={15} /> تفعيل القفل لمدة {minutes} دقيقة
      </button>
    </Modal>
  );
}

// شريط تنبيه ثابت أعلى صفحة الفصل يظهر أثناء تفعيل وضع الاختبار، مع عدّاد
// تنازلي مباشر وزر لإنهاء القفل يدويًا قبل انتهاء المدة.
function ExamModeBanner({ examModeUntil, onUnlock }) {
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, []);
  const remainingMs = new Date(examModeUntil).getTime() - now;
  if (remainingMs <= 0) return null;
  const mins = Math.floor(remainingMs / 60000);
  const secs = Math.floor((remainingMs % 60000) / 1000);
  return (
    <div className="flex items-center justify-between gap-2 px-3 py-2 rounded-xl mb-1.5" style={{ background: `linear-gradient(135deg, ${DASH_GREEN}, ${DASH_GREEN_DARK})`, color: "#fff" }}>
      <span className="flex items-center gap-2 text-xs font-bold">
        <Lock size={14} /> وضع الاختبار مفعّل — يُفتح تلقائيًا خلال {mins}:{String(secs).padStart(2, "0")}
      </span>
      <button onClick={onUnlock} className="text-xs font-bold px-3 py-1 rounded-lg shrink-0 hover:bg-white/10" style={{ background: "rgba(255,255,255,0.18)" }}>
        <Unlock size={12} className="inline ml-1" /> إنهاء الآن
      </button>
    </div>
  );
}

// يفحص هل صف معيّن يطابق "قاعدة تلوين الصف" المحددة بالفصل (نسبة حضور أو
// قيمة عمود عدّاد)، ويرجّع لون التمييز إذا تحقق الشرط، أو null إذا لا.
function rowColorRuleMatch(cls, row) {
  const rule = cls.rowColorRule;
  if (!rule || !rule.color) return null;
  let actual = null;
  if (rule.type === "attendance") {
    actual = attendancePercent(cls, row.id);
  } else if (rule.type === "column" && rule.colId) {
    const raw = cls.cells[`${row.id}:${rule.colId}`] ?? lastReportedValue(cls, row.id, rule.colId);
    actual = raw === null || raw === undefined || raw === "" ? null : Number(raw);
  }
  if (actual === null || Number.isNaN(actual)) return null;
  const threshold = Number(rule.value);
  const matched =
    rule.operator === "lt" ? actual < threshold :
    rule.operator === "gt" ? actual > threshold :
    rule.operator === "eq" ? actual === threshold : false;
  return matched ? rule.color : null;
}

function RowColorRuleModal({ cls, onClose, onSave, onClear }) {
  const counterColumns = cls.columns.filter((c) => c.type === "counter");
  const initial = cls.rowColorRule;
  const [type, setType] = useState(initial?.type || "attendance");
  const [colId, setColId] = useState(initial?.colId || counterColumns[0]?.id || "");
  const [operator, setOperator] = useState(initial?.operator || "lt");
  const [value, setValue] = useState(initial?.value ?? 80);
  const [color, setColor] = useState(initial?.color || "#F9D8D8");

  return (
    <Modal title="تلوين الصف تلقائيًا حسب شرط" onClose={onClose} accent="magic">
      <p className="text-xs mb-4" style={{ color: MUTED }}>
        يُلوَّن الصف كاملًا تلقائيًا عند تحقق الشرط — يلفت انتباهك فورًا بدون فحص كل عمود يدويًا (مثال: نسبة حضور أقل من ٨٠٪).
      </p>
      <Field label="نوع الشرط">
        <div className="flex gap-2">
          <button onClick={() => setType("attendance")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: type === "attendance" ? INK : "transparent", color: type === "attendance" ? "#fff" : MUTED, border: `1px solid ${type === "attendance" ? INK : LINE}` }}>نسبة الحضور</button>
          <button onClick={() => setType("column")} disabled={counterColumns.length === 0} className="flex-1 text-xs font-semibold py-2 rounded-lg disabled:opacity-40" style={{ background: type === "column" ? INK : "transparent", color: type === "column" ? "#fff" : MUTED, border: `1px solid ${type === "column" ? INK : LINE}` }}>قيمة عمود</button>
        </div>
      </Field>
      {type === "column" && (
        <Field label="العمود">
          <select value={colId} onChange={(e) => setColId(e.target.value)} style={inputStyle}>
            {counterColumns.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </Field>
      )}
      <div className="grid grid-cols-2 gap-2 mb-1">
        <Field label="الشرط">
          <select value={operator} onChange={(e) => setOperator(e.target.value)} style={inputStyle}>
            <option value="lt">أقل من</option>
            <option value="gt">أكبر من</option>
            <option value="eq">يساوي</option>
          </select>
        </Field>
        <Field label="القيمة">
          <input type="number" value={value} onChange={(e) => setValue(Number(e.target.value) || 0)} style={inputStyle} />
        </Field>
      </div>
      <Field label="لون التمييز">
        <ColorSwatches value={color} onChange={setColor} />
      </Field>
      <div className="flex justify-between items-center mt-4">
        <div>
          {initial && <button onClick={() => { onClear(); onClose(); }} className="px-3 py-2 rounded-lg text-sm font-medium" style={{ color: "#9A3B2E" }}>إزالة القاعدة</button>}
        </div>
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-lg text-sm font-medium" style={{ color: MUTED }}>إلغاء</button>
          <button
            onClick={() => { onSave({ type, colId: type === "column" ? colId : null, operator, value, color }); onClose(); }}
            className="px-5 py-2 rounded-lg text-sm font-bold text-white"
            style={{ background: "#26423B" }}
          >حفظ</button>
        </div>
      </div>
    </Modal>
  );
}

// ---------- الأدوات التفاعلية (ألعاب HTML مربوطة بالفصل) ----------
//
// أي لعبة (جاهزة من التطبيق، أو رفعها المعلم من المكتبة) لازم تنادي هذا
// السطر بالضبط عند إجابة صحيحة، حتى تُحتسب النقطة للطالب النشط تلقائيًا:
//   window.parent.postMessage({ type: "fasooli:answer", correct: true }, "*");
// ولو الإجابة غلط (اختياري، ما يؤثر على الدرجة):
//   window.parent.postMessage({ type: "fasooli:answer", correct: false }, "*");

// لعبة جاهزة مدمجة بالتطبيق كمثال حي — أسئلة اختيار من متعدد سريعة، تطبّق
// نفس البروتوكول أعلاه بالضبط.
// يبني صفحة لعبة "الأسئلة السريعة" ديناميكيًا حسب أي مجموعة أسئلة تُمرَّر
// له — سواء من بنك أسئلة جاهز حسب المادة، أو أسئلة كتبها المعلم بنفسه.
function buildQuizGameHtml(questions, gameTitle) {
  const safeJson = JSON.stringify(questions).replace(/</g, "\\u003c").replace(/`/g, "\\`");
  return `<!DOCTYPE html>
<html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { font-family: Tahoma, Arial, sans-serif; background: #FAF8F3; margin: 0; padding: 20px; display: flex; align-items: center; justify-content: center; min-height: 100vh; }
  .card { background: #fff; border-radius: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); padding: 28px; width: 100%; max-width: 460px; text-align: center; }
  .badge { display: inline-block; background: #26423B; color: #fff; font-size: 12px; font-weight: bold; padding: 5px 14px; border-radius: 999px; margin-bottom: 16px; }
  h1 { font-size: 19px; color: #232622; margin: 0 0 24px; line-height: 1.5; }
  .options { display: flex; flex-direction: column; gap: 10px; }
  button.opt { font-family: inherit; font-size: 15px; padding: 14px; border-radius: 12px; border: 2px solid #E4DFD2; background: #fff; cursor: pointer; transition: all 0.15s; }
  button.opt:active { transform: scale(0.97); }
  button.opt.correct { background: #E3F1EC; border-color: #0F9D58; color: #0F9D58; font-weight: bold; }
  button.opt.wrong { background: #FBEAE7; border-color: #C0392B; color: #C0392B; }
  .feedback { margin-top: 16px; font-size: 14px; font-weight: bold; min-height: 20px; }
  .score { position: absolute; top: 16px; left: 16px; font-size: 12px; color: #7A7768; }
</style></head>
<body>
  <div class="score" id="score">النقاط: 0</div>
  <div class="card">
    <span class="badge" id="qnum">سؤال 1</span>
    <h1 id="question"></h1>
    <div class="options" id="options"></div>
    <p class="feedback" id="feedback"></p>
  </div>
<script>
  const QUESTIONS = ${safeJson};
  let idx = Math.floor(Math.random() * QUESTIONS.length);
  let score = 0, qCount = 0, locked = false;
  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  }
  function render() {
    locked = false; qCount++;
    const item = QUESTIONS[idx];
    document.getElementById("qnum").textContent = "سؤال " + qCount;
    document.getElementById("question").textContent = item.q;
    document.getElementById("feedback").textContent = "";
    const optsDiv = document.getElementById("options");
    optsDiv.innerHTML = "";
    const order = shuffle(item.opts.map((t, i) => ({ t, correct: i === item.correct })));
    order.forEach((o) => {
      const btn = document.createElement("button");
      btn.className = "opt"; btn.textContent = o.t;
      btn.onclick = () => choose(btn, o.correct);
      optsDiv.appendChild(btn);
    });
  }
  function choose(btn, isCorrect) {
    if (locked) return;
    locked = true;
    btn.classList.add(isCorrect ? "correct" : "wrong");
    document.getElementById("feedback").textContent = isCorrect ? "✓ إجابة صحيحة! نقطة للطالب النشط 🎉" : "✗ إجابة خاطئة";
    document.getElementById("feedback").style.color = isCorrect ? "#0F9D58" : "#C0392B";
    if (isCorrect) { score++; document.getElementById("score").textContent = "النقاط: " + score; }
    window.parent.postMessage({ type: "fasooli:answer", correct: isCorrect }, "*");
    setTimeout(() => { idx = Math.floor(Math.random() * QUESTIONS.length); render(); }, 1400);
  }
  render();
</script>
</body></html>`;
}

// بنوك أسئلة عامة جاهزة حسب المادة — أسئلة ثقافة عامة بأسلوب كل مادة، مو
// محتوى مطابق حرفيًا للمنهج (ما نملك وصول لمحتوى الكتب الرسمية)، لكنها
// نقطة بداية جاهزة تغني عن الكتابة من الصفر.
const SUBJECT_QUESTION_BANKS = {
  "لغتي الخالدة": [
    { q: "أي مما يلي حرف من حروف الجر؟", opts: ["في", "قام", "كتاب", "جميل"], correct: 0 },
    { q: "ما جمع كلمة \"قلم\"؟", opts: ["أقلام", "قلمون", "قلمين", "قلوم"], correct: 0 },
    { q: "ما نوع الأسلوب في جملة \"ما أجمل السماء!\"؟", opts: ["أسلوب استفهام", "أسلوب تعجب", "أسلوب نهي", "أسلوب أمر"], correct: 1 },
    { q: "ما مرادف كلمة \"سعيد\"؟", opts: ["حزين", "فرح", "غاضب", "خائف"], correct: 1 },
    { q: "ما ضد كلمة \"كبير\"؟", opts: ["طويل", "صغير", "عريض", "قصير"], correct: 1 },
    { q: "أي الكلمات التالية اسم؟", opts: ["يكتب", "مدرسة", "في", "هل"], correct: 1 },
  ],
  "الرياضيات": [
    { q: "ما ناتج 8 × 7؟", opts: ["54", "56", "58", "64"], correct: 1 },
    { q: "كم يساوي نصف العدد 90؟", opts: ["40", "45", "50", "35"], correct: 1 },
    { q: "ما هو العدد الأولي من بين التالي؟", opts: ["9", "15", "17", "21"], correct: 2 },
    { q: "كم ضلعًا للمثلث؟", opts: ["اثنان", "ثلاثة", "أربعة", "خمسة"], correct: 1 },
    { q: "ما ناتج 144 ÷ 12؟", opts: ["10", "11", "12", "13"], correct: 2 },
    { q: "أي الكسور التالية يساوي النصف؟", opts: ["1/3", "2/4", "1/5", "3/5"], correct: 1 },
  ],
  "العلوم": [
    { q: "ما الغاز الذي يتنفسه الإنسان للحياة؟", opts: ["ثاني أكسيد الكربون", "الأكسجين", "النيتروجين", "الهيدروجين"], correct: 1 },
    { q: "كم عدد حالات المادة الأساسية؟", opts: ["اثنتان", "ثلاث", "أربع", "خمس"], correct: 1 },
    { q: "ما الكوكب الأقرب للشمس؟", opts: ["الأرض", "الزهرة", "عطارد", "المريخ"], correct: 2 },
    { q: "أي الأعضاء مسؤول عن ضخ الدم؟", opts: ["الرئة", "القلب", "الكبد", "المعدة"], correct: 1 },
    { q: "ما مصدر الطاقة الأساسي لكوكب الأرض؟", opts: ["القمر", "الشمس", "الرياح", "الماء"], correct: 1 },
    { q: "ماذا نسمي عملية تحويل النبات لضوء الشمس إلى غذاء؟", opts: ["التبخر", "البناء الضوئي", "التنفس", "الترسيب"], correct: 1 },
  ],
  "التربية الإسلامية": [
    { q: "كم عدد أركان الإسلام؟", opts: ["ثلاثة", "أربعة", "خمسة", "ستة"], correct: 2 },
    { q: "كم عدد الصلوات المفروضة يوميًا؟", opts: ["ثلاث", "أربع", "خمس", "ست"], correct: 2 },
    { q: "في أي شهر يصوم المسلمون؟", opts: ["شعبان", "رمضان", "شوال", "رجب"], correct: 1 },
    { q: "ما هي أول سورة نزلت في القرآن الكريم؟", opts: ["الفاتحة", "العلق", "البقرة", "الإخلاص"], correct: 1 },
    { q: "كم عدد أركان الإيمان؟", opts: ["أربعة", "خمسة", "ستة", "سبعة"], correct: 2 },
    { q: "ما اسم المدينة التي هاجر إليها النبي ﷺ؟", opts: ["مكة", "المدينة المنورة", "الطائف", "جدة"], correct: 1 },
  ],
  "الدراسات الاجتماعية": [
    { q: "ما عاصمة المملكة العربية السعودية؟", opts: ["جدة", "الرياض", "الدمام", "أبها"], correct: 1 },
    { q: "كم عدد قارات العالم؟", opts: ["خمس", "ستة", "سبع", "ثمان"], correct: 2 },
    { q: "أي البحار يحد المملكة من الغرب؟", opts: ["الخليج العربي", "البحر الأحمر", "بحر العرب", "المحيط الهندي"], correct: 1 },
    { q: "ما أطول نهر في العالم؟", opts: ["الأمازون", "النيل", "دجلة", "الفرات"], correct: 1 },
    { q: "كم عدد مناطق المملكة العربية السعودية؟", opts: ["١٠", "١١", "١٢", "١٣"], correct: 3 },
    { q: "أي مما يلي من مصادر الدخل الرئيسية بالمملكة؟", opts: ["الزراعة فقط", "النفط", "الصيد", "السياحة فقط"], correct: 1 },
  ],
};

// خطوة إعداد لعبة الأسئلة: اختيار مادة من بنك جاهز، أو كتابة أسئلة خاصة —
// بشكل اختياري بالكامل، حسب رغبة المعلم.
function QuizSetupModal({ onStart, onClose, onBack }) {
  const [mode, setMode] = useState("bank"); // bank | custom
  const [subject, setSubject] = useState(Object.keys(SUBJECT_QUESTION_BANKS)[0]);
  const [customQuestions, setCustomQuestions] = useState([{ id: uid(), text: "", options: ["", ""], correct: 0 }]);

  const addCustomQuestion = () => setCustomQuestions((qs) => [...qs, { id: uid(), text: "", options: ["", ""], correct: 0 }]);
  const removeCustomQuestion = (id) => setCustomQuestions((qs) => qs.filter((q) => q.id !== id));
  const updateQuestion = (id, patch) => setCustomQuestions((qs) => qs.map((q) => (q.id === id ? { ...q, ...patch } : q)));
  const updateOption = (id, i, val) => setCustomQuestions((qs) => qs.map((q) => (q.id === id ? { ...q, options: q.options.map((o, oi) => (oi === i ? val : o)) } : q)));
  const addOption = (id) => setCustomQuestions((qs) => qs.map((q) => (q.id === id ? { ...q, options: [...q.options, ""] } : q)));
  const removeOption = (id, i) => setCustomQuestions((qs) => qs.map((q) => (q.id === id ? { ...q, options: q.options.filter((_, oi) => oi !== i), correct: q.correct === i ? 0 : q.correct } : q)));

  const customValid = customQuestions.length > 0 && customQuestions.every((q) => q.text.trim() && q.options.length >= 2 && q.options.every((o) => o.trim()));

  const start = () => {
    if (mode === "bank") {
      const questions = SUBJECT_QUESTION_BANKS[subject];
      onStart({ name: `لعبة الأسئلة السريعة — ${subject}`, html: buildQuizGameHtml(questions) });
    } else {
      const questions = customQuestions.map((q) => ({ q: q.text.trim(), opts: q.options.map((o) => o.trim()), correct: q.correct }));
      onStart({ name: "لعبة الأسئلة السريعة — أسئلتي الخاصة", html: buildQuizGameHtml(questions) });
    }
  };

  return (
    <Modal title="إعداد لعبة الأسئلة السريعة" onClose={onClose} onBack={onBack} accent="magic" wide>
      <div className="flex gap-2 mb-4">
        <button onClick={() => setMode("bank")} className="flex-1 text-sm font-semibold py-2.5 rounded-lg" style={{ background: mode === "bank" ? INK : "transparent", color: mode === "bank" ? "#fff" : MUTED, border: `1px solid ${mode === "bank" ? INK : LINE}` }}>بنك أسئلة جاهز حسب المادة</button>
        <button onClick={() => setMode("custom")} className="flex-1 text-sm font-semibold py-2.5 rounded-lg" style={{ background: mode === "custom" ? INK : "transparent", color: mode === "custom" ? "#fff" : MUTED, border: `1px solid ${mode === "custom" ? INK : LINE}` }}>أكتب أسئلتي الخاصة</button>
      </div>

      {mode === "bank" ? (
        <Field label="اختر المادة">
          <select value={subject} onChange={(e) => setSubject(e.target.value)} style={inputStyle}>
            {Object.keys(SUBJECT_QUESTION_BANKS).map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <p className="text-xs mt-2" style={{ color: MUTED }}>
            أسئلة ثقافة عامة بأسلوب هذي المادة (مو منسوخة من كتاب رسمي معيّن) — نقطة بداية سريعة تقدر تستبدلها بأسئلتك الخاصة بأي وقت.
          </p>
        </Field>
      ) : (
        <div className="space-y-3 mb-3 max-h-96 overflow-y-auto">
          {customQuestions.map((q, qi) => (
            <div key={q.id} className="p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
              <div className="flex items-center gap-2 mb-2">
                <span className="text-xs font-bold px-2 py-1 rounded-full shrink-0" style={{ background: "#F3F1E9", color: MUTED }}>س{qi + 1}</span>
                <input style={{ ...inputStyle, flex: 1 }} value={q.text} onChange={(e) => updateQuestion(q.id, { text: e.target.value })} placeholder="نص السؤال" />
                {customQuestions.length > 1 && <button onClick={() => removeCustomQuestion(q.id)} className="p-1.5 rounded hover:bg-black/5 shrink-0"><Trash2 size={14} color="#C0392B" /></button>}
              </div>
              <div className="space-y-1.5 mr-8">
                {q.options.map((o, oi) => (
                  <div key={oi} className="flex items-center gap-2">
                    <input type="radio" name={`correct-${q.id}`} checked={q.correct === oi} onChange={() => updateQuestion(q.id, { correct: oi })} />
                    <input style={{ ...inputStyle, flex: 1, padding: "6px 10px" }} value={o} onChange={(e) => updateOption(q.id, oi, e.target.value)} placeholder="خيار" />
                    {q.options.length > 2 && <button onClick={() => removeOption(q.id, oi)} className="p-1 rounded hover:bg-black/5"><X size={12} color={MUTED} /></button>}
                  </div>
                ))}
                <button onClick={() => addOption(q.id)} className="text-xs font-semibold flex items-center gap-1" style={{ color: "#26423B" }}><Plus size={12} /> إضافة خيار</button>
              </div>
            </div>
          ))}
          <button onClick={addCustomQuestion} className="text-sm font-semibold flex items-center gap-1" style={{ color: "#26423B" }}><Plus size={15} /> إضافة سؤال</button>
        </div>
      )}

      <button
        disabled={mode === "custom" && !customValid}
        onClick={start}
        className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
      >
        ابدأ اللعبة
      </button>
    </Modal>
  );
}

// ---------- ٤ ألعاب إضافية جاهزة (تطبّق نفس بروتوكول postMessage) ----------

const BUILTIN_GAME_TRUEFALSE_HTML = `<!DOCTYPE html>
<html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { font-family: Tahoma, Arial, sans-serif; background: #FAF3EE; margin: 0; padding: 20px; display: flex; align-items: center; justify-content: center; min-height: 100vh; }
  .card { background: #fff; border-radius: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); padding: 28px; width: 100%; max-width: 460px; text-align: center; }
  .badge { display: inline-block; background: #C97A2B; color: #fff; font-size: 12px; font-weight: bold; padding: 5px 14px; border-radius: 999px; margin-bottom: 16px; }
  h1 { font-size: 20px; color: #232622; margin: 0 0 24px; line-height: 1.6; }
  .row { display: flex; gap: 12px; }
  button.opt { flex: 1; font-family: inherit; font-size: 16px; font-weight: bold; padding: 18px; border-radius: 14px; border: 2px solid #E4DFD2; background: #fff; cursor: pointer; }
  button.opt.correct { background: #E3F1EC; border-color: #0F9D58; color: #0F9D58; }
  button.opt.wrong { background: #FBEAE7; border-color: #C0392B; color: #C0392B; }
  .feedback { margin-top: 16px; font-size: 14px; font-weight: bold; min-height: 20px; }
  .score { position: absolute; top: 16px; left: 16px; font-size: 12px; color: #7A7768; }
</style></head>
<body>
  <div class="score" id="score">النقاط: 0</div>
  <div class="card">
    <span class="badge" id="qnum">عبارة 1</span>
    <h1 id="question"></h1>
    <div class="row">
      <button class="opt" id="btnTrue">✓ صح</button>
      <button class="opt" id="btnFalse">✗ خطأ</button>
    </div>
    <p class="feedback" id="feedback"></p>
  </div>
<script>
  const STATEMENTS = [
    { q: "الشمس أكبر من الأرض بكثير.", correct: true },
    { q: "عدد أيام السنة الميلادية 300 يوم.", correct: false },
    { q: "الماء يتجمد عند درجة الصفر المئوي.", correct: true },
    { q: "الفيل من الحيوانات الأليفة الصغيرة.", correct: false },
    { q: "الرياض هي عاصمة المملكة العربية السعودية.", correct: true },
    { q: "عدد أضلاع المربع خمسة.", correct: false },
    { q: "النحل ينتج العسل.", correct: true },
    { q: "الليل أطول من النهار طوال أيام السنة في كل مكان.", correct: false },
  ];
  let idx = Math.floor(Math.random() * STATEMENTS.length);
  let score = 0, qCount = 0, locked = false;
  function render() {
    locked = false; qCount++;
    document.getElementById("qnum").textContent = "عبارة " + qCount;
    document.getElementById("question").textContent = STATEMENTS[idx].q;
    document.getElementById("feedback").textContent = "";
    document.getElementById("btnTrue").className = "opt";
    document.getElementById("btnFalse").className = "opt";
  }
  function choose(answer) {
    if (locked) return;
    locked = true;
    const isCorrect = answer === STATEMENTS[idx].correct;
    document.getElementById(answer ? "btnTrue" : "btnFalse").className = "opt " + (isCorrect ? "correct" : "wrong");
    document.getElementById("feedback").textContent = isCorrect ? "✓ إجابة صحيحة! 🎉" : "✗ إجابة خاطئة";
    document.getElementById("feedback").style.color = isCorrect ? "#0F9D58" : "#C0392B";
    if (isCorrect) { score++; document.getElementById("score").textContent = "النقاط: " + score; }
    window.parent.postMessage({ type: "fasooli:answer", correct: isCorrect }, "*");
    setTimeout(() => { idx = Math.floor(Math.random() * STATEMENTS.length); render(); }, 1200);
  }
  document.getElementById("btnTrue").onclick = () => choose(true);
  document.getElementById("btnFalse").onclick = () => choose(false);
  render();
</script>
</body></html>`;

const BUILTIN_GAME_FILLBLANK_HTML = `<!DOCTYPE html>
<html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { font-family: Tahoma, Arial, sans-serif; background: #EEF3FA; margin: 0; padding: 20px; display: flex; align-items: center; justify-content: center; min-height: 100vh; }
  .card { background: #fff; border-radius: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); padding: 28px; width: 100%; max-width: 460px; text-align: center; }
  .badge { display: inline-block; background: #2E7DA6; color: #fff; font-size: 12px; font-weight: bold; padding: 5px 14px; border-radius: 999px; margin-bottom: 16px; }
  h1 { font-size: 19px; color: #232622; margin: 0 0 24px; line-height: 1.7; }
  .blank { display: inline-block; min-width: 60px; border-bottom: 3px solid #2E7DA6; color: #2E7DA6; font-weight: bold; }
  .options { display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; }
  button.opt { font-family: inherit; font-size: 15px; padding: 10px 18px; border-radius: 999px; border: 2px solid #E4DFD2; background: #fff; cursor: pointer; }
  button.opt.correct { background: #E3F1EC; border-color: #0F9D58; color: #0F9D58; font-weight: bold; }
  button.opt.wrong { background: #FBEAE7; border-color: #C0392B; color: #C0392B; }
  .feedback { margin-top: 16px; font-size: 14px; font-weight: bold; min-height: 20px; }
  .score { position: absolute; top: 16px; left: 16px; font-size: 12px; color: #7A7768; }
</style></head>
<body>
  <div class="score" id="score">النقاط: 0</div>
  <div class="card">
    <span class="badge" id="qnum">جملة 1</span>
    <h1 id="question"></h1>
    <div class="options" id="options"></div>
    <p class="feedback" id="feedback"></p>
  </div>
<script>
  const ITEMS = [
    { before: "تشرق الشمس من جهة ", after: ".", opts: ["الشرق", "الغرب", "الشمال"], correct: 0 },
    { before: "يتكوّن الأسبوع من ", after: " أيام.", opts: ["خمسة", "ستة", "سبعة"], correct: 2 },
    { before: "كلمة يكتبُ فعل ", after: ".", opts: ["ماضٍ", "مضارع", "أمر"], correct: 1 },
    { before: "أكبر كوكب في مجموعتنا الشمسية هو ", after: ".", opts: ["الأرض", "المشتري", "زحل"], correct: 1 },
    { before: "الماء يغلي عند درجة حرارة ", after: " مئوية.", opts: ["50", "80", "100"], correct: 2 },
    { before: "عاصمة المملكة العربية السعودية هي ", after: ".", opts: ["جدة", "الرياض", "الدمام"], correct: 1 },
  ];
  let idx = Math.floor(Math.random() * ITEMS.length);
  let score = 0, qCount = 0, locked = false;
  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  }
  function render() {
    locked = false; qCount++;
    const item = ITEMS[idx];
    document.getElementById("qnum").textContent = "جملة " + qCount;
    document.getElementById("question").innerHTML = item.before + '<span class="blank">___</span>' + item.after;
    document.getElementById("feedback").textContent = "";
    const optsDiv = document.getElementById("options");
    optsDiv.innerHTML = "";
    const order = shuffle(item.opts.map((t, i) => ({ t, correct: i === item.correct })));
    order.forEach((o) => {
      const btn = document.createElement("button");
      btn.className = "opt"; btn.textContent = o.t;
      btn.onclick = () => choose(btn, o.correct);
      optsDiv.appendChild(btn);
    });
  }
  function choose(btn, isCorrect) {
    if (locked) return;
    locked = true;
    btn.classList.add(isCorrect ? "correct" : "wrong");
    document.getElementById("feedback").textContent = isCorrect ? "✓ أحسنت! 🎉" : "✗ إجابة خاطئة";
    document.getElementById("feedback").style.color = isCorrect ? "#0F9D58" : "#C0392B";
    if (isCorrect) { score++; document.getElementById("score").textContent = "النقاط: " + score; }
    window.parent.postMessage({ type: "fasooli:answer", correct: isCorrect }, "*");
    setTimeout(() => { idx = Math.floor(Math.random() * ITEMS.length); render(); }, 1400);
  }
  render();
</script>
</body></html>`;

const BUILTIN_GAME_MEMORY_HTML = `<!DOCTYPE html>
<html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { font-family: Tahoma, Arial, sans-serif; background: #F3EEFA; margin: 0; padding: 20px; display: flex; align-items: center; justify-content: center; min-height: 100vh; }
  .card { background: #fff; border-radius: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); padding: 24px; width: 100%; max-width: 440px; text-align: center; }
  .badge { display: inline-block; background: #7A4E9E; color: #fff; font-size: 12px; font-weight: bold; padding: 5px 14px; border-radius: 999px; margin-bottom: 16px; }
  .grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; }
  .tile { aspect-ratio: 1; border-radius: 10px; background: #7A4E9E; display: flex; align-items: center; justify-content: center; font-size: 13px; font-weight: bold; color: #fff; cursor: pointer; user-select: none; padding: 4px; text-align: center; }
  .tile.flipped { background: #fff; border: 2px solid #7A4E9E; color: #232622; }
  .tile.matched { background: #E3F1EC; border: 2px solid #0F9D58; color: #0F9D58; }
  .feedback { margin-top: 16px; font-size: 14px; font-weight: bold; min-height: 20px; }
  .score { position: absolute; top: 16px; left: 16px; font-size: 12px; color: #7A7768; }
</style></head>
<body>
  <div class="score" id="score">أزواج: 0</div>
  <div class="card">
    <span class="badge">لعبة الذاكرة — طابق الأزواج</span>
    <div class="grid" id="grid"></div>
    <p class="feedback" id="feedback"></p>
  </div>
<script>
  const PAIRS = [
    ["واحد", "1"], ["اثنان", "2"], ["ثلاثة", "3"], ["أربعة", "4"],
    ["خمسة", "5"], ["ستة", "6"], ["سبعة", "7"], ["ثمانية", "8"],
  ];
  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  }
  function buildDeck() {
    const chosen = shuffle(PAIRS).slice(0, 6);
    let cards = [];
    chosen.forEach((p, i) => { cards.push({ id: i + "a", pairId: i, text: p[0] }); cards.push({ id: i + "b", pairId: i, text: p[1] }); });
    return shuffle(cards);
  }
  let deck = buildDeck();
  let flipped = [];
  let matchedCount = 0;
  let locked = false;

  function render() {
    const grid = document.getElementById("grid");
    grid.innerHTML = "";
    deck.forEach((card) => {
      const el = document.createElement("div");
      el.className = "tile";
      el.textContent = "؟";
      el.onclick = () => flip(card, el);
      el.dataset.id = card.id;
      grid.appendChild(el);
    });
  }
  function flip(card, el) {
    if (locked || el.classList.contains("matched") || el.classList.contains("flipped")) return;
    el.classList.add("flipped");
    el.textContent = card.text;
    flipped.push({ card, el });
    if (flipped.length === 2) {
      locked = true;
      const [a, b] = flipped;
      if (a.card.pairId === b.card.pairId) {
        setTimeout(() => {
          a.el.classList.add("matched"); b.el.classList.add("matched");
          matchedCount++;
          document.getElementById("score").textContent = "أزواج: " + matchedCount;
          document.getElementById("feedback").textContent = "✓ تطابق صحيح! 🎉";
          document.getElementById("feedback").style.color = "#0F9D58";
          window.parent.postMessage({ type: "fasooli:answer", correct: true }, "*");
          flipped = []; locked = false;
          if (matchedCount === 6) { document.getElementById("feedback").textContent = "🏆 أكملت كل الأزواج! لعبة جديدة..."; setTimeout(() => { deck = buildDeck(); matchedCount = 0; render(); document.getElementById("score").textContent = "أزواج: 0"; }, 1600); }
        }, 400);
      } else {
        document.getElementById("feedback").textContent = "✗ غير متطابق";
        document.getElementById("feedback").style.color = "#C0392B";
        window.parent.postMessage({ type: "fasooli:answer", correct: false }, "*");
        setTimeout(() => {
          a.el.classList.remove("flipped"); a.el.textContent = "؟";
          b.el.classList.remove("flipped"); b.el.textContent = "؟";
          flipped = []; locked = false;
        }, 900);
      }
    }
  }
  render();
</script>
</body></html>`;

const BUILTIN_GAME_SPEEDROUND_HTML = `<!DOCTYPE html>
<html dir="rtl" lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  * { box-sizing: border-box; }
  body { font-family: Tahoma, Arial, sans-serif; background: #FDF1F1; margin: 0; padding: 20px; display: flex; align-items: center; justify-content: center; min-height: 100vh; }
  .card { background: #fff; border-radius: 20px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); padding: 28px; width: 100%; max-width: 460px; text-align: center; }
  .badge { display: inline-block; background: #C0392B; color: #fff; font-size: 12px; font-weight: bold; padding: 5px 14px; border-radius: 999px; margin-bottom: 12px; }
  .timerbar { width: 100%; height: 8px; background: #F3E6E6; border-radius: 999px; overflow: hidden; margin-bottom: 18px; }
  .timerfill { height: 100%; background: #C0392B; width: 100%; transition: width 0.1s linear; }
  h1 { font-size: 19px; color: #232622; margin: 0 0 20px; line-height: 1.5; }
  .options { display: flex; flex-direction: column; gap: 10px; }
  button.opt { font-family: inherit; font-size: 15px; padding: 14px; border-radius: 12px; border: 2px solid #E4DFD2; background: #fff; cursor: pointer; }
  button.opt.correct { background: #E3F1EC; border-color: #0F9D58; color: #0F9D58; font-weight: bold; }
  button.opt.wrong { background: #FBEAE7; border-color: #C0392B; color: #C0392B; }
  .feedback { margin-top: 16px; font-size: 14px; font-weight: bold; min-height: 20px; }
  .score { position: absolute; top: 16px; left: 16px; font-size: 12px; color: #7A7768; }
</style></head>
<body>
  <div class="score" id="score">النقاط: 0</div>
  <div class="card">
    <span class="badge" id="qnum">تحدي 1 — ٨ ثواني</span>
    <div class="timerbar"><div class="timerfill" id="timerfill"></div></div>
    <h1 id="question"></h1>
    <div class="options" id="options"></div>
    <p class="feedback" id="feedback"></p>
  </div>
<script>
  const QUESTIONS = [
    { q: "ما ناتج 6 + 7؟", opts: ["12", "13", "14", "11"], correct: 1 },
    { q: "عاصمة مصر؟", opts: ["الإسكندرية", "القاهرة", "الأقصر", "أسوان"], correct: 1 },
    { q: "كم شهرًا بالسنة؟", opts: ["10", "11", "12", "13"], correct: 2 },
    { q: "لون السماء الصافية؟", opts: ["أخضر", "أزرق", "أحمر", "أصفر"], correct: 1 },
    { q: "ما ناتج 9 × 3؟", opts: ["24", "27", "30", "21"], correct: 1 },
    { q: "كم عدد أيام الأسبوع؟", opts: ["6", "7", "8", "5"], correct: 1 },
  ];
  let idx = Math.floor(Math.random() * QUESTIONS.length);
  let score = 0, qCount = 0, locked = false, timerInterval = null;
  const DURATION = 8000;

  function shuffle(arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  }
  function startTimer() {
    const start = Date.now();
    clearInterval(timerInterval);
    timerInterval = setInterval(() => {
      const elapsed = Date.now() - start;
      const pct = Math.max(0, 100 - (elapsed / DURATION) * 100);
      document.getElementById("timerfill").style.width = pct + "%";
      if (elapsed >= DURATION) { clearInterval(timerInterval); timeUp(); }
    }, 80);
  }
  function timeUp() {
    if (locked) return;
    locked = true;
    document.getElementById("feedback").textContent = "⏱️ انتهى الوقت!";
    document.getElementById("feedback").style.color = "#C0392B";
    window.parent.postMessage({ type: "fasooli:answer", correct: false }, "*");
    setTimeout(next, 1200);
  }
  function render() {
    locked = false; qCount++;
    const item = QUESTIONS[idx];
    document.getElementById("qnum").textContent = "تحدي " + qCount + " — ٨ ثواني";
    document.getElementById("question").textContent = item.q;
    document.getElementById("feedback").textContent = "";
    document.getElementById("timerfill").style.width = "100%";
    const optsDiv = document.getElementById("options");
    optsDiv.innerHTML = "";
    const order = shuffle(item.opts.map((t, i) => ({ t, correct: i === item.correct })));
    order.forEach((o) => {
      const btn = document.createElement("button");
      btn.className = "opt"; btn.textContent = o.t;
      btn.onclick = () => choose(btn, o.correct);
      optsDiv.appendChild(btn);
    });
    startTimer();
  }
  function choose(btn, isCorrect) {
    if (locked) return;
    locked = true;
    clearInterval(timerInterval);
    btn.classList.add(isCorrect ? "correct" : "wrong");
    document.getElementById("feedback").textContent = isCorrect ? "✓ سريع وصحيح! 🎉" : "✗ إجابة خاطئة";
    document.getElementById("feedback").style.color = isCorrect ? "#0F9D58" : "#C0392B";
    if (isCorrect) { score++; document.getElementById("score").textContent = "النقاط: " + score; }
    window.parent.postMessage({ type: "fasooli:answer", correct: isCorrect }, "*");
    setTimeout(next, 1200);
  }
  function next() { idx = Math.floor(Math.random() * QUESTIONS.length); render(); }
  render();
</script>
</body></html>`;

const BUILTIN_GAMES = [
  {
    id: "builtin-quiz",
    name: "لعبة الأسئلة السريعة",
    description: "أسئلة اختيار من متعدد — بنك جاهز حسب المادة، أو اكتب أسئلتك الخاصة.",
    needsSetup: true,
  },
  {
    id: "builtin-truefalse",
    name: "لعبة صح أو خطأ",
    description: "عبارات سريعة يحدد الطالب هل هي صحيحة أو خاطئة.",
    html: BUILTIN_GAME_TRUEFALSE_HTML,
  },
  {
    id: "builtin-fillblank",
    name: "أكمل الجملة",
    description: "جملة ناقصة واختيار الكلمة الصحيحة لإكمالها.",
    html: BUILTIN_GAME_FILLBLANK_HTML,
  },
  {
    id: "builtin-memory",
    name: "لعبة الذاكرة",
    description: "طابق الأزواج المتشابهة (أرقام وكلماتها) — كل تطابق صحيح نقطة.",
    html: BUILTIN_GAME_MEMORY_HTML,
  },
  {
    id: "builtin-speedround",
    name: "التحدي السريع",
    description: "أسئلة بعدّاد زمني ٨ ثوانٍ لكل سؤال — إثارة أكبر بالسرعة.",
    html: BUILTIN_GAME_SPEEDROUND_HTML,
  },
];

// انفجار علامة ✓ خضراء قصير فوق صف الطالب اللي حصل على نقطة لحظتها —
// نفس أسلوب التأكيد البصري المطلوب.
// تأثير احتفالي عند إجابة صحيحة: علامة ✓ كبيرة + قصاصات ملوّنة متطايرة —
// أو وجه حزين بسيط لو كانت الإجابة خاطئة (بدون نقطة).
function GamePointBurst({ kind = "correct" }) {
  const confettiColors = ["#4ADE80", "#FBBF24", "#60A5FA", "#F472B6", "#A78BFA"];
  return (
    <div style={{ position: "absolute", inset: 0, zIndex: 6, pointerEvents: "none", overflow: "visible" }}>
      {kind === "correct" ? (
        <>
          {Array.from({ length: 10 }).map((_, i) => {
            const angle = (i / 10) * 2 * Math.PI;
            const dist = 46 + Math.random() * 18;
            const tx = Math.cos(angle) * dist;
            const ty = Math.sin(angle) * dist;
            return (
              <span
                key={i}
                className="confetti-bit"
                style={{
                  width: 7, height: 7, borderRadius: i % 2 ? "50%" : 2,
                  background: confettiColors[i % confettiColors.length],
                  "--tx": `${tx}px`, "--ty": `${ty}px`, "--rot": `${Math.round(Math.random() * 360)}deg`,
                  animationDelay: `${i * 15}ms`,
                }}
              />
            );
          })}
          <div className="big-check-in" style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div className="w-14 h-14 rounded-full flex items-center justify-center shadow-lg" style={{ background: "#22C55E", boxShadow: "0 4px 16px rgba(34,197,94,0.5)" }}>
              <Check size={30} color="#fff" strokeWidth={4} />
            </div>
          </div>
        </>
      ) : (
        <div className="sad-wobble-in" style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div className="w-14 h-14 rounded-full flex items-center justify-center shadow-lg text-2xl" style={{ background: "#FBEAE7" }}>
            😢
          </div>
        </div>
      )}
    </div>
  );
}

// نافذة تشغيل اللعبة فعليًا بعد ربطها بفصل — يتحول "الطالب النشط" تلقائيًا
// بشكل عشوائي بعد كل سؤال (إجابة صح أو خطأ)، وأي إجابة صحيحة تضيف نقطة
// مباشرة بعمود "المشاركة" لذاك الطالب تلقائيًا (يُنشأ العمود تلقائيًا أول
// مرة لو ما كان موجودًا بالفصل).
function GamePlayerModal({ cls, updateClass, game, onClose, onBack }) {
  const [displayRows] = useState(() => shuffleArr(cls.rows));
  const [activeRowId, setActiveRowId] = useState(() => displayRows[0]?.id || null);
  const [flash, setFlash] = useState(null); // { rowId, kind }

  const ensureParticipationColumn = () => {
    const existing = cls.columns.find((c) => c.name === "المشاركة" && c.type === "counter");
    if (existing) return existing.id;
    const newId = uid();
    updateClass((c) => {
      if (c.columns.some((cc) => cc.name === "المشاركة" && cc.type === "counter")) return c;
      return { ...c, columns: [...c.columns, { id: newId, name: "المشاركة", type: "counter", options: [], color: "#0F9D58", autoRenew: false, pinned: false }] };
    });
    return newId;
  };

  const rotateToRandomStudent = (currentId) => {
    const others = displayRows.filter((r) => r.id !== currentId);
    const pool = others.length > 0 ? others : displayRows;
    if (pool.length === 0) return;
    setActiveRowId(pool[Math.floor(Math.random() * pool.length)].id);
  };

  useEffect(() => {
    const handler = (e) => {
      if (!e.data || e.data.type !== "fasooli:answer") return;
      if (!activeRowId) return;
      const rowId = activeRowId;
      if (e.data.correct) {
        const colId = ensureParticipationColumn();
        updateClass((c) => {
          const col = c.columns.find((cc) => cc.id === colId) || { id: colId, name: "المشاركة", color: "#0F9D58" };
          const key = `${rowId}:${colId}`;
          const current = Number(c.cells[key]) || 0;
          const next = current + 1;
          const meta = nowMeta();
          const entry = { id: uid(), colId, colName: col.name, colColor: col.color, value: String(next), ...meta };
          return {
            ...c,
            cells: { ...c.cells, [key]: String(next) },
            reports: { ...(c.reports || {}), [rowId]: [...(c.reports?.[rowId] || []), entry] },
          };
        });
        setFlash({ rowId, kind: "correct" });
      } else {
        setFlash({ rowId, kind: "wrong" });
      }
      setTimeout(() => {
        setFlash(null);
        rotateToRandomStudent(rowId);
      }, 1100);
    };
    window.addEventListener("message", handler);
    return () => window.removeEventListener("message", handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeRowId, cls.columns]);

  const participationCol = cls.columns.find((c) => c.name === "المشاركة" && c.type === "counter");

  // هذي الشاشة تحديدًا مبنية كطبقة ثابتة تملأ الشاشة كاملة بترتيب عمودي
  // (رأس مضغوط + شريط الطلاب + اللعبة تاخذ الباقي) بدل الاعتماد على تمرير
  // الصفحة — لأن اللمس داخل الـ iframe لا ينقل حركة التمرير للصفحة اللي
  // ورا، وكان يخلي شريط الطلاب "يختفي" على الجوال بدون أي طريقة ترجعه.
  return (
    <div className="fixed inset-0 flex flex-col" style={{ background: PAPER, zIndex: 60, height: "100dvh" }}>
      <div className="flex items-center justify-between px-4 py-3 shrink-0" style={{ background: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)" }}>
        <div className="flex items-center gap-2 min-w-0">
          {onBack && (
            <button onClick={onBack} className="p-1 rounded-lg shrink-0 hover:bg-white/10"><ArrowRight size={18} color="#fff" /></button>
          )}
          <h3 className="font-bold text-sm truncate" style={{ color: "#fff" }}>{game.name} — {cls.subject}</h3>
        </div>
        <button onClick={onClose} className="p-1 rounded-lg shrink-0 hover:bg-white/10"><X size={20} color="#fff" /></button>
      </div>

      <div className="shrink-0 px-3 py-2" style={{ background: "#fff", borderBottom: `1px solid ${LINE}` }}>
        <div className="flex items-center justify-between mb-1.5">
          <p className="text-xs font-bold" style={{ color: INK }}>الطالب النشط (يتغيّر عشوائيًا كل سؤال)</p>
          <p className="text-[11px]" style={{ color: MUTED }}>النقاط بعمود "المشاركة"</p>
        </div>
        <div className="flex gap-2 overflow-x-auto pb-1" style={{ scrollSnapType: "x proximity" }}>
          {displayRows.map((row) => {
            const isActive = row.id === activeRowId;
            const pts = participationCol ? (cls.cells[`${row.id}:${participationCol.id}`] || 0) : 0;
            return (
              <button
                key={row.id}
                onClick={() => setActiveRowId(row.id)}
                className="relative flex flex-col items-center gap-1 shrink-0 p-2 rounded-xl transition-all"
                style={{ border: `2px solid ${isActive ? "#0F9D58" : LINE}`, background: isActive ? "#E3F1EC" : "#fff", width: 72, scrollSnapAlign: "start" }}
              >
                {flash?.rowId === row.id && <GamePointBurst kind={flash.kind} />}
                <span className="w-8 h-8 rounded-full flex items-center justify-center font-bold text-sm text-white shrink-0" style={{ background: row.color }}>
                  {(row.name || "؟").trim().charAt(0)}
                </span>
                <span className="text-[10px] font-semibold text-center leading-tight truncate w-full" style={{ color: INK }}>{row.name}</span>
                <span className="text-[9px] font-bold px-1.5 py-0.5 rounded-full shrink-0" style={{ background: "#26423B", color: "#fff" }}>{pts}</span>
              </button>
            );
          })}
        </div>
      </div>

      <div className="flex-1 min-h-0">
        <iframe
          title="لعبة"
          srcDoc={game.html}
          sandbox="allow-scripts"
          style={{ width: "100%", height: "100%", border: "none", display: "block" }}
        />
      </div>
    </div>
  );
}

// خطوة اختيار الفصل اللي راح تُربَط فيه اللعبة قبل بدء اللعب.
function GameClassPickerModal({ classes, onSelect, onClose, onBack }) {
  const activeClasses = classes.filter((c) => !c.archived);
  return (
    <div className="fixed inset-0 flex flex-col" style={{ background: "rgba(25,28,25,0.55)", zIndex: 60 }} onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="w-full mx-auto mt-auto mb-auto sm:mt-16" style={{ maxWidth: 460 }}>
        <div className="rounded-2xl overflow-hidden" style={{ background: PAPER, border: `1px solid ${LINE}`, maxHeight: "80dvh", display: "flex", flexDirection: "column" }}>
          <div className="flex items-center justify-between px-5 py-4 shrink-0" style={{ background: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)" }}>
            <div className="flex items-center gap-2">
              {onBack && <button onClick={onBack} className="p-1 rounded-lg hover:bg-white/10"><ArrowRight size={18} color="#fff" /></button>}
              <h3 className="font-bold text-base" style={{ color: "#fff" }}>اربط اللعبة بفصل</h3>
            </div>
            <button onClick={onClose} className="p-1 rounded-lg hover:bg-white/10"><X size={20} color="#fff" /></button>
          </div>
          <div className="p-4 overflow-y-auto">
            {activeClasses.length === 0 ? (
              <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد فصول بعد.</p>
            ) : (
              <div className="space-y-2">
                {activeClasses.map((c) => (
                  <button
                    key={c.id}
                    onClick={() => onSelect(c.id)}
                    className="w-full flex items-center gap-3 p-3 rounded-xl text-right hover:bg-black/5"
                    style={{ border: `1px solid ${LINE}`, background: "#fff" }}
                  >
                    <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: c.color }} />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-semibold truncate" style={{ color: INK }}>{c.subject}</p>
                      <p className="text-xs" style={{ color: MUTED }}>{c.grade} • {c.rows.length} طالب</p>
                    </div>
                    <ChevronLeft size={16} color={MUTED} />
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// تبويب "الأدوات التفاعلية" بالصفحة الرئيسية: قائمة الألعاب (جاهزة من
// التطبيق، أو رفعتها بنفسك من المكتبة كملف HTML)، تختار لعبة ثم تربطها
// بفصل، وبعدها تبدأ اللعب مباشرة.
function GamesHub({ classes, library, updateClassById, bare = false }) {
  const [selectedGame, setSelectedGame] = useState(null);
  const [linkedClassId, setLinkedClassId] = useState(null);
  const [standalone, setStandalone] = useState(false);
  const [showQuizSetup, setShowQuizSetup] = useState(false);
  const [pendingLibraryGame, setPendingLibraryGame] = useState(null);

  const libraryGames = (library || []).filter((f) => (f.mimeType || "").includes("html") || (f.name || "").toLowerCase().endsWith(".html"));

  const resetAll = () => { setSelectedGame(null); setLinkedClassId(null); setStandalone(false); };

  if (showQuizSetup) {
    return (
      <QuizSetupModal
        onStart={(game) => { setSelectedGame(game); setShowQuizSetup(false); }}
        onClose={() => setShowQuizSetup(false)}
        onBack={() => setShowQuizSetup(false)}
      />
    );
  }

  if (pendingLibraryGame) {
    return (
      <Modal title={pendingLibraryGame.name} onClose={() => setPendingLibraryGame(null)} onBack={() => setPendingLibraryGame(null)} accent="magic">
        <p className="text-sm mb-4" style={{ color: MUTED }}>
          هذي لعبة رفعتها بنفسك — ما أقدر أضمن ربطها التلقائي بأسماء الطلاب لأني مو مطّلع على كودها الداخلي. اختر كيف تبي تشغّلها:
        </p>
        <div className="space-y-2">
          <button
            onClick={() => { setSelectedGame(pendingLibraryGame); setStandalone(false); setPendingLibraryGame(null); }}
            className="w-full flex items-center gap-3 p-3 rounded-xl text-right hover:bg-black/5"
            style={{ border: `1px solid ${LINE}`, background: "#fff" }}
          >
            <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ background: "#EAF3F0" }}><Users size={18} color="#26423B" /></div>
            <div className="flex-1">
              <p className="text-sm font-bold" style={{ color: INK }}>ربط بفصل (تسجيل نقاط تلقائي)</p>
              <p className="text-xs" style={{ color: MUTED }}>يعمل فقط لو اللعبة فعليًا تستخدم بروتوكول postMessage الموضّح بالأسفل.</p>
            </div>
          </button>
          <button
            onClick={() => { setSelectedGame(pendingLibraryGame); setStandalone(true); setPendingLibraryGame(null); }}
            className="w-full flex items-center gap-3 p-3 rounded-xl text-right hover:bg-black/5"
            style={{ border: `1px solid ${LINE}`, background: "#fff" }}
          >
            <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ background: "#F3F1E9" }}><Gamepad2 size={18} color={MUTED} /></div>
            <div className="flex-1">
              <p className="text-sm font-bold" style={{ color: INK }}>تشغيل مباشرة بدون ربط</p>
              <p className="text-xs" style={{ color: MUTED }}>يفتح اللعبة فقط للعرض/اللعب، بدون تسجيل أي نقاط بجدول الفصل.</p>
            </div>
          </button>
        </div>
      </Modal>
    );
  }

  if (selectedGame && standalone) {
    return (
      <div className="fixed inset-0 flex flex-col" style={{ background: PAPER, zIndex: 60, height: "100dvh" }}>
        <div className="flex items-center justify-between px-4 py-3 shrink-0" style={{ background: "linear-gradient(135deg, #7C5CE0, #4E6FE0, #2E9FD6)" }}>
          <h3 className="font-bold text-sm truncate" style={{ color: "#fff" }}>{selectedGame.name}</h3>
          <button onClick={resetAll} className="p-1 rounded-lg shrink-0 hover:bg-white/10"><X size={20} color="#fff" /></button>
        </div>
        <div className="flex-1 min-h-0">
          <iframe title="لعبة" srcDoc={selectedGame.html} sandbox="allow-scripts" style={{ width: "100%", height: "100%", border: "none", display: "block" }} />
        </div>
      </div>
    );
  }

  if (selectedGame && linkedClassId) {
    const cls = classes.find((c) => c.id === linkedClassId);
    if (!cls) { setLinkedClassId(null); return null; }
    return (
      <GamePlayerModal
        cls={cls}
        updateClass={(fn) => updateClassById(linkedClassId, fn)}
        game={selectedGame}
        onBack={() => setLinkedClassId(null)}
        onClose={resetAll}
      />
    );
  }

  if (selectedGame) {
    return (
      <GameClassPickerModal
        classes={classes}
        onSelect={(classId) => setLinkedClassId(classId)}
        onBack={() => setSelectedGame(null)}
        onClose={resetAll}
      />
    );
  }

  const content = (
    <>
      <p className="text-xs mb-4" style={{ color: MUTED }}>
        اختر لعبة، اربطها بفصل، وابدأ اللعب — أي إجابة صحيحة تضيف نقطة مباشرة لعمود "المشاركة" للطالب النشط.
      </p>
      <p className="text-xs font-bold mb-2" style={{ color: INK }}>ألعاب جاهزة (٥)</p>
      <div className="space-y-2 mb-5">
        {BUILTIN_GAMES.map((g) => (
          <button key={g.id} onClick={() => (g.needsSetup ? setShowQuizSetup(true) : setSelectedGame(g))} className="w-full flex items-center gap-3 p-3 rounded-xl text-right hover:bg-black/5" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
            <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ background: "#EAF3F0" }}>
              <Gamepad2 size={18} color="#26423B" />
            </div>
            <div className="flex-1">
              <p className="text-sm font-bold" style={{ color: INK }}>{g.name}</p>
              <p className="text-xs" style={{ color: MUTED }}>{g.description}</p>
            </div>
            <ChevronLeft size={16} color={MUTED} className="shrink-0" />
          </button>
        ))}
      </div>
      <p className="text-xs font-bold mb-2" style={{ color: INK }}>ألعابك المرفوعة (من المكتبة)</p>
      {libraryGames.length === 0 ? (
        <div className="rounded-2xl overflow-hidden" style={{ border: `1px solid ${LINE}` }}>
          <div className="p-4" style={{ background: "#EAF3F0" }}>
            <div className="flex items-center gap-2 mb-1.5">
              <Gamepad2 size={16} color="#26423B" />
              <p className="text-sm font-bold" style={{ color: "#26423B" }}>اصنع لعبتك الخاصة</p>
            </div>
            <p className="text-xs" style={{ color: "#26423B" }}>
              ارفع ملف HTML من تبويب "المكتبة" وسيظهر هنا تلقائيًا. تقدر تشغّلها مباشرة بدون ربط بفصل، أو تربطها لو تطبّق بروتوكول تسجيل النقاط بالأسفل.
            </p>
          </div>
          <div className="p-4" style={{ background: "#fff" }}>
            <p className="text-xs font-semibold mb-2" style={{ color: INK }}>لو تبي ربطها بنقاط تلقائية، نادِ هذا السطر عند كل إجابة صحيحة (اختياري)</p>
            <div className="relative rounded-xl p-3" style={{ background: "#1E2A26", direction: "ltr", textAlign: "left" }}>
              <code style={{ color: "#7DE3B3", fontSize: 11, fontFamily: "monospace", wordBreak: "break-all" }}>
                window.parent.postMessage({"{"} type: "fasooli:answer", correct: true {"}"}, "*");
              </code>
              <button
                onClick={async () => {
                  try { await navigator.clipboard.writeText(`window.parent.postMessage({ type: "fasooli:answer", correct: true }, "*");`); } catch (e) { /* ignore */ }
                }}
                className="mt-2 flex items-center gap-1.5 text-xs font-semibold px-2.5 py-1 rounded-lg"
                style={{ background: "rgba(255,255,255,0.1)", color: "#fff" }}
              >
                <ClipboardCopy size={12} /> نسخ السطر
              </button>
            </div>
          </div>
        </div>
      ) : (
        <div className="space-y-2">
          {libraryGames.map((f) => (
            <button
              key={f.id}
              onClick={() => {
                try {
                  const base64 = f.dataUrl.split(",")[1];
                  const html = decodeURIComponent(escape(atob(base64)));
                  setPendingLibraryGame({ id: f.id, name: f.name, html });
                } catch (e) {
                  alert("تعذّرت قراءة هذا الملف كصفحة HTML صحيحة.");
                }
              }}
              className="w-full flex items-center gap-3 p-3 rounded-xl text-right hover:bg-black/5"
              style={{ border: `1px solid ${LINE}`, background: "#fff" }}
            >
              <div className="w-10 h-10 rounded-lg flex items-center justify-center shrink-0" style={{ background: "#F3F1E9" }}>
                <FileText size={18} color="#26423B" />
              </div>
              <p className="text-sm font-bold flex-1" style={{ color: INK }}>{f.name}</p>
              <ChevronLeft size={16} color={MUTED} className="shrink-0" />
            </button>
          ))}
        </div>
      )}
    </>
  );

  if (bare) return content;
  return (
    <Modal title="الأدوات التفاعلية" onClose={() => {}} accent="magic" wide>
      {content}
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
          <span className="text-xs px-2.5 py-1 rounded-full font-bold" style={{ background: "#E3F0ED", color: "#26423B" }}>حاضر: {presentCount}</span>
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
                  style={{ background: status === "present" ? "#E3F0ED" : "#F5DEDB", color: status === "present" ? "#26423B" : "#C0392B" }}
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

function RemindersModal({ reminders, onAdd, onDelete, notifPermission, onRequestPermission, onClose }) {
  const [title, setTitle] = useState("");
  const [date, setDate] = useState(todayKey());
  const [time, setTime] = useState("08:00");

  const submit = () => {
    if (!title.trim()) return;
    onAdd({ id: uid(), title: title.trim(), date, time, notified: false });
    setTitle("");
  };

  const sorted = [...(reminders || [])].sort((a, b) => `${a.date}${a.time}`.localeCompare(`${b.date}${b.time}`));
  const now = new Date();

  return (
    <Modal title="تذكير" onClose={onClose}>
      {notifPermission !== "granted" && (
        <div className="flex items-start gap-2 p-3 rounded-xl mb-4" style={{ background: "#FCEFE2", border: "1px solid #F0D2CB" }}>
          <Bell size={15} color="#C97A2B" className="shrink-0 mt-0.5" />
          <div className="flex-1">
            <p className="text-xs mb-2" style={{ color: "#8A4A1E" }}>فعّل إشعارات المتصفح لتصلك التذكيرات كإشعار بالجهاز — يعمل ما دام الموقع مفتوحًا (حتى بتبويب بالخلفية)، ولا يعمل والمتصفح مغلق تمامًا.</p>
            <button onClick={onRequestPermission} className="text-xs font-bold px-3 py-1.5 rounded-lg text-white" style={{ background: "#C97A2B" }}>تفعيل الإشعارات</button>
          </div>
        </div>
      )}
      <Field label="عنوان التذكير">
        <input style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="مثال: اختبار الوحدة الثانية" />
      </Field>
      <div className="grid grid-cols-2 gap-2 mb-1">
        <Field label="التاريخ">
          <input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} />
        </Field>
        <Field label="الوقت">
          <input type="time" style={inputStyle} value={time} onChange={(e) => setTime(e.target.value)} />
        </Field>
      </div>
      <button disabled={!title.trim()} onClick={submit} className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 mb-4" style={{ background: "#26423B" }}>
        إضافة تذكير
      </button>

      {sorted.length === 0 ? (
        <p className="text-sm text-center py-6" style={{ color: MUTED }}>لا يوجد تذكيرات بعد.</p>
      ) : (
        <div className="space-y-2">
          {sorted.map((r) => {
            const due = new Date(`${r.date}T${r.time}`);
            const passed = due < now;
            return (
              <div key={r.id} className="flex items-center gap-2 p-3 rounded-xl" style={{ border: `1px solid ${LINE}`, background: passed ? "#F8F7F2" : "#fff", opacity: passed ? 0.6 : 1 }}>
                <Bell size={15} color={passed ? MUTED : "#26423B"} className="shrink-0" />
                <div className="flex-1">
                  <p className="text-sm font-semibold" style={{ color: INK }}>{r.title}</p>
                  <p className="text-xs" style={{ color: MUTED }}>{formatDateDisplay(r.date)} — {r.time}</p>
                </div>
                <button onClick={() => onDelete(r.id)} title="حذف" className="p-1.5 rounded hover:bg-black/5 shrink-0"><X size={14} color={MUTED} /></button>
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
        <button onClick={submitNew} className="px-4 py-2 rounded-lg text-sm font-bold text-white shrink-0" style={{ background: "#26423B" }}>إضافة</button>
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
                <button onClick={saveEdit} title="حفظ" className="p-1.5 rounded-lg hover:bg-black/5"><Check size={15} color="#26423B" /></button>
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
  const [selectedDate, setSelectedDate] = useState(todayKey());
  const t = selectedDate;
  const isToday = t === todayKey();
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
    <Modal title={isToday ? "نشاطي اليوم" : "نشاطي"} onClose={onClose} wide>
      <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
        <p className="text-sm" style={{ color: MUTED }}>{formatDateDisplay(t)} — {totalCount} رصد إجمالي عبر {groups.length} فصل</p>
        <DateField value={selectedDate} onChange={setSelectedDate} />
      </div>
      {groups.length === 0 ? (
        <p className="text-sm text-center py-10" style={{ color: MUTED }}>{isToday ? "لم تسجّل أي شيء اليوم بعد." : "لا يوجد رصد بهذا التاريخ."}</p>
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
              <s.icon size={16} color="#26423B" />
            </div>
            <div>
              <p className="text-sm font-bold mb-0.5" style={{ color: INK }}>{s.title}</p>
              <p className="text-xs" style={{ color: MUTED }}>{s.text}</p>
            </div>
          </div>
        ))}
      </div>
      <button onClick={onClose} className="w-full mt-4 py-2.5 rounded-xl text-sm font-bold text-white" style={{ background: "#26423B" }}>فهمت، ابدأ الآن</button>
    </Modal>
  );
}

function ScheduleMiniCard({ schedule, image, onOpen }) {
  const hasData = image || (schedule && Object.values(schedule).some((v) => v && v.trim()));
  return (
    <button
      onClick={onOpen}
      className="flex items-center gap-2.5 text-right rounded-xl px-3 py-2 mb-3 hover:opacity-90 transition-opacity"
      style={{ background: "#fff", border: `1px solid ${LINE}`, width: "100%" }}
    >
      <CalendarRange size={15} color="#26423B" className="shrink-0" />
      {image && <img src={image} alt="" className="w-6 h-6 rounded object-cover shrink-0 dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />}
      <span className="font-bold text-xs flex-1" style={{ color: INK }}>الجدول الدراسي</span>
      <span className="text-xs" style={{ color: MUTED }}>{hasData ? "عرض" : "إضافة"}</span>
      <ChevronLeft size={14} color={MUTED} className="shrink-0" />
    </button>
  );
}

// ينشئ (أو يعيد استخدام) "مؤشر مشاركة" دائم لتقرير طالب معيّن — سطر خفيف
// بجدول shared_reports يحتوي فقط رقم المعلم والفصل والطالب (بدون أي بيانات
// فعلية)، ويُقرأ لاحقًا عبر دالة RPC مقيّدة الصلاحيات وقت مسح رمز QR.
async function publishStudentReport(cls, row) {
  const { data: authData } = await supabase.auth.getUser();
  const teacherId = authData?.user?.id;
  if (!teacherId) throw new Error("يجب تسجيل الدخول أولًا.");
  const { data, error } = await supabase
    .from("shared_reports")
    .insert({ teacher_id: teacherId, class_id: cls.id, row_id: row.id })
    .select("id")
    .single();
  if (error) throw error;
  return data.id;
}

function buildSharedReportUrl(shareId) {
  return `${window.location.origin}${window.location.pathname}?shared=${shareId}`;
}

// نافذة توليد رمز QR دائم لتقرير طالب — أول مرة تُنشئ سطر مشاركة وتحفظ
// معرّفه بالطالب نفسه (row.shareId) حتى لا يتكرر إنشاء رابط جديد كل مرة؛
// المرات اللي بعدها تعيد استخدام نفس الرابط والرمز.
function StudentQrModal({ cls, row, onSaveShareId, onClose }) {
  const [state, setState] = useState(row.shareId ? "ready" : "loading");
  const [shareId, setShareId] = useState(row.shareId || null);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (row.shareId) return;
    (async () => {
      try {
        const id = await publishStudentReport(cls, row);
        setShareId(id);
        onSaveShareId(id);
        setState("ready");
      } catch (e) {
        setError(e.message || "تعذّر إنشاء الرابط.");
        setState("error");
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const url = shareId ? buildSharedReportUrl(shareId) : "";
  const qrImgSrc = shareId ? `https://api.qrserver.com/v1/create-qr-code/?size=260x260&data=${encodeURIComponent(url)}` : "";

  const copyLink = async () => {
    try { await navigator.clipboard.writeText(url); setCopied(true); setTimeout(() => setCopied(false), 2000); } catch (e) { /* ignore */ }
  };

  return (
    <Modal title={`رمز QR — ${row.name}`} onClose={onClose} accent="magic">
      <div className="p-3 rounded-xl mb-4 flex items-start gap-2" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
        <Info size={15} color="#26423B" className="shrink-0 mt-0.5" />
        <p className="text-xs" style={{ color: "#26423B" }}>
          هذا الرمز دائم ويتحدّث تلقائيًا مع أي رصد جديد — اطبعه مرة وحدة وألصقه بملف الطالب، وولي الأمر يمسحه بأي وقت ويشوف آخر تحديث بدون تسجيل دخول.
        </p>
      </div>
      {state === "loading" && (
        <div className="flex flex-col items-center py-10">
          <RefreshCw size={22} color={MUTED} className="animate-spin mb-2" />
          <p className="text-xs" style={{ color: MUTED }}>جارٍ إنشاء الرابط...</p>
        </div>
      )}
      {state === "error" && (
        <p className="text-sm text-center py-8" style={{ color: "#C0392B" }}>{error}</p>
      )}
      {state === "ready" && (
        <>
          <div className="flex justify-center mb-4">
            <img src={qrImgSrc} alt="رمز QR" width={220} height={220} style={{ borderRadius: 12, border: `1px solid ${LINE}` }} />
          </div>
          <div className="flex items-center gap-2 p-2 rounded-lg mb-3" style={{ border: `1px solid ${LINE}`, background: "#F8F7F2" }}>
            <span className="flex-1 text-xs truncate" style={{ color: MUTED, direction: "ltr", textAlign: "left" }}>{url}</span>
            <button onClick={copyLink} className="text-xs font-bold px-2.5 py-1 rounded-lg text-white shrink-0" style={{ background: "#26423B" }}>
              {copied ? "تم النسخ ✓" : "نسخ"}
            </button>
          </div>
          <a
            href={qrImgSrc}
            download={`QR-${row.name}.png`}
            className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-bold text-white transition-all hover:brightness-110"
            style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
          >
            <ImageDown size={16} /> تنزيل الرمز كصورة
          </a>
        </>
      )}
    </Modal>
  );
}

// يربط تقرير طالب (أو أي رصد بالفصل) بشاهد أداء وظيفي — إما ينشئ شاهدًا
// جديدًا أو يرفق الصورة على شاهد موجود مسبقًا، فيصير محفوظًا دائمًا بتبويب
// الشواهد كمرجع يرجع له المعلم لاحقًا.
function LinkToShawahedModal({ cls, row, entries, shawahed, onLink, onClose }) {
  const shawahedEntries = shawahed.entries || {};
  const [catKey, setCatKey] = useState(getAllShawahedCategories(shawahed)[0].key);
  const [mode, setMode] = useState("new"); // new | existing
  const [existingId, setExistingId] = useState("");
  const [title, setTitle] = useState(`تقرير أداء — ${row.name}`);
  const [notes, setNotes] = useState(`مرتبط بتقرير الطالب ${row.name} — ${cls.subject} (${cls.grade})`);
  const [generating, setGenerating] = useState(false);

  const cat = getAllShawahedCategories(shawahed).find((c) => c.key === catKey);
  const existingForCat = shawahedEntries[catKey] || [];

  const confirm = async () => {
    setGenerating(true);
    try {
      const groups = groupEntries(entries);
      let photoImageElement = null;
      if (row.photo) { try { photoImageElement = await loadImage(row.photo); } catch (e) { photoImageElement = null; } }
      const { canvas } = await buildReportCanvas({ title: `تقرير الطالب: ${row.name}`, subtitle: `${cls.subject} • ${cls.grade} • ${cls.teacher}`, groups, photoImageElement });
      const dataUrl = canvas.toDataURL("image/png");
      onLink({ catKey, mode, existingId, title: title.trim(), notes: notes.trim(), photo: dataUrl });
      onClose();
    } finally {
      setGenerating(false);
    }
  };

  return (
    <Modal title="ربط التقرير بشاهد" onClose={onClose} accent="magic">
      <p className="text-xs mb-4" style={{ color: MUTED }}>
        يُرفق نسخة من هذا التقرير كصورة داخل الشاهد المختار — يبقى محفوظًا دائمًا بتبويب الشواهد كمرجع ترجع له لاحقًا.
      </p>
      <Field label="المعيار (الفئة)">
        <select value={catKey} onChange={(e) => { setCatKey(e.target.value); setExistingId(""); }} style={inputStyle}>
          {getAllShawahedCategories(shawahed).map((c) => <option key={c.key} value={c.key}>{c.title}</option>)}
        </select>
      </Field>
      <div className="flex gap-2 mb-3">
        <button onClick={() => setMode("new")} className="flex-1 text-xs font-semibold py-2 rounded-lg" style={{ background: mode === "new" ? INK : "transparent", color: mode === "new" ? "#fff" : MUTED, border: `1px solid ${mode === "new" ? INK : LINE}` }}>شاهد جديد</button>
        <button onClick={() => setMode("existing")} disabled={existingForCat.length === 0} className="flex-1 text-xs font-semibold py-2 rounded-lg disabled:opacity-40" style={{ background: mode === "existing" ? INK : "transparent", color: mode === "existing" ? "#fff" : MUTED, border: `1px solid ${mode === "existing" ? INK : LINE}` }}>إضافة لشاهد موجود</button>
      </div>
      {mode === "new" ? (
        <>
          <Field label="عنوان الشاهد">
            <input value={title} onChange={(e) => setTitle(e.target.value)} style={inputStyle} />
          </Field>
          <Field label="ملاحظات (اختياري)">
            <textarea value={notes} onChange={(e) => setNotes(e.target.value)} style={{ ...inputStyle, minHeight: 60 }} />
          </Field>
        </>
      ) : (
        <Field label="اختر الشاهد">
          <select value={existingId} onChange={(e) => setExistingId(e.target.value)} style={inputStyle}>
            <option value="">اختر...</option>
            {existingForCat.map((e) => <option key={e.id} value={e.id}>{e.title}</option>)}
          </select>
        </Field>
      )}
      <button
        disabled={generating || (mode === "existing" && !existingId) || (mode === "new" && !title.trim())}
        onClick={confirm}
        className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
        style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
      >
        {generating ? "جارٍ الإنشاء..." : "ربط الآن"}
      </button>
    </Modal>
  );
}

function ReportModal({ cls, row, entries, reportTrash, schoolName, principalName, countryName, ministryName, logoImage, onClose, onBack, onEditEntry, onDeleteEntry, onDeleteCategory, onDeleteAllEntries, onAddNote, onRestoreLatest, onRestoreEntry, onClearTrash, onPrint, onPrintParent, onSaveShareId, shawahed, onLinkShawahed }) {
  const [editing, setEditing] = useState(false);
  const [noteText, setNoteText] = useState("");
  const [showTrash, setShowTrash] = useState(false);
  const [behaviorDraftGroup, setBehaviorDraftGroup] = useState(null);
  const [showMedicalBanner, setShowMedicalBanner] = useState(true);
  const [showCertificate, setShowCertificate] = useState(false);
  const [showBehaviorBanner, setShowBehaviorBanner] = useState(true);
  const [shareStudentError, setShareStudentError] = useState("");
  const [confirmDeleteAll, setConfirmDeleteAll] = useState(false);
  const [confirmDeleteCategory, setConfirmDeleteCategory] = useState(null);
  const [showRemedialPlan, setShowRemedialPlan] = useState(false);
  const [showQr, setShowQr] = useState(false);
  const [showLinkShawahed, setShowLinkShawahed] = useState(false);

  const addNote = () => {
    if (!noteText.trim()) return;
    const meta = nowMeta();
    onAddNote({ id: uid(), colId: "note", colName: "ملاحظة", colColor: COLORS[7].hex, value: noteText.trim(), ...meta, manual: true });
    setNoteText("");
  };

  const shareStudentReadOnly = () => {
    setShareStudentError("");
    const html = buildReadOnlyStudentHtml(cls, row, entries);
    const filename = `تقرير-${row.name}.html`;
    const blob = new Blob([html], { type: "text/html" });
    try {
      const file = new File([blob], filename, { type: "text/html" });
      if (isTouchPrimary() && navigator.canShare && navigator.canShare({ files: [file] })) {
        navigator.share({ files: [file], title: filename }).catch(() => {});
        return;
      }
    } catch (e) { /* fall through */ }
    const win = window.open("", "_blank");
    try {
      if (win) { win.document.write(html); win.document.close(); }
      else downloadBlob(blob, filename);
    } catch (e) {
      if (win) win.close();
      setShareStudentError("تعذّرت المشاركة، حاول مرة أخرى.");
    }
  };

  const groups = groupEntries(entries);
  const flaggedGroups = groups.filter((g) => {
    const col = cls.columns.find((c) => c.id === g.colId);
    return col?.behaviorFlag && g.items.length >= (col.behaviorThreshold || 3);
  });

  const toolCards = [
    { key: "edit", icon: Pencil, label: editing ? "إنهاء التعديل" : "تعديل", color: "#3B4C8C", onClick: () => setEditing((s) => !s) },
    { key: "restore", icon: RotateCcw, label: "استعادة آخر حذف", color: "#26423B", onClick: onRestoreLatest },
    { key: "trash", icon: FolderOpen, label: "سجل المحذوفات", color: "#6B7A3A", onClick: () => setShowTrash(true) },
    { key: "remedial", icon: Activity, label: "خطة علاجية", color: "#7A4E9E", onClick: () => setShowRemedialPlan(true) },
    { key: "print", icon: Printer, label: "طباعة", color: "#2E7DA6", onClick: () => onPrint() },
    { key: "printParent", icon: Users, label: "تقرير لولي الأمر", color: "#0F9D58", onClick: () => onPrintParent() },
    { key: "certificate", icon: Award, label: "شهادة تقدير", color: "#C97A2B", onClick: () => setShowCertificate(true) },
    { key: "share", icon: Share2, label: "مشاركة التقرير", color: "#5B6472", onClick: shareStudentReadOnly },
    { key: "qr", icon: QrCode, label: "رمز QR دائم", color: "#7A4E9E", onClick: () => setShowQr(true) },
    { key: "linkShawahed", icon: FileCheck, label: "ربط بشاهد", color: "#0F6B5C", onClick: () => setShowLinkShawahed(true) },
    ...(entries.length > 0 ? [{ key: "deleteAll", icon: Trash2, label: "حذف كل التقرير", color: "#C0392B", onClick: () => setConfirmDeleteAll(true) }] : []),
  ];

  return (
    <Modal title={`تقرير الطالب — ${row.name}`} onClose={onClose} onBack={onBack} lg>
      <div className="rounded-2xl p-4 mb-3" style={{ background: "#F3F1E9", border: `1px solid ${LINE}` }}>
        <div className="flex flex-wrap items-center gap-4 mb-3">
          {row.photo ? (
            <img src={row.photo} alt={row.name} className="w-12 h-12 rounded-full object-cover shrink-0 dark-mode-img-fix" style={{ border: `1px solid ${LINE}` }} />
          ) : (
            <div className="w-12 h-12 rounded-full flex items-center justify-center font-bold text-lg shrink-0" style={{ background: row.color, color: "#fff" }}>
              {(row.name || "؟").trim().charAt(0)}
            </div>
          )}
          <div className="flex-1" style={{ minWidth: 160 }}>
            <div className="flex items-center gap-2">
              <p className="font-bold text-base" style={{ color: INK }}>{row.name}</p>
              {(() => {
                const pct = attendancePercent(cls, row.id);
                if (pct === null) return null;
                const color = pct >= 90 ? "#26423B" : pct >= 75 ? "#C97A2B" : "#C0392B";
                return (
                  <span title="نسبة الحضور" className="px-2 py-0.5 rounded-full text-xs font-bold" style={{ background: `${color}18`, color }}>
                    {pct}٪ حضور
                  </span>
                );
              })()}
            </div>
            <p className="text-xs" style={{ color: MUTED }}>{cls.subject} • {cls.grade} • {cls.teacher}</p>
          </div>
          <p className="text-xs font-semibold shrink-0" style={{ color: MUTED }}>{entries.length} رصد إجمالي عبر {groups.length} تصنيف</p>
        </div>

        {row.medicalNote && row.medicalNote.trim() && (
          <div className="rounded-xl p-3 mb-3 relative" style={{ background: "#FBEDEA", border: "2px solid #C0392B" }}>
            <button
              onClick={() => setShowMedicalBanner((s) => !s)}
              title={showMedicalBanner ? "إخفاء التفاصيل" : "إظهار التفاصيل"}
              className="absolute top-1.5 left-1.5 p-1 rounded-full hover:bg-black/10"
            >
              {showMedicalBanner ? <Eye size={12} color="#9A3B2E" /> : <EyeOff size={12} color="#9A3B2E" />}
            </button>
            <div className="flex items-center gap-2">
              <AlertTriangle size={16} color="#9A3B2E" strokeWidth={2.5} />
              <p className="text-xs font-extrabold" style={{ color: "#9A3B2E" }}>تنبيه خاص</p>
            </div>
            {showMedicalBanner && (
              <p className="text-xs leading-snug mt-1.5" style={{ color: "#7A2E22", wordBreak: "break-word", whiteSpace: "pre-wrap" }}>{row.medicalNote}</p>
            )}
          </div>
        )}

        <p className="text-xs font-bold mb-2" style={{ color: MUTED }}>أدوات التقرير</p>
        <div className="flex gap-3 overflow-x-auto pb-1">
          {toolCards.map((t) => (
            <button
              key={t.key}
              onClick={t.onClick}
              className="rounded-2xl p-3 text-center shrink-0 hover:opacity-90 active:scale-95 transition-all"
              style={{ background: `${t.color}14`, minWidth: 88 }}
            >
              <div className="w-9 h-9 rounded-full mx-auto mb-2 flex items-center justify-center" style={{ background: t.color }}>
                <t.icon size={16} color="#fff" strokeWidth={2.5} />
              </div>
              <p className="text-xs font-semibold" style={{ color: INK }}>{t.label}</p>
            </button>
          ))}
        </div>
        {shareStudentError && <p className="text-xs mt-2" style={{ color: "#C0392B" }}>{shareStudentError}</p>}
      </div>

      {flaggedGroups.map((g) => (
        showBehaviorBanner ? (
          <div key={g.colId} className="flex items-center justify-between gap-3 p-3 rounded-xl mb-3 flex-wrap" style={{ background: "#FBEDEA", border: "1px solid #F0D2CB" }}>
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
        <button onClick={() => setShowBehaviorBanner(true)} className="flex items-center gap-1.5 text-xs font-semibold mb-3 px-2.5 py-1 rounded-full" style={{ color: "#9A3B2E", border: "1px solid #F0D2CB" }}>
          <EyeOff size={13} /> تنبيهات السلوك مخفية — إظهار
        </button>
      )}

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
              {(() => {
                const col = cls.columns.find((c) => c.id === g.colId);
                if (!col || col.type !== "level" || !(col.levels || []).length) return null;
                const levels = col.levels;
                return (
                  <div className="px-4 py-3" style={{ background: "#FAF8F3", borderBottom: `1px solid ${LINE}` }}>
                    <p className="text-xs font-semibold mb-2" style={{ color: MUTED }}>تطور المستوى عبر الوقت</p>
                    <div className="flex items-center gap-1 overflow-x-auto pb-1">
                      {g.items.map((e, i) => {
                        const idx = levels.findIndex((l) => l.label === e.value);
                        const dotColor = idx >= 0 ? levels[idx].color : LINE;
                        return (
                          <div key={e.id} className="flex items-center gap-1 shrink-0">
                            {i > 0 && <span style={{ width: 14, height: 1, background: LINE }} />}
                            <div className="flex flex-col items-center gap-0.5" title={`${e.date || ""} — ${e.value}`}>
                              <span className="w-3 h-3 rounded-full shrink-0" style={{ background: dotColor }} />
                              <span className="text-[9px] whitespace-nowrap" style={{ color: MUTED }}>{(e.date || "").split(" ").pop()}</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}
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
          <button onClick={addNote} className="px-4 py-2 rounded-lg text-sm font-bold text-white shrink-0" style={{ background: "#26423B" }}>إضافة</button>
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
      {showQr && (
        <StudentQrModal
          cls={cls}
          row={row}
          onSaveShareId={(id) => onSaveShareId(row.id, id)}
          onClose={() => setShowQr(false)}
        />
      )}
      {showLinkShawahed && (
        <LinkToShawahedModal
          cls={cls}
          row={row}
          entries={entries}
          shawahed={shawahed}
          onLink={onLinkShawahed}
          onClose={() => setShowLinkShawahed(false)}
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
              <td className="p-1 text-center font-bold" style={{ border: `1px solid ${LINE}`, color: status === "absent" ? "#C0392B" : "#26423B" }}>{status === "absent" ? "غائب" : "حاضر"}</td>
              {cls.columns.map((col) => {
                const val = dateKey ? valueOnDate(cls, row.id, col.id, dateKey) : (cls.cells[`${row.id}:${col.id}`] || lastReportedValue(cls, row.id, col.id));
                return <td key={col.id} className="p-1 text-center" style={{ border: `1px solid ${LINE}` }}>{val || ""}</td>;
              })}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function DisplayBoard({ cls, onClose, onPrint }) {
  const [boardDate, setBoardDate] = useState(todayKey());
  const [shareError, setShareError] = useState("");
  const shareReadOnly = () => {
    setShareError("");
    const html = buildReadOnlyBoardHtml(cls, boardDate);
    const filename = `${cls.subject || "الفصل"}-لوحة-العرض.html`;
    const blob = new Blob([html], { type: "text/html" });

    // المحاولة الأولى: صفحة اختيار المشاركة الحقيقية للجهاز (واتساب،
    // الرسائل، البريد...) — يجب استدعاؤها هنا مباشرة (بدون أي await قبلها)
    // ليعتبرها المتصفح ناتجة عن ضغطة المستخدم مباشرة.
    try {
      const file = new File([blob], filename, { type: "text/html" });
      if (isTouchPrimary() && navigator.canShare && navigator.canShare({ files: [file] })) {
        navigator.share({ files: [file], title: filename }).catch(() => {});
        return;
      }
    } catch (e) {
      // المتصفح ما يدعم مشاركة الملفات — نكمل للخيار البديل بالأسفل
    }

    // خيار بديل: فتح نافذة جديدة فيها المحتوى مباشرة (يعمل في كل المتصفحات
    // دون استثناء). نفتحها هنا بشكل متزامن أيضًا لنفس السبب أعلاه.
    const win = window.open("", "_blank");
    try {
      if (win) {
        win.document.write(html);
        win.document.close();
      } else {
        downloadBlob(blob, filename);
      }
    } catch (e) {
      if (win) win.close();
      setShareError("تعذّرت المشاركة، حاول مرة أخرى.");
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
            <IconBtn icon={Share2} label="مشاركة نسخة للقراءة فقط" onClick={shareReadOnly} />
            <button onClick={onClose} className="p-1.5 rounded-full hover:bg-black/5 active:scale-90 transition-transform"><X size={18} color={MUTED} /></button>
          </div>
        </div>
        {shareError && <p className="text-xs px-5 pt-2" style={{ color: "#C0392B" }}>{shareError}</p>}
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
      className={`rounded-2xl overflow-hidden relative card-in transition-all duration-200 hover:-translate-y-0.5 ${animating ? "trash-toss" : ""}`}
      style={{ background: "#fff", border: cls.pinned ? `2px solid #26423B` : `1px solid ${LINE}`, boxShadow: "0 1px 3px rgba(35,38,34,0.06)" }}
      onMouseEnter={(e) => { e.currentTarget.style.boxShadow = "0 10px 24px rgba(35,38,34,0.10)"; }}
      onMouseLeave={(e) => { e.currentTarget.style.boxShadow = "0 1px 3px rgba(35,38,34,0.06)"; }}
    >
      <div className="h-2" style={{ background: cls.color }} />
      <button
        title="حذف الفصل"
        onClick={(e) => { e.stopPropagation(); onDelete(cls.id); }}
        className="absolute top-2.5 left-2.5 z-10 p-1.5 rounded-full hover:opacity-80"
        style={{ background: "#FBEDEA" }}
      >
        <Trash2 size={13} color="#9A3B2E" />
      </button>
      <div className="p-4 cursor-pointer relative active:scale-[0.98] transition-transform" onClick={() => !locked && onOpen(cls.id)} style={{ cursor: locked ? "not-allowed" : "pointer" }}>
        {(cls.pinned || locked) && (
          <div className="absolute top-2 right-2 flex gap-1">
            {cls.pinned && <span className="p-1 rounded-full" style={{ background: "#EAF3F0" }}><Pin size={11} color="#26423B" /></span>}
            {locked && <span className="p-1 rounded-full" style={{ background: "#FBEDEA" }}><Lock size={11} color="#9A3B2E" /></span>}
          </div>
        )}
        <div className="flex items-center gap-2 mb-1 pl-6">
          {cls.emoji ? <span className="text-base leading-none">{cls.emoji}</span> : <BookOpen size={16} color={cls.color} />}
          <h3 className="font-bold text-base" style={{ color: INK }}>{cls.subject}</h3>
        </div>
        <p className="text-sm mb-0.5" style={{ color: MUTED }}>{cls.grade}</p>
        <p className="text-sm mb-0.5" style={{ color: MUTED }}>{cls.teacher}</p>
        <p className="text-xs" style={{ color: MUTED }}>{cls.yearHijri}{cls.yearHijri && cls.yearGregorian ? " هـ  •  " : ""}{cls.yearGregorian}{cls.yearGregorian ? " م" : ""}</p>
        {locked && <p className="text-xs mt-1 font-semibold" style={{ color: "#9A3B2E" }}>مقفل — اضغط أيقونة القفل للدخول</p>}
      </div>
      <div className="flex items-center gap-1 px-3 py-2 relative flex-wrap" style={{ borderTop: `1px solid ${LINE}` }}>
        <div className="flex items-center gap-1" style={{ borderInlineEnd: `1px solid ${LINE}`, paddingInlineEnd: 4 }}>
          <button title={cls.pinned ? "إلغاء التثبيت" : "تثبيت الفصل"} onClick={() => onTogglePin(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5">
            <Pin size={15} color={cls.pinned ? "#26423B" : MUTED} />
          </button>
          <button title={locked ? "فتح القفل" : "قفل الفصل"} onClick={() => onToggleLock(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5">
            {locked ? <Lock size={15} color="#9A3B2E" /> : <Unlock size={15} color={MUTED} />}
          </button>
        </div>
        {!cls.pinned && (
          <div className="flex items-center gap-1" style={{ borderInlineEnd: `1px solid ${LINE}`, paddingInlineEnd: 4 }}>
            <button title="نقل لأعلى" disabled={isFirst} onClick={() => onMove(cls.id, -1)} className="p-1.5 rounded-lg hover:bg-black/5 disabled:opacity-30"><ChevronUp size={15} color={MUTED} /></button>
            <button title="نقل لأسفل" disabled={isLast} onClick={() => onMove(cls.id, 1)} className="p-1.5 rounded-lg hover:bg-black/5 disabled:opacity-30"><ChevronDown size={15} color={MUTED} /></button>
          </div>
        )}
        <div className="flex items-center gap-1" style={{ borderInlineEnd: `1px solid ${LINE}`, paddingInlineEnd: 4 }}>
          <button title="تعديل" onClick={() => onEdit(cls)} className="p-1.5 rounded-lg hover:bg-black/5"><Pencil size={15} color={MUTED} /></button>
          <button title="الوان" onClick={() => setShowColors((s) => !s)} className="p-1.5 rounded-lg hover:bg-black/5"><Palette size={15} color={MUTED} /></button>
        </div>
        <div className="flex items-center gap-1">
          <button title="تكرار الفصل" onClick={() => onDuplicate(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5"><Copy size={15} color={MUTED} /></button>
          <button title="بدء فصل دراسي جديد" onClick={() => onNewTerm(cls)} className="p-1.5 rounded-lg hover:bg-black/5"><CalendarPlus size={15} color={MUTED} /></button>
        </div>
        <button title={cls.archived ? "إلغاء الأرشفة" : "أرشفة"} onClick={() => onArchive(cls.id)} className="p-1.5 rounded-lg hover:bg-black/5 mr-auto">
          {cls.archived ? <ArchiveRestore size={15} color={MUTED} /> : <Archive size={15} color={MUTED} />}
        </button>
        {showColors && (
          <div className="absolute bottom-12 right-3 z-20 p-2 rounded-xl shadow-lg" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <ColorSwatches value={cls.color} onChange={(hex) => { onColor(cls.id, hex); setShowColors(false); }} />
          </div>
        )}
      </div>
    </div>
  );
}

function HomePage({ data, setData, onOpen, userEmail, userId, onSignOut, siteSettings, updateSiteSettings, isOwner, isOnline, syncStatus }) {
  const [modal, setModal] = useState(null);
  const [mainTab, setMainTab] = useState("classes");
  const [tab, setTab] = useState("active");
  const [toolsExpanded, setToolsExpanded] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [showTodayActivity, setShowTodayActivity] = useState(false);
  const [shawahedPreview, setShawahedPreview] = useState(null);
  const [showGradeSheetFlow, setShowGradeSheetFlow] = useState(false);
  const [gradeSheetClassId, setGradeSheetClassId] = useState(null);
  const [gradeSheetPreview, setGradeSheetPreview] = useState(null);
  const [showAdminPanel, setShowAdminPanel] = useState(false);
  const [showTestBuilder, setShowTestBuilder] = useState(false);
  const [gradingTestId, setGradingTestId] = useState(null);
  const [printingTestId, setPrintingTestId] = useState(null);
  const [analyzingTestId, setAnalyzingTestId] = useState(null);
  const [omrTestId, setOmrTestId] = useState(null);
  const addTest = (test) => setData((d) => ({ ...d, tests: [...(d.tests || []), test] }));
  const deleteTest = (id) => setData((d) => ({ ...d, tests: (d.tests || []).filter((t) => t.id !== id) }));
  const archiveTest = (id) => setData((d) => {
    const test = (d.tests || []).find((t) => t.id === id);
    if (!test) return d;
    return { ...d, tests: d.tests.filter((t) => t.id !== id), archivedTests: [...(d.archivedTests || []), test] };
  });
  const restoreTest = (id) => setData((d) => {
    const test = (d.archivedTests || []).find((t) => t.id === id);
    if (!test) return d;
    return { ...d, archivedTests: d.archivedTests.filter((t) => t.id !== id), tests: [...(d.tests || []), test] };
  });
  const deleteArchivedTestForever = (id) => setData((d) => ({ ...d, archivedTests: (d.archivedTests || []).filter((t) => t.id !== id) }));
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
  const addFooterContact = () => updateSiteSettings((s) => ({ ...s, footerContacts: [...(s.footerContacts || []), { id: uid(), label: "", value: "" }] }));
  const updateFooterContact = (id, patch) => updateSiteSettings((s) => ({ ...s, footerContacts: (s.footerContacts || []).map((c) => (c.id === id ? { ...c, ...patch } : c)) }));
  const removeFooterContact = (id) => updateSiteSettings((s) => ({ ...s, footerContacts: (s.footerContacts || []).filter((c) => c.id !== id) }));
  const addFooterBadge = (image) => updateSiteSettings((s) => ({ ...s, footerBadges: [...(s.footerBadges || []), { id: uid(), image }] }));
  const removeFooterBadge = (id) => updateSiteSettings((s) => ({ ...s, footerBadges: (s.footerBadges || []).filter((b) => b.id !== id) }));
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
    <div className="max-w-5xl xl:max-w-7xl 2xl:max-w-[1600px] mx-auto px-4 md:px-6 py-6 page-fade-in">
      {siteSettings.announcementActive && siteSettings.announcement && (
        <div className="flex items-start gap-2 p-3 rounded-xl mb-4" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
          <Info size={16} color="#26423B" className="shrink-0 mt-0.5" />
          <p className="text-sm" style={{ color: "#26423B" }}>{siteSettings.announcement}</p>
        </div>
      )}
      <div className="sticky top-0 z-20 pb-2" style={{ background: PAPER }}>
        <div className="flex items-center justify-between mb-5">
          <div>
            {siteSettings.siteLogo ? (
              <img src={siteSettings.siteLogo} alt="فصولي" className="max-h-12 object-contain" />
            ) : (
              <h1 className="text-2xl font-extrabold" style={{ color: INK, fontFamily: "'Cairo', sans-serif" }}>فصولي</h1>
            )}
            <p className="text-sm mt-1" style={{ color: MUTED }}>{siteSettings.siteTagline || "فصولك الدراسية في مكان واحد"}</p>
          </div>
          <div className="flex items-center gap-2">
            <SyncStatusBadge isOnline={isOnline} syncStatus={syncStatus} />
            {isOwner && (
              <button onClick={() => setShowAdminPanel(true)} title="لوحة التحكم" className="w-10 h-10 rounded-full flex items-center justify-center hover:opacity-90" style={{ background: "#26423B" }}>
                <ShieldCheck size={18} color="#fff" />
              </button>
            )}
            <button onClick={() => setShowGuide(true)} title="كيف أبدأ؟" className="w-10 h-10 rounded-full flex items-center justify-center hover:bg-black/5" style={{ border: `1px solid ${LINE}` }}>
              <HelpCircle size={18} color={MUTED} />
            </button>
            <button onClick={() => setShowSettings(true)} title="الإعدادات" className="w-10 h-10 rounded-full flex items-center justify-center hover:bg-black/5" style={{ border: `1px solid ${LINE}` }}>
              <Settings size={18} color={MUTED} />
            </button>
          </div>
        </div>

        <div className="flex items-center gap-1.5 mb-4 overflow-x-auto pb-1">
          {[
            { key: "classes", label: "الفصول", icon: BookOpen },
            { key: "shawahed", label: "شواهد", icon: FileCheck },
            { key: "tests", label: "الاختبارات", icon: ListChecks },
            { key: "library", label: "المكتبة", icon: FolderOpen },
            { key: "games", label: "الأدوات التفاعلية", icon: Gamepad2 },
            { key: "archive", label: "المؤرشفة", icon: Archive },
          ].map((t) => (
            <button
              key={t.key}
              onClick={() => setMainTab(t.key)}
              className="flex items-center gap-1.5 px-4 py-2 rounded-full text-sm font-bold shrink-0 transition-all"
              style={mainTab === t.key
                ? { background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})`, color: "#fff", boxShadow: "0 2px 8px rgba(38,66,59,0.28)" }
                : { background: "#fff", color: MUTED, border: `1px solid ${LINE}` }}
            >
              <t.icon size={15} /> {t.label}
            </button>
          ))}
        </div>


        {mainTab === "classes" && (<>
        <ScheduleMiniCard schedule={data.schedule} image={data.scheduleImage} onOpen={() => setShowSchedule(true)} />

        <div className="flex flex-wrap items-center gap-2 mb-2.5">
          <button onClick={() => setTab("active")} className="px-4 py-1.5 rounded-full text-sm font-semibold"
            style={{ background: tab === "active" ? INK : "transparent", color: tab === "active" ? "#fff" : MUTED, border: `1px solid ${tab === "active" ? INK : LINE}` }}>
            النشطة ({data.classes.filter((c) => !c.archived).length})
          </button>
          <button onClick={() => setModal({ mode: "add" })} className="mr-auto flex items-center gap-1.5 px-4 py-2.5 rounded-full text-sm font-bold text-white transition-all hover:brightness-105 hover:-translate-y-px active:scale-95 active:translate-y-0" style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})`, boxShadow: "0 3px 10px rgba(38,66,59,0.32)" }}>
            <Plus size={18} strokeWidth={2.5} /> إضافة فصل جديد
          </button>
        </div>

        <button
          onClick={() => setToolsExpanded((s) => !s)}
          className="w-full flex items-center justify-center gap-1.5 py-1.5 mb-2.5 rounded-xl text-xs font-bold hover:opacity-80"
          style={{ background: "#fff", border: `1px solid ${LINE}`, color: MUTED }}
        >
          {toolsExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          {toolsExpanded ? "إخفاء الأدوات" : "إظهار الأدوات"}
        </button>

        {toolsExpanded && (
        <div className="rounded-2xl p-2.5 flex flex-wrap items-center gap-2 mb-2.5" style={{ background: "#fff", boxShadow: "0 2px 8px rgba(35,38,34,0.06)", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>أدوات</span>
          <div className="flex flex-wrap items-center gap-2" style={{ borderInlineEnd: `1px solid ${LINE}`, paddingInlineEnd: 8 }}>
            <IconBtn icon={Search} label="بحث عن طالب في كل الفصول" onClick={() => setShowSearch((s) => !s)} />
            <IconBtn icon={ListTodo} label="نشاطي اليوم" onClick={() => setShowTodayActivity(true)} />
            <IconBtn icon={FileSpreadsheet} label="كشف رصد درجات" magic onClick={() => setShowGradeSheetFlow(true)} />
          </div>
          <div className="flex flex-wrap items-center gap-2" style={{ borderInlineEnd: classes.length > 0 ? `1px solid ${LINE}` : "none", paddingInlineEnd: 8 }}>
            <IconBtn icon={RotateCcw} label="تراجع" onClick={restoreLatestClass} />
            <IconBtn icon={FolderOpen} label="استعادة" onClick={() => setShowClassTrash(true)} />
          </div>
          {classes.length > 0 && (
            <div className="flex flex-wrap items-center gap-2">
              <IconBtn icon={Archive} label={`أرشفة الكل (${classes.length})`} onClick={archiveAllClasses} />
              <IconBtn icon={Trash2} label={`حذف الكل (${classes.length})`} tone="danger" onClick={() => setConfirmDeleteAllClasses(true)} />
            </div>
          )}
        </div>
        )}
        </>)}
      </div>

      <div className="mt-5">
      {mainTab === "classes" && (<>
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
                        <span className="text-xs font-bold px-2 py-0.5 rounded-full shrink-0" style={{ background: status === "absent" ? "#F5DEDB" : "#E3F0ED", color: status === "absent" ? "#C0392B" : "#26423B" }}>
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
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
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
      </>)}
      </div>

      {mainTab === "shawahed" && (
        <ShawahedHub
          bare
          shawahed={data.shawahed || {}}
          teacherName={data.classes[0]?.teacher || ""}
          onUpdate={(next) => setData((d) => ({ ...d, shawahed: next }))}
          onExport={(selectedKeys, description) => setShawahedPreview({
            type: "shawahedReportV2",
            shawahed: data.shawahed || {},
            selectedKeys,
            meta: {
              countryName: data.settings?.countryName,
              ministryName: data.settings?.ministryName,
              schoolName: data.settings?.schoolName,
              logoImage: data.settings?.logoImage,
              teacherName: data.classes[0]?.teacher || "",
              principalName: data.settings?.principalName,
              description,
            },
          })}
          onQuickPrint={(catKey) => {
            const cat = getAllShawahedCategories(data.shawahed || {}).find((c) => c.key === catKey);
            const count = (data.shawahed?.entries?.[catKey] || []).length;
            const teacherName = data.classes[0]?.teacher || "";
            const autoDesc = cat
              ? `يوثّق هذا التقرير ${count} شاهدًا على أداء ${teacherName || "المعلم/ـة"} فيما يخص معيار "${cat.title}"، ويُظهر التزامه/ـا التطبيقي بهذا الجانب من الأداء الوظيفي.`
              : "";
            setShawahedPreview({
              type: "shawahedReportV2",
              shawahed: data.shawahed || {},
              selectedKeys: [catKey],
              meta: {
                countryName: data.settings?.countryName,
                ministryName: data.settings?.ministryName,
                schoolName: data.settings?.schoolName,
                logoImage: data.settings?.logoImage,
                teacherName,
                principalName: data.settings?.principalName,
                description: autoDesc,
              },
            });
          }}
          onShareReadOnly={(selectedKeys, description) => {
            const html = buildReadOnlyShawahedHtml(data.shawahed || {}, selectedKeys, {
              countryName: data.settings?.countryName,
              ministryName: data.settings?.ministryName,
              schoolName: data.settings?.schoolName,
              teacherName: data.classes[0]?.teacher || "",
              principalName: data.settings?.principalName,
              description,
            });
            const blob = new Blob([html], { type: "text/html" });
            shareOrDownloadFile(blob, "تقرير-شواهد.html", "text/html");
          }}
          onPrintProgram={(entry, cat) => setShawahedPreview({
            type: "programReport",
            entry,
            cat,
            meta: {
              countryName: data.settings?.countryName,
              ministryName: data.settings?.ministryName,
              schoolName: data.settings?.schoolName,
              logoImage: data.settings?.logoImage,
              teacherName: data.classes[0]?.teacher || "",
              region: data.settings?.region,
              office: data.settings?.office,
            },
          })}
          onShareReadOnlyOne={(entry, cat) => {
            const fakeShawahed = { entries: { [cat.key]: [entry] } };
            const html = buildReadOnlyShawahedHtml(fakeShawahed, [cat.key], {
              countryName: data.settings?.countryName,
              ministryName: data.settings?.ministryName,
              schoolName: data.settings?.schoolName,
              teacherName: data.classes[0]?.teacher || "",
              principalName: data.settings?.principalName,
              description: "",
            });
            const blob = new Blob([html], { type: "text/html" });
            shareOrDownloadFile(blob, `شاهد-${entry.title}.html`, "text/html");
          }}
        />
      )}
      {mainTab === "tests" && (
        <TestsListModal
          bare
          tests={data.tests || []}
          onCreateNew={() => setShowTestBuilder(true)}
          onGrade={(id) => setGradingTestId(id)}
          onGradeCamera={(id) => setOmrTestId(id)}
          onPrint={(id) => setPrintingTestId(id)}
          onDelete={deleteTest}
          onArchive={archiveTest}
          onAnalyze={(id) => setAnalyzingTestId(id)}
        />
      )}
      {mainTab === "library" && (
        <LibraryHub
          bare
          library={data.library || []}
          classes={data.classes}
          onUpload={(item) => setData((d) => ({ ...d, library: [...(d.library || []), item] }))}
          onDelete={(id) => setData((d) => ({ ...d, library: (d.library || []).filter((x) => x.id !== id) }))}
          onAssign={(id, classId) => setData((d) => ({ ...d, library: (d.library || []).map((x) => (x.id === id ? { ...x, classId } : x)) }))}
        />
      )}
      {mainTab === "games" && (
        <GamesHub
          bare
          classes={data.classes}
          library={data.library || []}
          updateClassById={(classId, fn) => setData((d) => ({ ...d, classes: d.classes.map((c) => (c.id === classId ? fn(c) : c)) }))}
        />
      )}
      {mainTab === "archive" && (
        <UnifiedArchiveModal
          bare
          archivedClasses={data.classes.filter((c) => c.archived)}
          shawahed={data.shawahed || {}}
          archivedTests={data.archivedTests || []}
          onRestoreClass={toggleArchive}
          onRestoreShawahed={(catKey, entryId) => {
            const shawahed = data.shawahed || {};
            const entries = shawahed.entries || {};
            const archivedEntries = shawahed.archivedEntries || {};
            const entry = (archivedEntries[catKey] || []).find((e) => e.id === entryId);
            if (!entry) return;
            setData((d) => ({
              ...d,
              shawahed: {
                ...shawahed,
                entries: { ...entries, [catKey]: [...(entries[catKey] || []), entry] },
                archivedEntries: { ...archivedEntries, [catKey]: archivedEntries[catKey].filter((e) => e.id !== entryId) },
              },
            }));
          }}
          onRestoreTest={restoreTest}
          onDeleteTestForever={deleteArchivedTestForever}
        />
      )}

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
          themeColor={data.settings?.themeColor}
          density={data.settings?.density}
          fontScale={data.settings?.fontScale}
          schoolName={data.settings?.schoolName}
          principalName={data.settings?.principalName}
          countryName={data.settings?.countryName}
          ministryName={data.settings?.ministryName}
          logoImage={data.settings?.logoImage}
          teacherPhoto={data.settings?.teacherPhoto}
          onChangeSchoolInfo={(patch) => setData((d) => ({ ...d, settings: { ...(d.settings || {}), ...patch } }))}
          footerContacts={siteSettings.footerContacts}
          footerBadges={siteSettings.footerBadges}
          onAddContact={addFooterContact}
          onUpdateContact={updateFooterContact}
          onRemoveContact={removeFooterContact}
          onAddBadge={addFooterBadge}
          onRemoveBadge={removeFooterBadge}
          userEmail={userEmail}
          onSignOut={onSignOut}
          isOwner={isOwner}
          onBackupData={() => {
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
            downloadBlob(blob, `فصولي-نسخة-احتياطية-${todayKey()}.json`);
          }}
          onRestoreData={(parsed) => setData(parsed)}
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
      {showAdminPanel && <AdminPanelModal currentUserId={userId} siteSettings={siteSettings} updateSiteSettings={updateSiteSettings} onClose={() => setShowAdminPanel(false)} />}
      {shawahedPreview && (
        <PrintPreviewModal
          job={shawahedPreview}
          format="pdf"
          onClose={() => setShawahedPreview(null)}
          onExport={(key) => {
            if (key === "pdf") exportPdfShare(shawahedPreview);
            else if (key === "png") exportPng(shawahedPreview);
            else if (key === "excel") exportExcel(shawahedPreview);
            setShawahedPreview(null);
          }}
        />
      )}
      {showGradeSheetFlow && !gradeSheetClassId && (
        <GradeSheetClassPicker
          classes={data.classes}
          onSelect={(id) => setGradeSheetClassId(id)}
          onClose={() => setShowGradeSheetFlow(false)}
        />
      )}
      {showGradeSheetFlow && gradeSheetClassId && (() => {
        const cls = data.classes.find((c) => c.id === gradeSheetClassId);
        if (!cls) { setGradeSheetClassId(null); return null; }
        return (
          <GradeSheetModal
            cls={cls}
            onClose={() => { setShowGradeSheetFlow(false); setGradeSheetClassId(null); }}
            onBack={() => setGradeSheetClassId(null)}
            onGenerate={({ shortTestIds, finalExamIds, reviewerName }) => {
              setGradeSheetPreview({ type: "gradeSheet", cls, shortTestIds, finalExamIds, reviewerName });
              setShowGradeSheetFlow(false);
              setGradeSheetClassId(null);
            }}
          />
        );
      })()}
      {gradeSheetPreview && (
        <PrintPreviewModal
          job={gradeSheetPreview}
          format="pdf"
          onClose={() => setGradeSheetPreview(null)}
          onExport={(key) => {
            if (key === "pdf") exportPdfShare(gradeSheetPreview);
            else if (key === "png") exportPng(gradeSheetPreview);
            else if (key === "excel") exportExcel(gradeSheetPreview);
            setGradeSheetPreview(null);
          }}
        />
      )}
      {showTestBuilder && (
        <TestBuilderModal
          questionBank={data.questionBank || []}
          onAddToBank={(qs) => setData((d) => ({ ...d, questionBank: [...(d.questionBank || []), ...qs] }))}
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
      {analyzingTestId && (
        <TestAnalysisModal
          test={(data.tests || []).find((t) => t.id === analyzingTestId)}
          onClose={() => setAnalyzingTestId(null)}
        />
      )}
      {isOwner && <SiteFooter contacts={siteSettings.footerContacts} badges={siteSettings.footerBadges} />}
    </div>
  );
}

// ---------- Class detail page ----------

function ClassPage({ cls, updateClass, onBack, requestPrint, feedbackEnabled, schoolName, principalName, countryName, ministryName, logoImage, allClasses, onMoveRowsToClass, isOwner, density, isOnline, syncStatus, shawahed, onLinkShawahed }) {
  const [colModal, setColModal] = useState(null);
  const [rowModal, setRowModal] = useState(null);
  const [showSearch, setShowSearch] = useState(false);
  const [search, setSearch] = useState("");
  const [toolsExpanded, setToolsExpanded] = useState(false);
  const [showBoard, setShowBoard] = useState(false);
  const [reportRowId, setReportRowId] = useState(null);
  const [confirmAction, setConfirmAction] = useState(null); // { type: 'deleteRow'|'deleteAll', row? }
  const [animatingRowId, setAnimatingRowId] = useState(null);
  const [animatingColId, setAnimatingColId] = useState(null);
  const [tableFading, setTableFading] = useState(false);
  const [showAttendance, setShowAttendance] = useState(false);
  const [showEvents, setShowEvents] = useState(false);
  const [showReminders, setShowReminders] = useState(false);
  const [showRandomPicker, setShowRandomPicker] = useState(false);
  const [showRandomGroups, setShowRandomGroups] = useState(false);
  const [selectedRowIds, setSelectedRowIds] = useState([]);
  const [activeFilter, setActiveFilter] = useState(null);
  const [showFilterModal, setShowFilterModal] = useState(false);
  const [showBulkRecordModal, setShowBulkRecordModal] = useState(false);
  const [showBulkCertificateModal, setShowBulkCertificateModal] = useState(false);
  const [showMoveModal, setShowMoveModal] = useState(false);
  const [confirmBulkDelete, setConfirmBulkDelete] = useState(false);
  const [showReportPicker, setShowReportPicker] = useState(false);
  const [showGradeSheet, setShowGradeSheet] = useState(false);
  const [showPeriodComparison, setShowPeriodComparison] = useState(false);
  const [showExamMode, setShowExamMode] = useState(false);
  const [showRowColorRule, setShowRowColorRule] = useState(false);
  const [showStatsRow, setShowStatsRow] = useState(false);
  const examLocked = isExamModeActive(cls);
  const compactDensity = density === "compact";
  const cellPadClass = compactDensity ? "p-0.5" : "p-1";
  const namePadClass = compactDensity ? "p-1" : "p-2";
  const [blinkRowId, setBlinkRowId] = useState(null);
  const timers = useRef({});

  const [toast, setToast] = useState(null);
  const toastTimer = useRef(null);
  const showToast = (msg) => {
    setToast(msg);
    if (toastTimer.current) clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 2200);
  };

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
    if (value && value.trim()) {
      playFeedback(feedbackEnabled);
      showToast(`تم رصد "${col.name}" — ${row.name}`);
    }
    if (col.autoRenew && value && value.trim()) {
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
      const newCols = drafts.map((d) => ({ id: uid(), name: d.name.trim(), type: d.type, options: d.options, levels: d.levels || [], color: d.color, maxValue: d.maxValue || "", colorScale: !!d.colorScale, colorBands: d.colorBands || [], autoRenew: !!d.autoRenew, pinned: !!d.pinned, behaviorFlag: !!d.behaviorFlag, behaviorThreshold: d.behaviorThreshold || 3 }));
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
    updateClass((c) => ({ ...c, rows: [...c.rows, ...drafts.map((d) => ({ id: uid(), name: d.name.trim(), type: d.type, options: d.options, color: d.color, autoRenew: !!d.autoRenew, medicalNote: d.medicalNote || "", nationalId: d.nationalId || "" }))] }));
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
    const row = cls.rows.find((r) => r.id === rowId);
    if (row) showToast(`تم تسجيل غياب ${row.name}`);
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

  const addReminder = (reminder) => updateClass((c) => ({ ...c, reminders: [...(c.reminders || []), reminder] }));
  const deleteReminder = (id) => updateClass((c) => ({ ...c, reminders: (c.reminders || []).filter((r) => r.id !== id) }));
  const [notifPermission, setNotifPermission] = useState(typeof Notification !== "undefined" ? Notification.permission : "unsupported");
  const requestNotifPermission = async () => {
    if (typeof Notification === "undefined") return;
    const perm = await Notification.requestPermission();
    setNotifPermission(perm);
  };
  // يجدول إشعارًا فعليًا بالجهاز لكل تذكير مستقبلي — يعمل ما دام الموقع
  // مفتوحًا (حتى بتبويب خلفي)، ولا يعمل والمتصفح مغلق تمامًا (يحتاج خادم).
  useEffect(() => {
    if (notifPermission !== "granted") return;
    const timers = (cls.reminders || [])
      .filter((r) => !r.notified)
      .map((r) => {
        const due = new Date(`${r.date}T${r.time}`).getTime();
        const ms = due - Date.now();
        if (ms <= 0 || ms > 24 * 60 * 60 * 1000) return null;
        return setTimeout(() => {
          new Notification(r.title, { body: `${cls.subject} — تذكير مجدول`, icon: "/favicon.ico" });
          updateClass((c) => ({ ...c, reminders: (c.reminders || []).map((x) => (x.id === r.id ? { ...x, notified: true } : x)) }));
        }, ms);
      })
      .filter(Boolean);
    return () => timers.forEach(clearTimeout);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cls.reminders, notifPermission]);

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
  const clearAllStudents = () => setConfirmAction({ type: "clearStudents" });

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
    if (confirmAction.type === "clearStudents") {
      setConfirmAction(null);
      setTableFading(true);
      setTimeout(() => {
        updateClass((c) => {
          const next = { ...c, rows: [], cells: {}, reports: {} };
          return pushTrash(next, { id: uid(), type: "clearStudents", when: nowMeta(), data: { rows: c.rows, cells: c.cells, reports: c.reports } });
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
      } else if (entry.type === "clearStudents") {
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

  const [previewJob, setPreviewJob] = useState(null);
  const openPrintPreview = (job, format = "pdf") => setPreviewJob({ job, format });

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
  const TOTAL_GRADE_W = 110;
  const hasTotalGradeCol = cls.columns.some((c) => c.type === "counter" && Number(c.maxValue) > 0);
  const nameEndOffset = (cls.showRowNumbers ? NUM_W : 0) + NAME_W;
  const leadWidth = nameEndOffset + (hasTotalGradeCol ? TOTAL_GRADE_W : 0);
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
    <div className="max-w-[1800px] mx-auto px-4 py-6 page-fade-in class-print-area">
      {toast && (
        <div className="fixed top-4 inset-x-0 z-[60] flex justify-center pointer-events-none">
          <div className="flex items-center gap-2 px-4 py-2.5 rounded-full shadow-lg toast-pop" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <span className="w-5 h-5 rounded-full flex items-center justify-center shrink-0" style={{ background: "#26423B" }}>
              <Check size={12} color="#fff" strokeWidth={3} />
            </span>
            <span className="text-sm font-semibold" style={{ color: INK }}>{toast}</span>
          </div>
        </div>
      )}
      <div className="sticky top-0 z-20 pb-2" style={{ background: PAPER }}>
        <div className="flex items-center justify-between mb-3">
          <button onClick={onBack} className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-70" style={{ color: MUTED }}>
            <ArrowRight size={16} /> رجوع للفصول
          </button>
          <SyncStatusBadge isOnline={isOnline} syncStatus={syncStatus} />
        </div>

        <div className="rounded-xl px-2.5 py-1.5 mb-1.5 flex flex-wrap items-center gap-x-4 gap-y-0.5" style={{ background: "#fff", border: `1px solid ${LINE}`, borderInlineStart: `4px solid ${cls.color}` }}>
          <p className="text-sm font-bold flex items-center gap-1" style={{ color: INK }}>{cls.emoji && <span>{cls.emoji}</span>}{cls.subject}</p>
          <p className="text-xs" style={{ color: MUTED }}>{cls.grade}</p>
          <p className="text-xs" style={{ color: MUTED }}>{cls.teacher}</p>
          <p className="text-xs" style={{ color: MUTED }}>{cls.yearHijri} هـ / {cls.yearGregorian} م</p>
        </div>

        {examLocked && (
          <ExamModeBanner
            examModeUntil={cls.examModeUntil}
            onUnlock={() => updateClass((c) => ({ ...c, examModeUntil: null }))}
          />
        )}

        <EventsTicker events={cls.events} speed={cls.tickerSpeed || 14} />

        <button
          onClick={() => setToolsExpanded((s) => !s)}
          className="w-full flex items-center justify-center gap-1.5 py-1.5 mb-1.5 rounded-xl text-xs font-bold hover:opacity-80"
          style={{ background: "#fff", border: `1px solid ${LINE}`, color: MUTED }}
        >
          {toolsExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          {toolsExpanded ? "إخفاء الأدوات" : "إظهار الأدوات"}
        </button>

        {toolsExpanded && (<>
        <div className="rounded-2xl p-1.5 mb-1.5 flex flex-wrap items-center gap-1.5" style={{ background: "#fff", boxShadow: "0 2px 8px rgba(35,38,34,0.06)", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>العرض والطباعة</span>
          <IconBtn icon={LayoutGrid} label="لوحة العرض" onClick={() => setShowBoard(true)} />
          <IconBtn icon={Printer} label="طباعة" onClick={printCurrentScreen} />
          <IconBtn icon={FileOutput} label="طباعة الجدول مفرغ" onClick={() => openPrintPreview({ type: "blank", cls })} />
          <IconBtn icon={Palette} label={cls.rowColorRule ? "تعديل تلوين الصف" : "تلوين الصف حسب شرط"} onClick={() => setShowRowColorRule(true)} />
          <IconBtn icon={BarChart3} label={showStatsRow ? "إخفاء صف المعدل" : "إظهار صف المعدل"} onClick={() => setShowStatsRow((s) => !s)} />
        </div>

        <div className="rounded-2xl p-1.5 mb-1.5 flex flex-wrap items-center gap-1.5" style={{ background: "#fff", boxShadow: "0 2px 8px rgba(35,38,34,0.06)", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>أدوات الحصة</span>
          <IconBtn icon={CalendarCheck} label="متابعة الحضور" onClick={() => setShowAttendance(true)} />
          <IconBtn icon={Newspaper} label="الأحداث" onClick={() => setShowEvents(true)} />
          <IconBtn icon={Bell} label="تذكير" onClick={() => setShowReminders(true)} />
          <IconBtn icon={Shuffle} label="اختر لي طالبًا" onClick={() => setShowRandomPicker(true)} />
          <IconBtn icon={Users} label="مجموعات عشوائية" onClick={() => setShowRandomGroups(true)} />
        </div>

        <div className="rounded-2xl p-1.5 mb-2 flex flex-wrap items-center gap-1.5" style={{ background: "#fff", boxShadow: "0 2px 8px rgba(35,38,34,0.06)", border: `1px solid ${LINE}` }}>
          <span className="text-xs font-bold px-1 shrink-0" style={{ color: MUTED }}>إدارة الجدول</span>
          <IconBtn icon={Plus} label="إضافة عمود" tone="primary" disabled={examLocked} onClick={() => setColModal({ mode: "add" })} />
          <IconBtn icon={Plus} label="إضافة صف" tone="primary" disabled={examLocked} onClick={() => setRowModal({ mode: "add" })} />
          <IconBtn icon={FileText} label="تقرير" magic onClick={() => setShowReportPicker(true)} />
          <IconBtn icon={ClipboardList} label="تقرير شامل للفصل" magic onClick={() => openPrintPreview({ type: "classFullReport", cls }, "pdf")} />
          <IconBtn icon={FileSpreadsheet} label="كشف رصد درجات" magic onClick={() => setShowGradeSheet(true)} />
          <IconBtn icon={BarChart3} label="مقارنة أداء بين فترتين" onClick={() => setShowPeriodComparison(true)} />
          <IconBtn icon={RotateCcw} label="تراجع" disabled={examLocked} onClick={restoreLatest} />
          <IconBtn icon={FolderOpen} label="استعادة" disabled={examLocked} onClick={() => setShowTrash(true)} />
          <IconBtn icon={Trash2} label="حذف الكل" tone="danger" disabled={examLocked} onClick={deleteAll} />
          <IconBtn icon={UserX} label="تفريغ الطلاب (يبقي الأعمدة)" tone="danger" disabled={examLocked} onClick={clearAllStudents} />
          <IconBtn
            icon={examLocked ? Unlock : Lock}
            label={examLocked ? "إنهاء وضع الاختبار" : "وضع الاختبار"}
            tone={examLocked ? "danger" : "default"}
            onClick={() => (examLocked ? updateClass((c) => ({ ...c, examModeUntil: null })) : setShowExamMode(true))}
          />
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
                <button onClick={() => setShowFilterModal(true)} className="flex items-center gap-1.5 px-3 py-2 text-sm font-medium" style={{ background: "#EAF3F0", color: "#26423B" }}>
                  <Filter size={15} />
                  {cls.columns.find((c) => c.id === activeFilter.colId)?.name}
                </button>
                <button onClick={() => setActiveFilter(null)} title="إزالة التصفية" className="px-2 py-2 hover:bg-black/5" style={{ background: "#EAF3F0" }}>
                  <X size={14} color="#26423B" />
                </button>
              </div>
            ) : (
              <IconBtn icon={Filter} label="تصفية" onClick={() => setShowFilterModal(true)} />
            )
          )}
        </div>
        </>)}
      </div>

      <div className="mt-3">
        {selectedRowIds.length > 0 && (
          <div className="flex flex-wrap items-center gap-2 mb-3 p-2.5 rounded-xl" style={{ background: "#EAF3F0", border: "1px solid #C9E2DB" }}>
            <span className="text-sm font-bold px-2" style={{ color: "#26423B" }}>{selectedRowIds.length} طالب محدد</span>
            <button onClick={selectAllRows} className="text-xs font-semibold px-2 py-1 rounded-lg hover:bg-black/5" style={{ color: INK }}>تحديد الكل</button>
            <button onClick={clearSelection} className="text-xs font-semibold px-2 py-1 rounded-lg hover:bg-black/5" style={{ color: INK }}>إلغاء التحديد</button>
            <div className="mr-auto flex gap-2">
              <IconBtn icon={Send} label="نقل إلى فصل آخر" disabled={examLocked} onClick={() => setShowMoveModal(true)} />
              <IconBtn icon={Award} label="شهادات جماعية" onClick={() => setShowBulkCertificateModal(true)} />
              <IconBtn icon={Users} label="رصد جماعي" onClick={() => setShowBulkRecordModal(true)} />
              <IconBtn icon={Trash2} label="حذف المحددين" tone="danger" disabled={examLocked} onClick={() => setConfirmBulkDelete(true)} />
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
                    style={{ background: GOLD_LIGHT, border: `1px solid ${LINE}`, position: "sticky", top: 0, insetInlineStart: 0, zIndex: 9, width: NUM_W, minWidth: NUM_W }}
                  >#</th>
                )}
                <th
                  className="p-2 text-right"
                  style={{ background: GOLD_LIGHT, border: `1px solid ${LINE}`, position: "sticky", top: 0, insetInlineStart: cls.showRowNumbers ? NUM_W : 0, zIndex: 9, width: NAME_W, minWidth: NAME_W }}
                >الاسم</th>
                {hasTotalGradeCol && (
                  <th
                    className="p-2 text-center"
                    style={{ background: DASH_GREEN, color: "#fff", border: `1px solid ${LINE}`, position: "sticky", top: 0, insetInlineStart: nameEndOffset, zIndex: 9, width: TOTAL_GRADE_W, minWidth: TOTAL_GRADE_W }}
                  >الدرجة الكلية</th>
                )}
                {columnMeta.map((col, i) => (
                  <th
                    key={col.id}
                    className={`p-1.5 text-center relative ${col.id === animatingColId ? "trash-toss" : ""}`}
                    style={{
                      background: colorLight(col.color),
                      border: `1px solid ${LINE}`,
                      borderBottom: `3px solid ${col.color}`,
                      color: INK,
                      minWidth: col.pinned ? PIN_W : "150px",
                      width: col.pinned ? PIN_W : undefined,
                      position: "sticky",
                      top: 0,
                      insetInlineStart: col.pinned ? col.pinOffset : undefined,
                      zIndex: col.pinned ? 9 : 7,
                    }}
                  >
                    <div className="flex items-center justify-center gap-1.5 mb-1">
                      <span className="w-1.5 h-1.5 rounded-full shrink-0" style={{ background: col.color }} />
                      <span className="font-semibold">{col.name}</span>
                      {col.autoRenew && <RefreshCw size={11} color="#26423B" title="تفريغ تلقائي مفعّل" />}
                    </div>
                    <div className="flex items-center justify-center gap-0.5">
                      <MiniIconBtn icon={ChevronRight} title="نقل لليمين" disabled={examLocked || i === 0} onClick={() => moveColumn(col.id, -1)} />
                      <MiniIconBtn icon={ChevronLeft} title="نقل لليسار" disabled={examLocked || i === columnMeta.length - 1} onClick={() => moveColumn(col.id, 1)} />
                      <MiniIconBtn icon={col.pinned ? Pin : PinOff} title={col.pinned ? "إلغاء التثبيت" : "تثبيت العمود"} color={col.pinned ? "#26423B" : MUTED} disabled={examLocked} onClick={() => togglePinned(col.id)} />
                      <MiniIconBtn icon={Users} title="رصد نفس القيمة لجميع الطلاب" onClick={() => setBulkSetColId(bulkSetColId === col.id ? null : col.id)} />
                      <MiniIconBtn icon={Pencil} title="تعديل العمود" disabled={examLocked} onClick={() => setColModal({ mode: "edit", data: col })} />
                    </div>
                    {bulkSetColId === col.id && (
                      <BulkSetPopover column={col} onApply={(v) => bulkSetColumnValue(col, v)} onClose={() => setBulkSetColId(null)} />
                    )}
                  </th>
                ))}
                <th className="p-1.5 text-center" style={{ background: "#FBEDEA", border: `1px solid ${LINE}`, color: "#9A3B2E", width: 60, minWidth: 60, position: "sticky", top: 0, insetInlineEnd: 0, zIndex: 9 }}>الغياب</th>
              </tr>
            </thead>
            <tbody>
              {showStatsRow && (
                <tr style={{ background: GOLD_LIGHT }}>
                  {cls.showRowNumbers && <td style={{ border: `1px solid ${LINE}` }} />}
                  <td className="p-2 text-xs font-bold" style={{ border: `1px solid ${LINE}`, color: DASH_GREEN }}>📊 المعدل</td>
                  {hasTotalGradeCol && <td style={{ border: `1px solid ${LINE}` }} />}
                  {columnMeta.map((col) => {
                    if (col.type !== "counter") return <td key={col.id} style={{ border: `1px solid ${LINE}` }} />;
                    const validNums = cls.rows.map((r) => cls.cells[`${r.id}:${col.id}`]).filter((v) => v !== undefined && v !== "" && !Number.isNaN(Number(v))).map(Number);
                    const avg = validNums.length ? validNums.reduce((a, b) => a + b, 0) / validNums.length : null;
                    return (
                      <td key={col.id} className="p-2 text-center text-xs font-bold" style={{ border: `1px solid ${LINE}`, color: DASH_GREEN }}>
                        {avg === null ? "—" : Math.round(avg * 100) / 100}
                      </td>
                    );
                  })}
                  <td style={{ border: `1px solid ${LINE}` }} />
                </tr>
              )}
              {visibleRows.map((row, vIdx) => {
                const i = cls.rows.findIndex((r) => r.id === row.id);
                const ruleColor = rowColorRuleMatch(cls, row);
                const nextVisibleRowId = visibleRows[vIdx + 1]?.id || null;
                return (
                  <tr key={row.id} className={`data-row ${row.id === animatingRowId ? "trash-toss" : ""}`}>
                    {cls.showRowNumbers && (
                      <td
                        className="p-2 text-center text-xs font-semibold"
                        style={{ border: `1px solid ${LINE}`, color: MUTED, background: ruleColor || "#fff", position: "sticky", insetInlineStart: 0, zIndex: 3, width: NUM_W, minWidth: NUM_W }}
                      >{i + 1}</td>
                    )}
                    <td
                      className={`${namePadClass} font-medium`}
                      style={{ border: `1px solid ${LINE}`, color: INK, borderInlineStart: `4px solid ${row.color}`, background: ruleColor || "#fff", position: "sticky", insetInlineStart: cls.showRowNumbers ? NUM_W : 0, zIndex: 3, width: NAME_W, minWidth: NAME_W, fontSize: compactDensity ? 12 : undefined }}
                    >
                      <div className="flex items-center gap-1.5 mb-1">
                        <input
                          type="checkbox"
                          checked={selectedRowIds.includes(row.id)}
                          onChange={() => toggleRowSelection(row.id)}
                          className="shrink-0"
                        />
                        <span className="font-semibold truncate flex-1">{row.name}</span>
                        {(() => {
                          const pct = attendancePercent(cls, row.id);
                          if (pct === null) return null;
                          const color = pct >= 90 ? "#26423B" : pct >= 75 ? "#C97A2B" : "#C0392B";
                          return (
                            <span title="نسبة الحضور" className="shrink-0 px-1.5 py-0.5 rounded-full text-[10px] font-bold" style={{ background: `${color}18`, color }}>
                              {pct}٪
                            </span>
                          );
                        })()}
                        {row.medicalNote && row.medicalNote.trim() && (
                          <span title={`تنبيه: ${row.medicalNote}`} className="shrink-0"><AlertTriangle size={12} color="#C97A2B" /></span>
                        )}
                      </div>
                      <div className="flex items-center gap-0.5" style={{ paddingInlineStart: 20 }}>
                        <MiniIconBtn icon={ChevronUp} title="نقل لأعلى" disabled={examLocked || i === 0} onClick={() => moveRow(row.id, -1)} />
                        <MiniIconBtn icon={ChevronDown} title="نقل لأسفل" disabled={examLocked || i === cls.rows.length - 1} onClick={() => moveRow(row.id, 1)} />
                        <span className="mx-1" style={{ width: 1, height: 14, background: LINE, display: "inline-block" }} />
                        <MiniIconBtn icon={Copy} title="نسخ الصف" disabled={examLocked} onClick={() => duplicateRowById(row.id)} />
                        <MiniIconBtn icon={Pencil} title="تعديل الصف" disabled={examLocked} onClick={() => setRowModal({ mode: "edit", data: row })} />
                      </div>
                    </td>
                    {hasTotalGradeCol && (() => {
                      const tg = totalGrade(cls, row.id);
                      return (
                        <td
                          className="p-1.5 text-center"
                          style={{ border: `1px solid ${LINE}`, background: tg ? tg.bandBg : "#fff", position: "sticky", insetInlineStart: nameEndOffset, zIndex: 2 }}
                        >
                          {tg ? (
                            <div>
                              <p className="font-bold text-sm" style={{ color: tg.bandColor }}>{tg.score}/{tg.max}</p>
                              <p className="text-[11px] font-semibold" style={{ color: tg.bandColor }}>{tg.band}</p>
                            </div>
                          ) : (
                            <span className="text-xs" style={{ color: MUTED }}>—</span>
                          )}
                        </td>
                      );
                    })()}
                    {columnMeta.map((col) => {
                      const cellVal = cls.cells[`${row.id}:${col.id}`];
                      const bandColor = col.type === "counter" && col.colorScale ? matchColorBand(col.colorBands, cellVal) : null;
                      const cellBg = bandColor || ruleColor || (col.pinned ? "#FBFAF6" : "#fff");
                      return (
                      <td
                        key={col.id}
                        className={`${cellPadClass} text-center ${col.id === animatingColId ? "trash-toss" : ""}`}
                        style={{
                          border: `1px solid ${LINE}`,
                          background: cellBg,
                          position: col.pinned ? "sticky" : "static",
                          insetInlineStart: col.pinned ? col.pinOffset : undefined,
                          zIndex: col.pinned ? 2 : 0,
                        }}
                      >
                        <Cell column={col} value={cellVal} onChange={(v) => setCell(row, col, v)} rowId={row.id} nextRowId={nextVisibleRowId} />
                      </td>
                      );
                    })}
                    <td className="p-1 text-center" style={{ border: `1px solid ${LINE}`, background: ruleColor || "#fff", position: "sticky", insetInlineEnd: 0, zIndex: 2 }}>
                      {blinkRowId === row.id ? (
                        <span className="inline-block w-10 h-4" />
                      ) : (
                        <button onClick={() => markAbsentToday(row.id)} className="text-xs font-bold hover:opacity-70" style={{ color: "#C0392B" }}>غياب</button>
                      )}
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
          isOwner={isOwner}
        />
      )}
      {showBoard && <DisplayBoard cls={cls} onClose={() => setShowBoard(false)} onPrint={(dateKey) => openPrintPreview({ type: "class", cls, dateKey })} />}
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
          onBack={() => { setReportRowId(null); setShowReportPicker(true); }}
          onEditEntry={(entryId, value) => updateReportEntryValue(reportRow.id, entryId, value)}
          onDeleteEntry={(entryId) => removeReportEntry(reportRow.id, entryId)}
          onDeleteCategory={(colId) => removeReportCategory(reportRow.id, colId)}
          onDeleteAllEntries={() => removeAllReportEntries(reportRow.id)}
          onAddNote={(entry) => addManualReportNote(reportRow.id, entry)}
          onRestoreLatest={() => restoreLatestReportEntry(reportRow.id)}
          onRestoreEntry={(trashId) => restoreReportEntryFromTrash(reportRow.id, trashId)}
          onClearTrash={() => clearReportTrash(reportRow.id)}
          onPrint={() => openPrintPreview({ type: "report", cls, row: reportRow, entries: reportEntries })}
          onPrintParent={() => openPrintPreview({ type: "parentReport", cls, row: reportRow, entries: reportEntries, meta: { schoolName, teacherName: cls.teacher } })}
          onSaveShareId={(rowId, shareId) => updateClass((c) => ({ ...c, rows: c.rows.map((r) => (r.id === rowId ? { ...r, shareId } : r)) }))}
          shawahed={shawahed || {}}
          onLinkShawahed={({ catKey, mode, existingId, title, notes, photo }) => {
            const entries = shawahed?.entries || {};
            if (mode === "existing" && existingId) {
              onLinkShawahed({
                ...shawahed,
                entries: {
                  ...entries,
                  [catKey]: (entries[catKey] || []).map((e) => (e.id === existingId ? { ...e, photo } : e)),
                },
              });
            } else {
              const newEntry = { id: uid(), title, notes, photo, date: todayKey() };
              onLinkShawahed({
                ...shawahed,
                entries: { ...entries, [catKey]: [...(entries[catKey] || []), newEntry] },
              });
            }
          }}
        />
      )}
      {confirmAction && (
        <ConfirmDialog
          title={
            confirmAction.type === "deleteAll" ? "حذف كل بيانات الفصل"
            : confirmAction.type === "clearStudents" ? "تفريغ جميع الطلاب"
            : confirmAction.type === "deleteColumn" ? "حذف العمود"
            : "حذف الصف"
          }
          message={
            confirmAction.type === "deleteAll"
              ? "سيتم حذف جميع الأعمدة والصفوف وبيانات الرصد في هذا الفصل. متابعة؟"
              : confirmAction.type === "clearStudents"
              ? `سيتم حذف جميع الطلاب الحاليين (${cls.rows.length}) وكل رصدهم، بينما تبقى الأعمدة كما هي بدون تغيير — مناسب لو تبي تعيد استخدام نفس الفصل بطلاب جدد. يمكن استعادة الطلاب المحذوفين لاحقًا من سلة المحذوفات. متابعة؟`
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
      {previewJob && (
        <PrintPreviewModal
          job={previewJob.job}
          format={previewJob.format}
          onClose={() => setPreviewJob(null)}
          onExport={(key) => {
            if (key === "pdf") exportPdfShare(previewJob.job);
            else if (key === "png") exportPng(previewJob.job);
            else if (key === "excel") exportExcel(previewJob.job);
            setPreviewJob(null);
          }}
        />
      )}
      {showAttendance && (
        <AttendanceModal
          cls={cls}
          updateClass={updateClass}
          onClose={() => setShowAttendance(false)}
          onPrint={(dateKey) => openPrintPreview({ type: "attendance", cls, dateKey })}
          onShare={(dateKey) => exportPdfShare({ type: "attendance", cls, dateKey })}
        />
      )}
      {showReminders && (
        <RemindersModal
          reminders={cls.reminders}
          onAdd={addReminder}
          onDelete={deleteReminder}
          notifPermission={notifPermission}
          onRequestPermission={requestNotifPermission}
          onClose={() => setShowReminders(false)}
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
      {showReportPicker && (
        <StudentPickerModal
          rows={cls.rows}
          onSelect={(rowId) => { setReportRowId(rowId); setShowReportPicker(false); }}
          onClose={() => setShowReportPicker(false)}
        />
      )}
      {showGradeSheet && (
        <GradeSheetModal
          cls={cls}
          onClose={() => setShowGradeSheet(false)}
          onGenerate={({ shortTestIds, finalExamIds, reviewerName }) => {
            setShowGradeSheet(false);
            openPrintPreview({ type: "gradeSheet", cls, shortTestIds, finalExamIds, reviewerName }, "pdf");
          }}
        />
      )}
      {showPeriodComparison && (
        <PeriodComparisonModal
          cls={cls}
          onClose={() => setShowPeriodComparison(false)}
          onPrint={(data) => openPrintPreview({ type: "periodComparison", cls, ...data })}
        />
      )}
      {showExamMode && (
        <ExamModeModal
          onClose={() => setShowExamMode(false)}
          onActivate={(minutes) => {
            const until = new Date(Date.now() + minutes * 60000).toISOString();
            updateClass((c) => ({ ...c, examModeUntil: until }));
            setShowExamMode(false);
          }}
        />
      )}
      {showRowColorRule && (
        <RowColorRuleModal
          cls={cls}
          onClose={() => setShowRowColorRule(false)}
          onSave={(rule) => updateClass((c) => ({ ...c, rowColorRule: rule }))}
          onClear={() => updateClass((c) => ({ ...c, rowColorRule: null }))}
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

function AuthScreen({ siteSettings }) {
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
    <div className="min-h-screen flex items-center justify-center px-4" style={{ background: "linear-gradient(180deg, #F3F1E9 0%, #FAF8F3 100%)", fontFamily: "'IBM Plex Sans Arabic', sans-serif" }} dir="rtl">
      <div className="w-full max-w-sm rounded-2xl p-6 modal-panel-in" style={{ background: "#fff", border: `1px solid ${LINE}`, boxShadow: "0 10px 40px rgba(35,38,34,0.08)" }}>
        {siteSettings?.siteLogo ? (
          <img src={siteSettings.siteLogo} alt="فصولي" className="max-h-28 mx-auto mb-4 object-contain" />
        ) : (
          <>
            <div className="w-14 h-14 rounded-2xl mx-auto mb-4 flex items-center justify-center" style={{ background: "#26423B" }}>
              <BookOpen size={26} color="#fff" strokeWidth={2} />
            </div>
            <h1 className="text-2xl font-extrabold text-center" style={{ color: INK, fontFamily: "'Cairo', sans-serif" }}>فصولي</h1>
            <p className="text-xs text-center mb-1 tracking-wide" style={{ color: MUTED }}>FOSOOLI</p>
          </>
        )}
        <p className="text-sm text-center mb-6" style={{ color: MUTED }}>{mode === "login" ? "سجّل دخولك للمتابعة" : "أنشئ حسابًا جديدًا"}</p>

        <Field label="البريد الإلكتروني">
          <input type="email" style={inputStyle} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@example.com" autoComplete="email" />
        </Field>
        <Field label="كلمة المرور">
          <input type="password" style={inputStyle} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" autoComplete={mode === "login" ? "current-password" : "new-password"} onKeyDown={(e) => { if (e.key === "Enter") submit(); }} />
        </Field>

        {error && <p className="text-xs mb-3" style={{ color: "#C0392B" }}>{error}</p>}
        {message && <p className="text-xs mb-3" style={{ color: "#26423B" }}>{message}</p>}

        <button disabled={loading || !email.trim() || !password} onClick={submit} className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 active:scale-95 transition-transform" style={{ background: "#26423B" }}>
          {loading ? "جارٍ..." : mode === "login" ? "تسجيل الدخول" : "إنشاء حساب"}
        </button>

        <div className="flex items-center gap-2 my-4">
          <div className="flex-1 h-px" style={{ background: LINE }} />
          <span className="text-xs" style={{ color: MUTED }}>أو</span>
          <div className="flex-1 h-px" style={{ background: LINE }} />
        </div>

        <button onClick={signInWithGoogle} className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold active:scale-95 transition-transform" style={{ border: `1px solid ${LINE}`, color: INK }}>
          <svg width="18" height="18" viewBox="0 0 48 48"><path fill="#FFC107" d="M43.6 20.5H42V20H24v8h11.3c-1.6 4.7-6.1 8-11.3 8-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.1 8 3l5.7-5.7C34.6 6 29.6 4 24 4 12.9 4 4 12.9 4 24s8.9 20 20 20 20-8.9 20-20c0-1.3-.1-2.7-.4-3.5z"/><path fill="#FF3D00" d="M6.3 14.7l6.6 4.8C14.6 16 18.9 13 24 13c3.1 0 5.8 1.1 8 3l5.7-5.7C34.6 6 29.6 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.5 0 10.4-1.9 14.3-5.1l-6.6-5.6C29.6 35 26.9 36 24 36c-5.2 0-9.6-3.3-11.3-7.9l-6.5 5C9.6 39.6 16.3 44 24 44z"/><path fill="#1976D2" d="M43.6 20.5H42V20H24v8h11.3c-.8 2.3-2.2 4.3-4.1 5.7l6.6 5.6C41.4 36 44 30.6 44 24c0-1.3-.1-2.7-.4-3.5z"/></svg>
          الدخول عبر Google
        </button>

        <button onClick={() => { setMode((m) => (m === "login" ? "signup" : "login")); setError(""); setMessage(""); }} className="w-full text-center text-xs font-semibold mt-4" style={{ color: "#26423B" }}>
          {mode === "login" ? "ما عندك حساب؟ أنشئ واحدًا" : "عندك حساب؟ سجّل دخولك"}
        </button>
      </div>
    </div>
  );
}

// ---------- App root ----------

export class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null, info: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, info) {
    console.error("حدث خطأ غير متوقع في الواجهة:", error, info);
    this.setState({ info });
  }
  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center px-4" style={{ background: PAPER, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }} dir="rtl">
          <div className="w-full max-w-lg rounded-2xl p-6 text-center" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <p className="text-sm font-bold mb-2" style={{ color: "#C0392B" }}>حدث خطأ غير متوقع</p>
            <p className="text-xs mb-4" style={{ color: MUTED }}>حاول إعادة تحميل الصفحة. لو تكرر الخطأ، خذ لقطة شاشة من التفاصيل بالأسفل وأرسلها.</p>
            <button onClick={() => window.location.reload()} className="px-5 py-2 rounded-lg text-sm font-bold text-white mb-4" style={{ background: "#26423B" }}>
              إعادة تحميل الصفحة
            </button>
            <div className="text-start p-3 rounded-lg overflow-auto" style={{ background: "#F8F7F2", border: `1px solid ${LINE}`, maxHeight: 260 }}>
              <p className="text-xs font-mono" style={{ color: "#C0392B", direction: "ltr", textAlign: "left" }}>{String(this.state.error?.message || this.state.error || "")}</p>
              {this.state.error?.stack && (
                <pre className="text-[10px] mt-2 whitespace-pre-wrap" style={{ color: MUTED, direction: "ltr", textAlign: "left" }}>{this.state.error.stack}</pre>
              )}
              {this.state.info?.componentStack && (
                <pre className="text-[10px] mt-2 whitespace-pre-wrap" style={{ color: MUTED, direction: "ltr", textAlign: "left" }}>{this.state.info.componentStack}</pre>
              )}
            </div>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  // توجيه بسيط: لو الرابط فيه ?shared=معرّف، نعرض تقرير طالب عام للقراءة
  // فقط (بدون تسجيل دخول) — هذا هو الرابط اللي يُشفَّر برمز QR بتقرير
  // الطالب. أي رابط عادي (بدون هذي الوسيطة) يذهب للتطبيق الكامل كالمعتاد.
  const sharedId = (() => {
    try { return new URLSearchParams(window.location.search).get("shared"); } catch (e) { return null; }
  })();
  if (sharedId) return <SharedReportView shareId={sharedId} />;
  return <AuthenticatedApp />;
}

// عرض عام للقراءة فقط لتقرير طالب واحد — يُفتح من رابط/رمز QR بدون تسجيل
// دخول. يستدعي دالة قاعدة بيانات (RPC) بصلاحيات مقيّدة تُرجع فقط بيانات
// هذا الطالب تحديدًا، مو كامل حساب المعلم — راجع ملاحظة الإعداد بالأسفل.
function SharedReportView({ shareId }) {
  const [state, setState] = useState("loading"); // loading | ready | error
  const [payload, setPayload] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const { data, error } = await supabase.rpc("get_shared_report", { p_share_id: shareId });
        if (error || !data) { setState("error"); return; }
        setPayload(data);
        setState("ready");
      } catch (e) {
        setState("error");
      }
    })();
  }, [shareId]);

  if (state === "loading") {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: PAPER }}>
        <RefreshCw size={22} color={MUTED} className="animate-spin" />
      </div>
    );
  }
  if (state === "error" || !payload) {
    return (
      <div className="min-h-screen flex items-center justify-center px-4" style={{ background: PAPER, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }} dir="rtl">
        <div className="text-center">
          <p className="text-sm font-bold mb-1" style={{ color: "#C0392B" }}>تعذّر عرض هذا التقرير</p>
          <p className="text-xs" style={{ color: MUTED }}>الرابط قد يكون غير صحيح، أو أُلغي من طرف المعلم.</p>
        </div>
      </div>
    );
  }

  const groups = groupEntries(payload.entries || []);

  return (
    <div dir="rtl" style={{ background: PAPER, minHeight: "100vh", fontFamily: "'IBM Plex Sans Arabic', sans-serif" }}>
      <div className="max-w-lg mx-auto px-4 py-8">
        <div className="text-center mb-6">
          <div className="w-14 h-14 rounded-2xl mx-auto mb-3 flex items-center justify-center" style={{ background: "#26423B" }}>
            <BookOpen size={26} color="#fff" />
          </div>
          <h1 className="text-xl font-extrabold" style={{ color: INK }}>{payload.studentName}</h1>
          <p className="text-sm mt-1" style={{ color: MUTED }}>{[payload.subject, payload.grade, payload.teacher].filter(Boolean).join(" • ")}</p>
        </div>
        {groups.length === 0 ? (
          <p className="text-sm text-center py-10" style={{ color: MUTED }}>لا يوجد رصد لهذا الطالب بعد.</p>
        ) : (
          <div className="space-y-4">
            {groups.map((g) => (
              <div key={g.colId} className="rounded-xl overflow-hidden" style={{ border: `1px solid ${LINE}`, background: "#fff" }}>
                <div className="px-4 py-2.5 flex items-center justify-between" style={{ background: g.colColor }}>
                  <span className="font-bold text-sm" style={{ color: "#fff" }}>{g.colName}</span>
                  <span className="text-xs px-2 py-0.5 rounded-full font-semibold" style={{ background: "rgba(255,255,255,0.25)", color: "#fff" }}>{g.items.length}</span>
                </div>
                <div>
                  {g.items.map((it, i) => (
                    <div key={it.id} className="flex items-center justify-between px-4 py-2" style={{ borderTop: i > 0 ? `1px solid ${LINE}` : "none" }}>
                      <span className="text-xs" style={{ color: MUTED }}>{it.day ? `${it.day}، ` : ""}{it.date}</span>
                      <span className="text-sm font-medium" style={{ color: INK }}>{it.value}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
        <p className="text-center text-xs mt-8" style={{ color: MUTED }}>يتحدّث هذا الرابط تلقائيًا مع أي رصد جديد • عبر تطبيق فصولي</p>
      </div>
    </div>
  );
}

function AuthenticatedApp() {
  useFonts();
  const [session, setSession] = useState(undefined); // undefined = still checking, null = logged out
  const [data, setData] = useState({ classes: [], trash: [], schedule: {}, scheduleImage: null, settings: { feedback: true } });
  const [loaded, setLoaded] = useState(false);
  const [view, setViewRaw] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem("fosooli-last-view") || "null");
      if (saved && saved.page === "class" && saved.id) return saved;
    } catch (e) { /* ignore */ }
    return { page: "home" };
  });
  const setView = (next) => {
    setViewRaw(next);
    try { localStorage.setItem("fosooli-last-view", JSON.stringify(next)); } catch (e) { /* ignore */ }
  };
  const [printJob, requestPrint] = usePrint();
  const [siteSettings, setSiteSettings] = useState({});
  const [isOwner, setIsOwner] = useState(false);
  const [disabledMessage, setDisabledMessage] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => setSession(session));
    return () => listener.subscription.unsubscribe();
  }, []);

  // يطبّق لون الهوية الشخصية للمعلم (لو حدده بالإعدادات) على كامل التطبيق
  // كل ما تغيّرت القيمة — أو يرجّع للون الافتراضي لو ما حدد شيء.
  useEffect(() => {
    applyThemeColor(data.settings?.themeColor);
  }, [data.settings?.themeColor]);

  // يكبّر/يصغّر حجم الخط بكامل التطبيق عبر تعديل حجم خط عنصر <html> —
  // معظم أحجام Tailwind مبنية على rem، فتتناسب تلقائيًا مع أي تغيير هنا.
  useEffect(() => {
    const scale = Number(data.settings?.fontScale) || 1;
    document.documentElement.style.fontSize = `${16 * scale}px`;
    return () => { document.documentElement.style.fontSize = ""; };
  }, [data.settings?.fontScale]);

  // إعدادات الموقع العامة (تذييل الصفحة، الشعار، شهادات الثقة): تُحمَّل حتى
  // قبل تسجيل الدخول (تظهر بشاشة الدخول نفسها)، ويقرأها الجميع، لكن فقط
  // المالك (is_owner) يقدر يكتبها فعليًا — القاعدة الأمنية بجهة الخادم.
  const loadSiteSettings = async () => {
    try {
      const { data: row } = await supabase.from("site_settings").select("data").eq("id", 1).maybeSingle();
      if (row && row.data) setSiteSettings(row.data);
    } catch (e) {
      console.error("تعذر تحميل إعدادات الموقع", e);
    }
  };
  useEffect(() => { loadSiteSettings(); }, []);

  useEffect(() => {
    if (!session) return;
    (async () => {
      try {
        const { data: profile } = await supabase.from("profiles").select("is_owner, is_disabled").eq("id", session.user.id).maybeSingle();
        if (profile?.is_disabled) {
          setDisabledMessage(true);
          try { localStorage.removeItem("fosooli-last-view"); } catch (e2) { /* ignore */ }
          await supabase.auth.signOut();
          return;
        }
        setIsOwner(!!profile?.is_owner);
      } catch (e) {
        console.error("تعذر التحقق من صلاحية المالك", e);
      }
    })();
  }, [session]);

  const updateSiteSettings = (fn) => {
    setSiteSettings((prev) => {
      const next = fn(prev);
      supabase.from("site_settings").upsert({ id: 1, data: next, updated_at: new Date().toISOString() }, { onConflict: "id" })
        .then(({ error }) => { if (error) console.error("تعذر حفظ إعدادات الموقع (يتطلب صلاحية المالك)", error); });
      return next;
    });
  };

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
    setSyncStatus("saving");
    const t = setTimeout(() => {
      supabase.from("user_data")
        .upsert({ user_id: session.user.id, data, updated_at: new Date().toISOString() }, { onConflict: "user_id" })
        .then(({ error }) => {
          if (error) { console.error("تعذر حفظ البيانات", error); setSyncStatus("error"); }
          else setSyncStatus("saved");
        });
    }, 600);
    return () => clearTimeout(t);
  }, [data, loaded, session]);

  // مؤشر "متصل / غير متصل": يتابع حالة الشبكة الفعلية + حالة آخر عملية حفظ.
  const [isOnline, setIsOnline] = useState(() => (typeof navigator !== "undefined" ? navigator.onLine : true));
  const [syncStatus, setSyncStatus] = useState("saved"); // saved | saving | error
  useEffect(() => {
    const goOnline = () => setIsOnline(true);
    const goOffline = () => setIsOnline(false);
    window.addEventListener("online", goOnline);
    window.addEventListener("offline", goOffline);
    return () => { window.removeEventListener("online", goOnline); window.removeEventListener("offline", goOffline); };
  }, []);

  const handleSignOut = async () => {
    try { localStorage.removeItem("fosooli-last-view"); } catch (e) { /* ignore */ }
    await supabase.auth.signOut();
  };

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
    if (disabledMessage) {
      return (
        <div className="min-h-screen flex items-center justify-center px-4" style={{ background: PAPER, fontFamily: "'IBM Plex Sans Arabic', sans-serif" }} dir="rtl">
          <div className="w-full max-w-sm rounded-2xl p-6 text-center" style={{ background: "#fff", border: `1px solid ${LINE}` }}>
            <p className="text-sm font-bold mb-2" style={{ color: "#C0392B" }}>هذا الحساب معطّل حاليًا</p>
            <p className="text-xs" style={{ color: MUTED }}>تواصل مع إدارة الموقع لمزيد من المعلومات.</p>
          </div>
        </div>
      );
    }
    return <AuthScreen siteSettings={siteSettings} />;
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
      {view.page === "home" && <HomePage data={data} setData={setData} onOpen={openClass} userEmail={session.user.email} userId={session.user.id} onSignOut={handleSignOut} siteSettings={siteSettings} updateSiteSettings={updateSiteSettings} isOwner={isOwner} isOnline={isOnline} syncStatus={syncStatus} />}
      {view.page === "class" && currentClass && <ClassPage cls={currentClass} updateClass={updateClass} onBack={backHome} requestPrint={requestPrint} feedbackEnabled={data.settings?.feedback !== false} schoolName={data.settings?.schoolName} principalName={data.settings?.principalName} countryName={data.settings?.countryName} ministryName={data.settings?.ministryName} logoImage={data.settings?.logoImage} allClasses={data.classes} onMoveRowsToClass={moveRowsToClass} isOwner={isOwner} density={data.settings?.density} isOnline={isOnline} syncStatus={syncStatus} shawahed={data.shawahed || {}} onLinkShawahed={(next) => setData((d) => ({ ...d, shawahed: next }))} />}
      {view.page === "class" && !currentClass && (
        <div className="max-w-md mx-auto py-20 text-center">
          <p style={{ color: MUTED }}>لم يتم العثور على هذا الفصل</p>
          <button onClick={backHome} className="mt-3 text-sm font-semibold" style={{ color: "#26423B" }}>العودة للرئيسية</button>
        </div>
      )}
      <div className="app-print-root" dir="rtl">
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
    <div
      dir="rtl"
      className="min-h-screen relative"
      style={{
        background: "#FDFCF9",
        fontFamily: "'IBM Plex Sans Arabic', sans-serif",
        borderTop: `4px solid ${hexToRgba(DASH_GREEN, 0.85)}`,
        boxShadow: `inset 6px 0 0 0 ${hexToRgba(DASH_GREEN, 0.12)}, inset -6px 0 0 0 ${hexToRgba(GOLD, 0.12)}`,
      }}
    >
      <div className="fixed inset-0 overflow-hidden pointer-events-none" style={{ zIndex: 0 }}>
        <div className="bg-blob-1" style={{ position: "absolute", top: "-8%", insetInlineEnd: "-6%", width: "50vw", height: "50vw", maxWidth: 680, maxHeight: 680, borderRadius: "9999px", background: `radial-gradient(circle, ${hexToRgba(DASH_GREEN, 0.26)} 0%, ${hexToRgba(DASH_GREEN, 0.14)} 45%, ${hexToRgba(DASH_GREEN, 0)} 72%)`, filter: "blur(4px)" }} />
        <div className="bg-blob-2" style={{ position: "absolute", bottom: "-10%", insetInlineStart: "-8%", width: "48vw", height: "48vw", maxWidth: 660, maxHeight: 660, borderRadius: "9999px", background: `radial-gradient(circle, ${hexToRgba(GOLD, 0.22)} 0%, ${hexToRgba(GOLD, 0)} 72%)`, filter: "blur(4px)" }} />
        <div className="bg-blob-1" style={{ position: "absolute", top: "18%", insetInlineStart: "22%", width: "22vw", height: "22vw", maxWidth: 300, maxHeight: 300, borderRadius: "9999px", background: `radial-gradient(circle, ${hexToRgba(DASH_GREEN, 0.14)} 0%, ${hexToRgba(DASH_GREEN, 0)} 72%)`, filter: "blur(4px)" }} />
      </div>
      <div className="relative" style={{ zIndex: 1 }}>
        {appContent}
      </div>
    </div>
  );
}
