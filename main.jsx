/* ============================================================================
   PATCH FILE — not a standalone app. Copy each block into your existing
   single-file app at the location described in the comment above it.
   ============================================================================ */


/* ----------------------------------------------------------------------------
   1) NEW COLOR CONSTANTS
   WHERE: add right after your existing line:
     const MUTED = "#7A7768";
   (keep INK/PAPER/LINE/MUTED as-is — everything else in the app still uses
   them, so nothing breaks. These are ADDITIONAL constants for the new
   dashboard chrome — sidebar, stat cards, report banners.)
---------------------------------------------------------------------------- */
const DASH_GREEN = "#26423B";       // dark green sidebar / banners
const DASH_GREEN_DARK = "#1B322C";  // deeper shade for gradients/hover
const GOLD = "#D9A441";             // gold accent (buttons, active nav, numbers)
const GOLD_LIGHT = "#F7EBD2";       // light gold background tint
const DASH_CARD_BG = "#FFFFFF";


/* ----------------------------------------------------------------------------
   2) SIDEBAR COMPONENT (new)
   WHERE: add this as a new top-level component, e.g. right before
     function HomePage(...) { ... }
---------------------------------------------------------------------------- */
function AppSidebar({ userEmail, mainTab, setMainTab, onOpenGuide, onOpenSettings, isOwner, onOpenAdmin, siteSettings }) {
  const NAV = [
    { key: "classes", label: "الفصول", icon: BookOpen },
    { key: "shawahed", label: "شواهد", icon: FileCheck },
    { key: "tests", label: "الاختبارات", icon: ListChecks },
    { key: "archive", label: "المؤرشفة", icon: Archive },
  ];
  const initials = (userEmail || "م").trim().charAt(0).toUpperCase();
  return (
    <div
      className="hidden lg:flex flex-col shrink-0"
      style={{
        width: 250,
        background: `linear-gradient(180deg, ${DASH_GREEN} 0%, ${DASH_GREEN_DARK} 100%)`,
        minHeight: "100vh",
        position: "sticky",
        top: 0,
        alignSelf: "flex-start",
        padding: "28px 18px",
        color: "#fff",
      }}
    >
      <div className="flex flex-col items-center mb-8 text-center">
        {siteSettings?.siteLogo ? (
          <img src={siteSettings.siteLogo} alt="" className="max-h-14 object-contain mb-3" />
        ) : (
          <div
            className="w-16 h-16 rounded-full flex items-center justify-center mb-3 font-extrabold text-xl"
            style={{ background: GOLD, color: DASH_GREEN, border: "3px solid rgba(255,255,255,0.25)" }}
          >
            {initials}
          </div>
        )}
        <p className="text-sm font-extrabold tracking-wide">{(userEmail || "معلم/ـة").split("@")[0]}</p>
        <p className="text-[11px]" style={{ color: "rgba(255,255,255,0.6)" }}>{userEmail}</p>
      </div>

      <nav className="flex flex-col gap-1.5 mb-8">
        {NAV.map((n) => {
          const active = mainTab === n.key;
          return (
            <button
              key={n.key}
              onClick={() => setMainTab(n.key)}
              className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-semibold transition-all"
              style={active
                ? { background: "rgba(255,255,255,0.14)", color: GOLD }
                : { background: "transparent", color: "rgba(255,255,255,0.75)" }}
            >
              <n.icon size={17} />
              {n.label}
            </button>
          );
        })}
      </nav>

      <div className="flex flex-col gap-1.5 mt-auto pt-4" style={{ borderTop: "1px solid rgba(255,255,255,0.12)" }}>
        <button onClick={onOpenGuide} className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-medium hover:bg-white/5" style={{ color: "rgba(255,255,255,0.72)" }}>
          <HelpCircle size={16} /> كيف أبدأ؟
        </button>
        <button onClick={onOpenSettings} className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-medium hover:bg-white/5" style={{ color: "rgba(255,255,255,0.72)" }}>
          <Settings size={16} /> الإعدادات
        </button>
        {isOwner && (
          <button onClick={onOpenAdmin} className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-bold hover:bg-white/5" style={{ color: GOLD }}>
            <ShieldCheck size={16} /> لوحة التحكم
          </button>
        )}
      </div>
    </div>
  );
}

/* Small stat card used on the dashboard strip, styled like the reference
   image (icon chip + big number + label). Drop-in replacement/companion
   for your existing DashboardStrip cards. */
function DashStatCard({ label, value, icon: Icon, tone = "green" }) {
  const bg = tone === "gold" ? GOLD : DASH_GREEN;
  return (
    <div className="rounded-2xl p-4 flex items-center justify-between" style={{ background: bg, color: "#fff", minHeight: 84 }}>
      <div>
        <p className="text-xs opacity-80 mb-1">{label}</p>
        <p className="text-2xl font-extrabold">{value}</p>
      </div>
      <div className="w-10 h-10 rounded-xl flex items-center justify-center" style={{ background: "rgba(255,255,255,0.18)" }}>
        <Icon size={18} color="#fff" />
      </div>
    </div>
  );
}

/* ----------------------------------------------------------------------------
   3) INTEGRATION — wrapping HomePage/App with the sidebar
   WHERE: in your `HomePage` component, the outer return currently starts
   with:
     return (
       <div className="max-w-5xl xl:max-w-7xl 2xl:max-w-[1600px] mx-auto px-4 md:px-6 py-6 page-fade-in">
   Wrap that whole return value like this instead (keeps everything else
   inside unchanged — sidebar is hidden below the `lg` breakpoint, so mobile
   still gets your existing top tab bar untouched):

   return (
     <div className="flex" style={{ minHeight: "100vh" }}>
       <AppSidebar
         userEmail={userEmail}
         mainTab={mainTab}
         setMainTab={setMainTab}
         onOpenGuide={() => setShowGuide(true)}
         onOpenSettings={() => setShowSettings(true)}
         isOwner={isOwner}
         onOpenAdmin={() => setShowAdminPanel(true)}
         siteSettings={siteSettings}
       />
       <div className="flex-1 max-w-5xl xl:max-w-7xl 2xl:max-w-[1600px] mx-auto px-4 md:px-6 py-6 page-fade-in">
         {/* ...everything that was already inside the old outer <div> stays exactly here... */}
       </div>
     </div>
   );

   Also: since the sidebar now carries navigation + settings + guide + admin
   panel, you can optionally hide the duplicate icons in your existing
   top-right header block (the one with HelpCircle/Settings/ShieldCheck
   buttons) on large screens by adding `lg:hidden` to that container's
   className — purely cosmetic, app works either way if you leave it.
---------------------------------------------------------------------------- */


/* ----------------------------------------------------------------------------
   4) SHAWAHED REPORT — new visual design + category picker
   WHERE: add this new canvas builder near your existing
     async function buildShawahedReportCanvas(shawahed, meta = {}) { ... }
   (keep the old one — nothing calls it once you switch the export button,
   but no harm leaving it in place)
---------------------------------------------------------------------------- */
async function buildShawahedReportCanvasV2(shawahed, selectedKeys, meta = {}) {
  const { countryName, ministryName, schoolName, logoImage, teacherName, principalName, subject, grade } = meta;
  const entries = shawahed.entries || {};
  const cats = SHAWAHED_CATEGORIES.filter((c) => selectedKeys.includes(c.key) && (entries[c.key] || []).length > 0);
  const scale = 3;
  const width = 900;
  const pad = 36;

  let totalEntries = 0;
  cats.forEach((c) => { totalEntries += entries[c.key].length; });

  let logoImageElement = null;
  if (logoImage) {
    try { logoImageElement = await loadImage(logoImage); } catch (e) { logoImageElement = null; }
  }

  const measure = document.createElement("canvas").getContext("2d");
  measure.font = "13px Tahoma, Arial";

  const bannerH = 108;
  const breadcrumbH = 48;
  const statsH = 92;
  const catHeaderH = 40;
  const rowH = 78;
  const catGap = 16;
  const footerH = 130;

  const rowLinesByCat = {};
  cats.forEach((cat) => {
    rowLinesByCat[cat.key] = (entries[cat.key] || []).map((e) =>
      wrapCanvasText(measure, e.notes || "", width - pad * 2 - 200)
    );
  });

  let bodyHeight = 0;
  cats.forEach((cat) => {
    bodyHeight += catHeaderH + (entries[cat.key] || []).length * rowH + catGap;
  });

  const height = bannerH + breadcrumbH + statsH + bodyHeight + footerH + pad;

  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  ctx.direction = "rtl";
  ctx.textBaseline = "middle";

  // ---- top banner ----
  const bannerGrad = ctx.createLinearGradient(0, 0, width, 0);
  bannerGrad.addColorStop(0, DASH_GREEN);
  bannerGrad.addColorStop(1, DASH_GREEN_DARK);
  ctx.fillStyle = bannerGrad;
  ctx.fillRect(0, 0, width, bannerH);

  if (logoImageElement) {
    const s = 58;
    ctx.save();
    ctx.beginPath();
    ctx.roundRect(width - pad - s, (bannerH - s) / 2, s, s, 12);
    ctx.clip();
    ctx.fillStyle = "#fff";
    ctx.fillRect(width - pad - s, (bannerH - s) / 2, s, s);
    ctx.drawImage(logoImageElement, width - pad - s, (bannerH - s) / 2, s, s);
    ctx.restore();
  }

  ctx.textAlign = "center";
  ctx.fillStyle = "#fff";
  ctx.font = "bold 24px Tahoma, Arial";
  ctx.fillText("سجل توثيق شواهد الأداء الوظيفي", width / 2, bannerH / 2 - 12);
  ctx.font = "13px Tahoma, Arial";
  ctx.fillStyle = GOLD_LIGHT;
  const sub = [countryName, ministryName, schoolName].filter(Boolean).join("  •  ");
  ctx.fillText(sub || " ", width / 2, bannerH / 2 + 16);

  // ---- breadcrumb pills ----
  let by = bannerH + breadcrumbH / 2;
  const crumbs = [subject, grade, teacherName].filter(Boolean);
  if (crumbs.length) {
    ctx.font = "12px Tahoma, Arial";
    let cx = width - pad;
    crumbs.forEach((c, i) => {
      const w = ctx.measureText(c).width + 28;
      cx -= w;
      ctx.fillStyle = "#F3F1E9";
      ctx.beginPath();
      ctx.roundRect(cx, by - 14, w, 28, 14);
      ctx.fill();
      ctx.fillStyle = INK;
      ctx.textAlign = "center";
      ctx.fillText(c, cx + w / 2, by + 1);
      cx -= 10;
    });
  }

  // ---- stat mini-cards ----
  let sy = bannerH + breadcrumbH + 10;
  const stats = [
    { label: "إجمالي الشواهد", value: totalEntries },
    { label: "عدد الفئات", value: cats.length },
    { label: "التاريخ", value: formatDateDisplay(todayKey()) },
  ];
  const cardGap = 14;
  const cardW = (width - pad * 2 - cardGap * (stats.length - 1)) / stats.length;
  stats.forEach((s, i) => {
    const x = pad + i * (cardW + cardGap);
    ctx.fillStyle = i === 1 ? GOLD : DASH_GREEN;
    ctx.beginPath();
    ctx.roundRect(x, sy, cardW, statsH - 16, 14);
    ctx.fill();
    ctx.fillStyle = "#fff";
    ctx.textAlign = "center";
    ctx.font = "bold 20px Tahoma, Arial";
    ctx.fillText(String(s.value), x + cardW / 2, sy + statsH / 2 - 22);
    ctx.font = "11px Tahoma, Arial";
    ctx.globalAlpha = 0.85;
    ctx.fillText(s.label, x + cardW / 2, sy + statsH / 2 - 2);
    ctx.globalAlpha = 1;
  });

  // ---- category sections ----
  let y = bannerH + breadcrumbH + statsH;
  for (const cat of cats) {
    const officialIndex = SHAWAHED_CATEGORIES.indexOf(cat) + 1;
    ctx.fillStyle = cat.color;
    ctx.beginPath();
    ctx.roundRect(pad, y, width - pad * 2, catHeaderH, 10);
    ctx.fill();
    ctx.fillStyle = "#fff";
    ctx.font = "bold 14px Tahoma, Arial";
    ctx.textAlign = "right";
    ctx.fillText(`${officialIndex}. ${cat.title}`, width - pad - 14, y + catHeaderH / 2);
    ctx.textAlign = "left";
    ctx.font = "12px Tahoma, Arial";
    ctx.fillText(`${(entries[cat.key] || []).length} شاهد`, pad + 14, y + catHeaderH / 2);
    y += catHeaderH + 4;

    const list = entries[cat.key] || [];
    for (let ri = 0; ri < list.length; ri++) {
      const entry = list[ri];
      const rY = y;
      ctx.fillStyle = ri % 2 ? "#FBFAF6" : "#fff";
      ctx.fillRect(pad, rY, width - pad * 2, rowH - 6);
      ctx.strokeStyle = LINE;
      ctx.lineWidth = 1;
      ctx.strokeRect(pad, rY, width - pad * 2, rowH - 6);

      // checkmark chip
      const chipR = 11;
      ctx.beginPath();
      ctx.arc(width - pad - 20, rY + 24, chipR, 0, Math.PI * 2);
      ctx.fillStyle = cat.color;
      ctx.fill();
      ctx.strokeStyle = "#fff";
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(width - pad - 25, rY + 24);
      ctx.lineTo(width - pad - 21, rY + 28);
      ctx.lineTo(width - pad - 14, rY + 19);
      ctx.stroke();

      let photoImg = null;
      if (entry.photo) { try { photoImg = await loadImage(entry.photo); } catch (e) { photoImg = null; } }
      let leftEdge = pad + 16;
      if (photoImg) {
        const thumb = 44;
        ctx.save();
        ctx.beginPath();
        ctx.roundRect(pad + 10, rY + 12, thumb, thumb, 8);
        ctx.clip();
        ctx.drawImage(photoImg, pad + 10, rY + 12, thumb, thumb);
        ctx.restore();
        leftEdge = pad + 10 + thumb + 12;
      }

      ctx.textAlign = "right";
      ctx.fillStyle = INK;
      ctx.font = "bold 13px Tahoma, Arial";
      ctx.fillText(entry.title, width - pad - 40, rY + 22);

      ctx.font = "11px Tahoma, Arial";
      ctx.fillStyle = MUTED;
      const lines = rowLinesByCat[cat.key][ri];
      if (lines && lines[0]) ctx.fillText(lines[0], width - pad - 40, rY + 42);

      ctx.textAlign = "left";
      ctx.font = "10px Tahoma, Arial";
      ctx.fillStyle = MUTED;
      ctx.fillText(formatDateDisplay(entry.date), leftEdge, rY + rowH - 22);

      y += rowH;
    }
    y += catGap;
  }

  // ---- footer signature cards ----
  y += 10;
  const boxW = (width - pad * 2 - 20) / 2;
  const boxH = footerH - 30;
  ctx.textAlign = "center";
  [
    { label: "توقيع المعلم/ـة", name: teacherName, x: pad },
    { label: "توقيع مدير/ة المدرسة", name: principalName, x: pad + boxW + 20 },
  ].forEach((box) => {
    ctx.fillStyle = GOLD_LIGHT;
    ctx.beginPath();
    ctx.roundRect(box.x, y, boxW, boxH, 14);
    ctx.fill();
    ctx.fillStyle = DASH_GREEN;
    ctx.font = "bold 13px Tahoma, Arial";
    ctx.fillText(box.label, box.x + boxW / 2, y + 26);
    ctx.font = "12px Tahoma, Arial";
    ctx.fillStyle = INK;
    ctx.fillText(box.name && box.name.trim() ? box.name : "....................................", box.x + boxW / 2, y + 52);
    ctx.strokeStyle = "#D8CDA8";
    ctx.beginPath();
    ctx.moveTo(box.x + 20, y + boxH - 18);
    ctx.lineTo(box.x + boxW - 20, y + boxH - 18);
    ctx.stroke();
  });

  return { canvas, logicalWidth: width, logicalHeight: height };
}


/* ----------------------------------------------------------------------------
   5) jobToCanvas — add a new case
   WHERE: inside `async function jobToCanvas(job) { ... }`, right after the
   existing:
     if (job.type === "shawahedReport") {
       return buildShawahedReportCanvas(job.shawahed, job.meta || {});
     }
   add:
---------------------------------------------------------------------------- */
/*
  if (job.type === "shawahedReportV2") {
    return buildShawahedReportCanvasV2(job.shawahed, job.selectedKeys, job.meta || {});
  }
*/

/* ----------------------------------------------------------------------------
   6) jobToTable — add a matching case (used for the Excel export option and
   for filename/title when previewing)
   WHERE: inside `function jobToTable(job) { ... }`, right after the existing
     if (job.type === "shawahedReport") { ... return {...}; }
   add:
---------------------------------------------------------------------------- */
/*
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
*/


/* ----------------------------------------------------------------------------
   7) CATEGORY PICKER MODAL (new)
   WHERE: add as a new component near ShawahedHub.
---------------------------------------------------------------------------- */
function ShawahedExportPickerModal({ shawahed, onClose, onConfirm }) {
  const entries = shawahed.entries || {};
  const available = SHAWAHED_CATEGORIES.filter((c) => (entries[c.key] || []).length > 0);
  const [selected, setSelected] = useState(available.map((c) => c.key)); // default: all

  const toggle = (key) => setSelected((s) => (s.includes(key) ? s.filter((k) => k !== key) : [...s, key]));
  const selectAll = () => setSelected(available.map((c) => c.key));
  const selectNone = () => setSelected([]);

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
          <div className="space-y-1.5 mb-4 max-h-72 overflow-y-auto">
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
          <button
            disabled={selected.length === 0}
            onClick={() => onConfirm(selected)}
            className="w-full py-2.5 rounded-xl text-sm font-bold text-white disabled:opacity-40 transition-all hover:brightness-110"
            style={{ background: `linear-gradient(135deg, ${GOLD}, ${DASH_GREEN})` }}
          >
            معاينة التقرير {selected.length > 0 ? `(${selected.length})` : ""}
          </button>
        </>
      )}
    </Modal>
  );
}


/* ----------------------------------------------------------------------------
   8) ShawahedHub — wire the new picker in, plus a quick per-category print
   WHERE: inside `function ShawahedHub(...)`:

   a) add state near the top of the component:
        const [showExportPicker, setShowExportPicker] = useState(null); // null | { mode: 'export' } — actual pick handled by picker itself

   b) replace this existing button:
        <IconBtn icon={FileText} label="تصدير التقرير" magic onClick={onExport} />
      with:
        <IconBtn icon={FileText} label="طباعة / مشاركة" magic onClick={() => setShowExportPicker(true)} />

   c) add a quick print icon on each category card — inside the `.map(cat => ...)`
      button block, add a second small icon button next to <FileCheck>, e.g.:
        <button
          title={`طباعة "${cat.title}" فقط`}
          onClick={(e) => { e.stopPropagation(); onQuickPrint(cat.key); }}
          className="p-1 rounded-full hover:bg-black/5"
        >
          <Printer size={13} color={cat.color} />
        </button>
      (requires passing a new `onQuickPrint` prop down from HomePage, see below)

   d) near the bottom, alongside the existing `{showArchive && (...)}` block, add:
        {showExportPicker && (
          <ShawahedExportPickerModal
            shawahed={shawahed}
            onClose={() => setShowExportPicker(false)}
            onConfirm={(keys) => { setShowExportPicker(false); onExport(keys); }}
          />
        )}

   e) `ShawahedHub`'s `onExport` prop is currently called with no arguments.
      Update its signature usage in HomePage (step 9 below) so it now expects
      a `selectedKeys` array.
---------------------------------------------------------------------------- */


/* ----------------------------------------------------------------------------
   9) HomePage — build the new job type instead of the old one
   WHERE: inside `function HomePage(...)`, find:
     onExport={() => setShawahedPreview({
       type: "shawahedReport",
       shawahed: data.shawahed || {},
       meta: { ... },
     })}
   replace with:
---------------------------------------------------------------------------- */
/*
  onExport={(selectedKeys) => setShawahedPreview({
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
      subject: "شواهد الأداء الوظيفي",
      grade: "",
    },
  })}
  onQuickPrint={(catKey) => setShawahedPreview({
    type: "shawahedReportV2",
    shawahed: data.shawahed || {},
    selectedKeys: [catKey],
    meta: {
      countryName: data.settings?.countryName,
      ministryName: data.settings?.ministryName,
      schoolName: data.settings?.schoolName,
      logoImage: data.settings?.logoImage,
      teacherName: data.classes[0]?.teacher || "",
      principalName: data.settings?.principalName,
    },
  })}
*/

/* That's it — the existing <PrintPreviewModal job={shawahedPreview} .../>
   block right below it needs no changes: it already calls exportPdfShare /
   exportPng / exportExcel(shawahedPreview), which now route through the
   new job.type === "shawahedReportV2" cases you added in steps 5 & 6. */
