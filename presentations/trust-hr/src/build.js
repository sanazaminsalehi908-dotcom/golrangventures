// Trust event — HR deck: Golrang Innovation Award with HR
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");
const fs = require("fs");
const { applyTheme } = require(process.env.SKILL_DIR + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Golrang Ventures",
  headFontFace: "Vazirmatn",
  bodyFontFace: "Vazirmatn",
  colors: {
    dk1: "16233D", // body text
    lt1: "FFFFFF",
    dk2: "0B2A5B", // navy
    lt2: "EEF3FA", // card tint
    accent1: "1E5FB4", // Golrang Ventures blue
    accent2: "F2B632", // yellow
    accent3: "5B8FD6", // light blue
    accent4: "5E6B80", // muted slate (captions)
    accent5: "2E9C7A", // green
    accent6: "C8553D", // brick (barriers)
    hlink: "1E5FB4",
    folHlink: "6B4FA0",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "نوآوری را منابع انسانی زنده نگه می‌دارد";
pres.author = "گلرنگ ونچرز";
pres.subject = "رویداد تراست";
const C = pres.SchemeColor;

const W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
const FA = "۰۱۲۳۴۵۶۷۸۹";
const fa = (s) => String(s).replace(/[0-9]/g, (d) => FA[d]).replace(/\./g, "٫");

// ---------- icons ----------
async function icon(name, hex, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(lu[name], { color: "#" + hex, size: px, strokeWidth: 1.75 })
  );
  const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- text helper (RTL) ----------
function T(slide, text, o) {
  const opts = Object.assign(
    { isTextBox: true, rtlMode: true, lang: "fa-IR", align: "right", valign: "top", margin: 0, color: C.text1, fontSize: 14, paraSpaceAfter: 0 },
    o
  );
  slide.addText(text, opts);
}
function card(slide, x, y, w, h, fill, name, extra = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name }, extra));
}
async function iconDisc(slide, x, y, d, iconName, discColor, iconHex, name) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: discColor }, line: { type: "none" }, objectName: name + " disc" });
  const p = d * 0.24;
  slide.addImage({ data: await icon(iconName, iconHex), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name + " icon", altText: name });
}

// ---------- layouts ----------
const FOOT = "جایزه نوآوری گلرنگ با همراهی منابع انسانی · رویداد تراست";
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 5.2, y: 1.55, w: 7.53, h: 0.4, fontSize: 14, bold: true, color: C.accent2, align: "right", rtlMode: true, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 5.2, y: 2.05, w: 7.53, h: 2.0, fontSize: 40, bold: true, color: C.background1, align: "right", valign: "top", rtlMode: true, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 5.2, y: 4.2, w: 7.53, h: 0.9, fontSize: 17, color: C.background2, align: "right", valign: "top", rtlMode: true, margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, M, 0.6, M],
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.42, w: CW, h: 0.34, fontSize: 13, bold: true, color: C.accent1, align: "right", rtlMode: true, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.78, w: CW, h: 0.75, fontSize: 30, bold: true, color: C.text2, align: "right", valign: "top", rtlMode: true, margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 5.0, y: 7.0, w: 7.73, h: 0.3, fontSize: 10, color: C.accent4, align: "right", rtlMode: true, margin: 0 } } },
  ],
  slideNumber: { x: M, y: 7.0, w: 0.8, h: 0.3, fontSize: 10, color: C.accent4, align: "left" },
});

const SECTIONS = ["شروع", "اطلاع‌رسانی", "مشارکت", "سامانه", "باشگاه نوآوری", "داشبورد", "جمع‌بندی"];
let curSec = null;
function slideIn(sec, layout = "CONTENT") {
  if (sec !== curSec) { pres.addSection({ title: sec }); curSec = sec; }
  return pres.addSlide({ masterName: layout, sectionTitle: sec });
}
function head(s, kicker, title) {
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
}

(async () => {
  // =========================================================
  // 1. Cover
  {
    const s = slideIn("شروع", "COVER");
    // concentric rings motif on the left
    const cx = 2.45, cy = 3.75;
    [[4.6, C.accent1, 70], [3.4, C.accent1, 45], [2.2, C.accent1, 0]].forEach(([d, col, tr], i) =>
      s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: col, transparency: tr }, line: { type: "none" }, objectName: "ring " + (i + 1) })
    );
    s.addImage({ data: await icon("LuLightbulb", HEX.accent2), x: cx - 0.6, y: cy - 0.6, w: 1.2, h: 1.2, objectName: "cover icon", altText: "لامپ ایده" });
    s.addText("رویداد تراست · ویژه‌ی مدیران منابع انسانی", { placeholder: "kicker" });
    s.addText("نوآوری را منابع انسانی زنده نگه می‌دارد", { placeholder: "title" });
    s.addText("جایزه نوآوری، سامانه، باشگاه و داشبورد: نقش منابع انسانی در هر قدم", { placeholder: "body" });
    T(s, "[نام ارائه‌دهنده] · [سمت] · گلرنگ ونچرز", { x: 5.2, y: 5.55, w: 7.53, h: 0.35, fontSize: 14, color: C.background1, objectName: "presenter" });
    T(s, "آبان ۱۴۰۵", { x: 5.2, y: 5.95, w: 7.53, h: 0.35, fontSize: 13, color: C.accent3, objectName: "date" });
    s.addNotes("خوشامد و معرفی کوتاه (۱ دقیقه). پیام اصلی امروز: جایزه نوآوری بدون همراهی منابع انسانی فقط یک رویداد سالانه است؛ با همراهی شما می‌تواند بخشی از فرهنگ کار در هر شرکت شود. امروز در ۳۰ دقیقه می‌گوییم دقیقاً چه کمکی از شما می‌خواهیم و چه ابزارهایی در اختیارتان می‌گذاریم.");
  }

  // =========================================================
  // 2. Why HR
  {
    const s = slideIn("شروع");
    head(s, "چرا منابع انسانی؟", "جایزه ایده‌ها را پیدا می‌کند؛ شما آدم‌ها را به میدان می‌آورید");
    const items = [
      ["LuUsers", "به همه دسترسی دارید", "منابع انسانی تنها واحدی است که با تک‌تک همکاران، از خط تولید و انبار تا دفتر، در ارتباط است."],
      ["LuTrendingUp", "ابزار انگیزش دست شماست", "ارزیابی عملکرد، پاداش، آموزش و مسیر شغلی؛ همان ابزارهایی که رفتار را واقعاً تغییر می‌دهند."],
      ["LuSprout", "فرهنگ را شما می‌سازید", "نوآوری وقتی ماندگار می‌شود که بخشی از روال کار باشد، نه یک رویداد در سال."],
    ];
    const gap = 0.4, cw = (CW - 2 * gap) / 3, y = 1.95, ch = 3.3;
    for (let i = 0; i < 3; i++) {
      const x = W - M - cw - i * (cw + gap);
      card(s, x, y, cw, ch, C.background2, "why card " + (i + 1));
      await iconDisc(s, x + cw - 0.4 - 0.9, y + 0.4, 0.9, items[i][0], C.accent1, HEX.lt1, "why " + (i + 1));
      T(s, items[i][1], { x: x + 0.4, y: y + 1.55, w: cw - 0.8, h: 0.5, fontSize: 19, bold: true, color: C.text2, objectName: "why title " + (i + 1) });
      T(s, items[i][2], { x: x + 0.4, y: y + 2.1, w: cw - 0.8, h: 1.1, fontSize: 14, color: C.text1, lineSpacingMultiple: 1.15, objectName: "why body " + (i + 1) });
    }
    card(s, M, 5.6, CW, 1.0, C.text2, "message band");
    T(s, [
      { text: "هدف مشترک ما: ", options: { bold: true, color: C.accent2 } },
      { text: "تبدیل نوآوری از یک رویداد سالانه به بخشی از فرهنگ کار در هر شرکت گروه", options: { color: C.background1 } },
    ], { x: M + 0.4, y: 5.6, w: CW - 0.8, h: 1.0, fontSize: 17, valign: "middle", objectName: "message" });
    s.addNotes("(۲ دقیقه) سه دلیل که این کار بدون منابع انسانی پیش نمی‌رود. تأکید: ما ایده‌ها را داوری و حمایت می‌کنیم، اما رساندن پیام به همکاران و ساختن انگیزه کار شماست.");
  }

  // =========================================================
  // 3. Agenda
  {
    const s = slideIn("شروع");
    head(s, "در این ۳۰ دقیقه", "پنج موضوع، یک درخواست");
    const ag = [
      ["LuMegaphone", "اطلاع‌رسانی", "جایزه با صدای منابع انسانی", "۵ دقیقه"],
      ["LuHeartHandshake", "مشارکت", "تشویق همکاران به ایده دادن", "۶ دقیقه"],
      ["LuMonitorSmartphone", "سامانه‌ی جدید", "جایزه در لایه‌ی هر شرکت", "۷ دقیقه"],
      ["LuAward", "باشگاه نوآوری", "مزایا و سهم هر طرف", "۴ دقیقه"],
      ["LuLayoutDashboard", "داشبورد زنده", "همه‌چیز در یک نگاه", "۳ دقیقه"],
    ];
    const n = ag.length, gap = 0.3, cw = (CW - (n - 1) * gap) / n, y = 2.05;
    // connector line behind discs
    s.addShape(pres.shapes.LINE, { x: M + cw / 2, y: y + 0.55, w: CW - cw, h: 0, line: { color: C.accent3, width: 1.5, dashType: "dash" }, objectName: "agenda line" });
    for (let i = 0; i < n; i++) {
      const x = W - M - cw - i * (cw + gap);
      await iconDisc(s, x + cw / 2 - 0.55, y, 1.1, ag[i][0], C.accent1, HEX.lt1, "agenda " + (i + 1));
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 + 0.22, y: y - 0.08, w: 0.42, h: 0.42, fill: { color: C.accent2 }, line: { color: C.background1, width: 2 }, objectName: "agenda num disc " + (i + 1) });
      T(s, fa(i + 1), { x: x + cw / 2 + 0.22, y: y - 0.08, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: C.text2, align: "center", valign: "middle", objectName: "agenda num " + (i + 1) });
      T(s, ag[i][1], { x, y: y + 1.35, w: cw, h: 0.45, fontSize: 19, bold: true, color: C.text2, align: "center", objectName: "agenda title " + (i + 1) });
      T(s, ag[i][2], { x, y: y + 1.85, w: cw, h: 0.7, fontSize: 14, color: C.text1, align: "center", objectName: "agenda desc " + (i + 1) });
      T(s, ag[i][3], { x, y: y + 2.55, w: cw, h: 0.35, fontSize: 12, bold: true, color: C.accent1, align: "center", objectName: "agenda time " + (i + 1) });
    }
    card(s, M, 5.55, CW, 1.05, C.background2, "ask band");
    await iconDisc(s, W - M - 0.3 - 0.65, 5.75, 0.65, "LuListChecks", C.accent2, HEX.dk2, "ask");
    T(s, [
      { text: "و در پایان (۳ دقیقه): ", options: { bold: true, color: C.text2 } },
      { text: "پنج کاری که از منابع انسانی هر شرکت می‌خواهیم، و گام بعدی", options: { color: C.text1 } },
    ], { x: M + 0.4, y: 5.55, w: CW - 1.6, h: 1.05, fontSize: 17, valign: "middle", objectName: "ask text" });
    s.addNotes("(۱ دقیقه) نقشه‌ی ارائه. چون همه با جایزه آشنا هستند، مستقیم سراغ نقش منابع انسانی می‌رویم.");
  }

  // =========================================================
  // 4. Communication channels
  {
    const s = slideIn("اطلاع‌رسانی");
    head(s, "۱ · اطلاع‌رسانی با همراهی منابع انسانی", "پیام جایزه باید از زبان خود شرکت شنیده شود");
    const ch = [
      ["LuMail", "اعلام رسمی مدیرعامل", "ایمیل یا ویدئوی کوتاه؛ نشان می‌دهد این کار مهم است"],
      ["LuMessageCircle", "کانال‌های داخلی", "پیام‌رسان سازمانی، اینترانت و «گلرنگ من»"],
      ["LuSun", "جلسات صبحگاهی", "پنج دقیقه معرفی توسط خود مدیر، نه فقط بخشنامه"],
      ["LuQrCode", "پوستر و QR در محل کار", "سالن غذاخوری، ورودی‌ها و خط تولید"],
      ["LuVideo", "روایت ایده‌پردازان قبلی", "ویدئوی کوتاه از همکاری که ایده‌اش اجرا شده"],
      ["LuGraduationCap", "آموزش و آنبوردینگ", "یک اسلاید ثابت درباره‌ی جایزه در دوره‌ها و معارفه‌ی نیروهای جدید"],
    ];
    const gridW = 8.2, gap = 0.3, cw = (gridW - gap) / 2, chh = 1.4, y0 = 1.85;
    for (let i = 0; i < ch.length; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = W - M - cw - col * (cw + gap), y = y0 + row * (chh + 0.25);
      card(s, x, y, cw, chh, C.background2, "channel " + (i + 1));
      await iconDisc(s, x + cw - 0.3 - 0.7, y + 0.35, 0.7, ch[i][0], C.accent1, HEX.lt1, "channel " + (i + 1));
      T(s, ch[i][1], { x: x + 0.3, y: y + 0.25, w: cw - 1.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, objectName: "channel title " + (i + 1) });
      T(s, ch[i][2], { x: x + 0.3, y: y + 0.66, w: cw - 1.5, h: 0.65, fontSize: 13, color: C.text1, objectName: "channel body " + (i + 1) });
    }
    // ambassador panel (left)
    const px = M, pw = W - M - gridW - 0.35 - M, py = y0, ph = 3 * chh + 2 * 0.25;
    card(s, px, py, pw, ph, C.text2, "ambassador panel");
    await iconDisc(s, px + pw - 0.35 - 0.85, py + 0.4, 0.85, "LuUserCheck", C.accent2, HEX.dk2, "ambassador");
    T(s, "سفیر نوآوری", { x: px + 0.35, y: py + 1.45, w: pw - 0.7, h: 0.5, fontSize: 22, bold: true, color: C.accent2, objectName: "ambassador title" });
    T(s, "یک نفر از منابع انسانی هر شرکت که مالک اطلاع‌رسانی است:", { x: px + 0.35, y: py + 2.0, w: pw - 0.7, h: 0.7, fontSize: 14, color: C.background1, objectName: "ambassador lead" });
    const amb = ["تقویم پیام‌ها را اجرا می‌کند", "به سؤال همکاران جواب می‌دهد", "رابط شرکت با گلرنگ ونچرز است"];
    for (let k = 0; k < amb.length; k++) {
      const yy = py + 2.85 + k * 0.5;
      s.addImage({ data: await icon("LuCircleCheck", HEX.accent2), x: px + pw - 0.35 - 0.28, y: yy + 0.04, w: 0.28, h: 0.28, objectName: "ambassador check " + (k + 1), altText: "تیک" });
      T(s, amb[k], { x: px + 0.35, y: yy, w: pw - 1.1, h: 0.38, fontSize: 14, color: C.background1, valign: "middle", objectName: "ambassador item " + (k + 1) });
    }
    s.addNotes("(۳ دقیقه) شش کانال پیشنهادی. نکته‌ی کلیدی: پیام از زبان مدیرعامل و مدیر مستقیم، اثر چند برابر بخشنامه دارد. پیشنهاد اصلی ما: هر شرکت یک «سفیر نوآوری» از منابع انسانی معرفی کند.");
  }

  // =========================================================
  // 5. Communication calendar
  {
    const s = slideIn("اطلاع‌رسانی");
    head(s, "۱ · اطلاع‌رسانی با همراهی منابع انسانی", "تقویم پیشنهادی اطلاع‌رسانی در هر شرکت");
    const st = [
      ["هفته‌ی ۱", "اعلام رسمی", "پیام مدیرعامل و معرفی سفیر نوآوری شرکت"],
      ["هفته‌ی ۲", "کارگاه «از مسئله تا ایده»", "یک ساعت در هر شرکت؛ پیدا کردن مسئله‌های واقعی کار"],
      ["هفته‌ی ۳ تا ۵", "روایت و یادآوری", "هر هفته یک داستان، یک پرسش یا یک مسئله‌ی باز"],
      ["هفته‌ی ۶", "کلینیک ایده", "کمک حضوری به تکمیل فرم پیش از پایان مهلت"],
      ["پس از مهلت", "بازخورد و تقدیر", "پاسخ به همه‌ی ایده‌پردازان، نه فقط برندگان"],
    ];
    const n = st.length, gap = 0.25, cw = (CW - (n - 1) * gap) / n, y = 2.2, ch = 3.15;
    s.addShape(pres.shapes.LINE, { x: M + cw / 2, y: y + 0.3, w: CW - cw, h: 0, line: { color: C.accent1, width: 2 }, objectName: "timeline line" });
    for (let i = 0; i < n; i++) {
      const x = W - M - cw - i * (cw + gap);
      const last = i === n - 1;
      s.addShape(pres.shapes.OVAL, { x: x + cw / 2 - 0.3, y, w: 0.6, h: 0.6, fill: { color: last ? C.accent2 : C.accent1 }, line: { color: C.background1, width: 3 }, objectName: "step disc " + (i + 1) });
      T(s, fa(i + 1), { x: x + cw / 2 - 0.3, y, w: 0.6, h: 0.6, fontSize: 16, bold: true, color: last ? C.text2 : C.background1, align: "center", valign: "middle", objectName: "step num " + (i + 1) });
      card(s, x, y + 0.9, cw, ch - 0.9, last ? C.text2 : C.background2, "step card " + (i + 1));
      T(s, st[i][0], { x: x + 0.25, y: y + 1.1, w: cw - 0.5, h: 0.35, fontSize: 13, bold: true, color: last ? C.accent2 : C.accent1, objectName: "step when " + (i + 1) });
      T(s, st[i][1], { x: x + 0.25, y: y + 1.5, w: cw - 0.5, h: 0.75, fontSize: 17, bold: true, color: last ? C.background1 : C.text2, objectName: "step title " + (i + 1) });
      T(s, st[i][2], { x: x + 0.25, y: y + 2.3, w: cw - 0.5, h: 0.9, fontSize: 13, color: last ? C.background2 : C.text1, objectName: "step body " + (i + 1) });
    }
    T(s, "هفته‌ها نسبت به تاریخ فراخوان شمرده شده‌اند و با تقویم و تعطیلات هر شرکت تنظیم می‌شوند. کیت آماده‌ی پیام‌ها و پوسترها را گلرنگ ونچرز در اختیار سفیران می‌گذارد.", { x: M, y: 5.75, w: CW, h: 0.7, fontSize: 13, color: C.accent4, objectName: "calendar note" });
    s.addNotes("(۲ دقیقه) یک ریتم ساده‌ی شش‌هفته‌ای. مهم‌ترین قدم آخر است: به همه‌ی ایده‌پردازان پاسخ بدهیم؛ همین، مشارکت دوره‌ی بعد را می‌سازد. محتوای آماده (متن پیام، پوستر، ویدئو) را ما تهیه می‌کنیم.");
  }

  // =========================================================
  // 6. Barriers -> HR response
  {
    const s = slideIn("مشارکت");
    head(s, "۲ · افزایش مشارکت همکاران", "چرا همکاران ایده نمی‌دهند و منابع انسانی چه می‌تواند بکند");
    const rows = [
      ["«وقت ندارم»", "ساعت آزاد نوآوری: چند ساعت در ماه برای کار روی ایده، با هماهنگی مدیر"],
      ["«ایده‌ی من کوچک است»", "ایده‌های بهبود هم پذیرفته و دیده می‌شوند؛ لازم نیست هر ایده یک استارتاپ باشد"],
      ["«مدیرم حمایت نمی‌کند»", "مشارکت تیم در جایزه، یکی از شاخص‌های ارزیابی مدیران شود"],
      ["«فایده‌ای ندارد»", "به هر ایده در زمان مشخص پاسخ و بازخورد داده شود، حتی اگر انتخاب نشود"],
    ];
    const y0 = 2.0, rh = 0.98, gap = 0.2;
    const bw = 3.6, aw = CW - bw - 0.9;
    T(s, "مانع", { x: W - M - bw, y: 1.62, w: bw, h: 0.3, fontSize: 12, bold: true, color: C.accent6, objectName: "col barrier" });
    T(s, "پاسخ منابع انسانی", { x: M, y: 1.62, w: aw, h: 0.3, fontSize: 12, bold: true, color: C.accent5, objectName: "col response" });
    for (let i = 0; i < rows.length; i++) {
      const y = y0 + i * (rh + gap);
      card(s, W - M - bw, y, bw, rh, C.background2, "barrier " + (i + 1));
      T(s, rows[i][0], { x: W - M - bw + 0.3, y, w: bw - 0.6, h: rh, fontSize: 18, bold: true, color: C.text2, valign: "middle", objectName: "barrier text " + (i + 1) });
      s.addImage({ data: await icon("LuArrowLeft", HEX.accent1), x: W - M - bw - 0.68, y: y + rh / 2 - 0.23, w: 0.46, h: 0.46, objectName: "arrow " + (i + 1), altText: "پیکان" });
      card(s, M, y, aw, rh, C.background1, "response " + (i + 1), { line: { color: C.accent3, width: 1 } });
      T(s, rows[i][1], { x: M + 0.3, y, w: aw - 0.6, h: rh, fontSize: 15, color: C.text1, valign: "middle", objectName: "response text " + (i + 1) });
    }
    s.addNotes("(۳ دقیقه) چهار جمله‌ای که در هر شرکتی شنیده می‌شود و پاسخی که فقط از دست منابع انسانی برمی‌آید. سؤال از حضار: در شرکت شما کدام‌یک پررنگ‌تر است؟");
  }

  // =========================================================
  // 7. Incentive package
  {
    const s = slideIn("مشارکت");
    head(s, "۲ · افزایش مشارکت همکاران", "بسته‌ی تشویقی پیشنهادی برای هر شرکت");
    const inc = [
      ["LuEye", "دیده‌شدن", ["تقدیر در جلسه‌ی عمومی شرکت", "گواهی مشارکت برای همه‌ی ایده‌پردازان", "معرفی در کانال‌های داخلی"]],
      ["LuTrendingUp", "رشد", ["امتیاز مشارکت در ارزیابی عملکرد", "اولویت در دوره‌های آموزشی", "ثبت در پرونده‌ی شغلی"]],
      ["LuClock", "زمان", ["ساعت آزاد نوآوری در ماه", "زمان برای کارگاه و منتورینگ", "حمایت مدیر مستقیم"]],
      ["LuGift", "پاداش", ["هدیه برای هر ایده‌ی کامل", "پاداش برای ایده‌های پایلوت‌شده", "عضویت در باشگاه نوآوری"]],
    ];
    const statW = 3.3, gridW = CW - statW - 0.35, gap = 0.25, cw = (gridW - gap) / 2, chh = 2.25, y0 = 1.85;
    for (let i = 0; i < 4; i++) {
      const col = i % 2, row = Math.floor(i / 2);
      const x = W - M - cw - col * (cw + gap), y = y0 + row * (chh + gap);
      card(s, x, y, cw, chh, C.background2, "incentive " + (i + 1));
      await iconDisc(s, x + cw - 0.3 - 0.7, y + 0.3, 0.7, inc[i][0], C.accent1, HEX.lt1, "incentive " + (i + 1));
      T(s, inc[i][1], { x: x + 0.3, y: y + 0.42, w: cw - 1.5, h: 0.45, fontSize: 19, bold: true, color: C.text2, objectName: "incentive title " + (i + 1) });
      T(s, inc[i][2].map((t, k) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: k < inc[i][2].length - 1 } })),
        { x: x + 0.3, y: y + 1.1, w: cw - 0.6, h: 1.05, fontSize: 14, color: C.text1, paraSpaceAfter: 3, objectName: "incentive list " + (i + 1) });
    }
    const sx = M, sy = y0, sh = 2 * chh + gap;
    card(s, sx, sy, statW, sh, C.text2, "target panel");
    T(s, "هدف پیشنهادی", { x: sx + 0.3, y: sy + 0.45, w: statW - 0.6, h: 0.4, fontSize: 15, bold: true, color: C.accent3, align: "center", objectName: "target label" });
    T(s, "[۱۰٪]", { x: sx + 0.2, y: sy + 1.1, w: statW - 0.4, h: 1.3, fontSize: 60, bold: true, color: C.accent2, align: "center", valign: "middle", objectName: "target number" });
    T(s, "از همکاران هر شرکت، دست‌کم یک ایده ثبت کنند", { x: sx + 0.3, y: sy + 2.55, w: statW - 0.6, h: 0.8, fontSize: 15, color: C.background1, align: "center", objectName: "target desc" });
    T(s, "عدد هدف با هر شرکت توافق می‌شود و در داشبورد دنبال می‌شود", { x: sx + 0.3, y: sy + 3.6, w: statW - 0.6, h: 0.75, fontSize: 12, color: C.background2, align: "center", objectName: "target note" });
    s.addNotes("(۳ دقیقه) چهار دسته مشوق. بیشتر این‌ها هزینه‌ی کمی دارند و فقط تصمیم منابع انسانی را لازم دارند. عدد هدف (۱۰٪) پیشنهادی است و باید با هر شرکت نهایی شود.");
  }

  // =========================================================
  // 8. Platform flow
  {
    const s = slideIn("سامانه");
    head(s, "۳ · سامانه‌ی جدید در لایه‌ی شرکت‌ها", "هر شرکت دور داخلی خودش را دارد؛ بهترین‌ها به سطح گروه می‌آیند");
    const fl = [
      ["LuLightbulb", "ثبت ایده", "همکار", false],
      ["LuSearch", "بررسی اولیه", "سفیر نوآوری شرکت", false],
      ["LuGavel", "داوری داخلی", "داوران خود شرکت", false],
      ["LuStar", "ایده‌های برتر شرکت", "معرفی به سطح گروه", false],
      ["LuRocket", "رقابت گروه و پایلوت", "گلرنگ ونچرز", true],
    ];
    const n = fl.length, aw = 0.5, cw = (CW - (n - 1) * aw) / n, y = 2.1, ch = 2.55;
    // bracket: company layer vs group layer
    const compW = 3 * cw + 2 * aw + cw + aw; // first four steps
    card(s, W - M - (4 * cw + 3 * aw), y - 0.05, 4 * cw + 3 * aw, ch + 0.1, C.background2, "company layer bg");
    for (let i = 0; i < n; i++) {
      const x = W - M - cw - i * (cw + aw);
      const g = fl[i][3];
      if (g) card(s, x, y - 0.05, cw, ch + 0.1, C.text2, "group layer bg");
      await iconDisc(s, x + cw / 2 - 0.5, y + 0.3, 1.0, fl[i][0], g ? C.accent2 : C.accent1, g ? HEX.dk2 : HEX.lt1, "flow " + (i + 1));
      T(s, fl[i][1], { x: x + 0.1, y: y + 1.45, w: cw - 0.2, h: 0.5, fontSize: 16, bold: true, color: g ? C.background1 : C.text2, align: "center", objectName: "flow title " + (i + 1) });
      T(s, fl[i][2], { x: x + 0.1, y: y + 1.95, w: cw - 0.2, h: 0.45, fontSize: 13, color: g ? C.accent3 : C.accent1, align: "center", objectName: "flow who " + (i + 1) });
      if (i < n - 1) s.addImage({ data: await icon("LuArrowLeft", HEX.accent1), x: x - aw + 0.07, y: y + 0.6, w: aw - 0.14, h: aw - 0.14, objectName: "flow arrow " + (i + 1), altText: "پیکان" });
    }
    T(s, "لایه‌ی شرکت (با مدیریت منابع انسانی)", { x: W - M - (4 * cw + 3 * aw), y: y + ch + 0.2, w: 4 * cw + 3 * aw, h: 0.35, fontSize: 13, bold: true, color: C.accent1, align: "center", objectName: "company layer label" });
    T(s, "لایه‌ی گروه", { x: M, y: y + ch + 0.2, w: cw, h: 0.35, fontSize: 13, bold: true, color: C.text2, align: "center", objectName: "group layer label" });
    card(s, M, 5.6, CW, 0.95, C.background2, "benefit band");
    T(s, [
      { text: "چرا این مدل؟ ", options: { bold: true, color: C.text2 } },
      { text: "ایده‌ها نزدیک به محل کار بررسی می‌شوند، همکاران سریع‌تر جواب می‌گیرند و هر شرکت مالک نوآوری خودش است.", options: { color: C.text1 } },
    ], { x: M + 0.35, y: 5.6, w: CW - 0.7, h: 0.95, fontSize: 15, valign: "middle", objectName: "benefit text" });
    s.addNotes("(۳ دقیقه) معرفی سامانه‌ی [نام سامانه]. تغییر اصلی نسبت به قبل: جایزه به لایه‌ی شرکت‌ها می‌رود. چهار مرحله‌ی اول داخل هر شرکت و با مدیریت منابع انسانی انجام می‌شود؛ ایده‌های برتر هر شرکت به رقابت گروه می‌آیند.");
  }

  // =========================================================
  // 9. Roles
  {
    const s = slideIn("سامانه");
    head(s, "۳ · سامانه‌ی جدید در لایه‌ی شرکت‌ها", "چه کسی در سامانه چه می‌کند");
    const roles = [
      ["LuUsers", "همکاران", ["ثبت ایده از موبایل یا رایانه", "پیگیری وضعیت ایده", "دریافت بازخورد داوران"], false],
      ["LuUserCheck", "منابع انسانی شرکت", ["راه‌اندازی فراخوان شرکت", "تعیین داوران داخلی", "پیگیری مشارکت و گزارش به مدیرعامل"], true],
      ["LuGavel", "داوران داخلی", ["امتیازدهی آنلاین", "معیارهای یکسان در همه‌ی شرکت‌ها", "نوشتن بازخورد کوتاه"], false],
      ["LuRocket", "گلرنگ ونچرز", ["تعیین معیارها و آموزش داوران", "داوری نهایی در سطح گروه", "پشتیبانی و اجرای پایلوت"], false],
    ];
    const gap = 0.3, cw = (CW - 3 * gap) / 4, y = 1.95, ch = 4.55;
    for (let i = 0; i < 4; i++) {
      const x = W - M - cw - i * (cw + gap);
      const hl = roles[i][3];
      card(s, x, y, cw, ch, hl ? C.text2 : C.background2, "role " + (i + 1));
      if (hl) {
        card(s, x + 0.3, y + 0.3, 1.25, 0.36, C.accent2, "role badge");
        T(s, "نقش شما", { x: x + 0.3, y: y + 0.3, w: 1.25, h: 0.36, fontSize: 12, bold: true, color: C.text2, align: "center", valign: "middle", objectName: "role badge text" });
      }
      await iconDisc(s, x + cw - 0.3 - 0.9, y + 0.3, 0.9, roles[i][0], hl ? C.accent2 : C.accent1, hl ? HEX.dk2 : HEX.lt1, "role " + (i + 1));
      T(s, roles[i][1], { x: x + 0.3, y: y + 1.45, w: cw - 0.6, h: 0.5, fontSize: 19, bold: true, color: hl ? C.background1 : C.text2, objectName: "role title " + (i + 1) });
      T(s, roles[i][2].map((t, k) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: k < 2 } })),
        { x: x + 0.3, y: y + 2.1, w: cw - 0.6, h: 2.2, fontSize: 14, color: hl ? C.background1 : C.text1, paraSpaceAfter: 8, objectName: "role list " + (i + 1) });
    }
    s.addNotes("(۲ دقیقه) چهار نقش. نقش منابع انسانی پررنگ شده: صاحب فراخوان در شرکت. سامانه کار اجرایی را ساده می‌کند: فرم، یادآوری و گزارش آماده است.");
  }

  // =========================================================
  // 10. Platform features
  {
    const s = slideIn("سامانه");
    head(s, "۳ · سامانه‌ی جدید در لایه‌ی شرکت‌ها", "قابلیت‌های اصلی سامانه");
    const ft = [
      ["LuSmartphone", "فرم ساده‌ی ثبت ایده", "روی موبایل و رایانه، در چند دقیقه"],
      ["LuClipboardCheck", "پیگیری وضعیت", "ایده‌پرداز همیشه می‌داند ایده‌اش کجاست"],
      ["LuGavel", "داوری آنلاین", "معیارهای یکسان، امتیاز و بازخورد در یک جا"],
      ["LuBell", "اعلان و یادآوری خودکار", "برای همکاران، داوران و سفیران"],
      ["LuFileText", "گزارش مشارکت برای منابع انسانی", "به تفکیک واحد، آماده برای مدیرعامل"],
      ["LuLayoutDashboard", "اتصال به داشبورد زنده", "داده‌ی همه‌ی شرکت‌ها در یک نما"],
    ];
    const lw = 7.0, x0 = W - M - lw, y0 = 1.9, rh = 0.78;
    for (let i = 0; i < ft.length; i++) {
      const y = y0 + i * rh;
      await iconDisc(s, W - M - 0.62, y + 0.04, 0.62, ft[i][0], C.accent1, HEX.lt1, "feature " + (i + 1));
      T(s, ft[i][1], { x: x0, y: y + 0.02, w: lw - 0.85, h: 0.36, fontSize: 16, bold: true, color: C.text2, objectName: "feature title " + (i + 1) });
      T(s, ft[i][2], { x: x0, y: y + 0.38, w: lw - 0.85, h: 0.32, fontSize: 13, color: C.accent4, objectName: "feature body " + (i + 1) });
    }
    // screenshot placeholder frame (left)
    const fx = M, fw = x0 - 0.45 - M, fy = 1.9, fh = 4.65;
    card(s, fx, fy, fw, fh, C.text2, "screen frame");
    card(s, fx + 0.2, fy + 0.45, fw - 0.4, fh - 0.65, C.background2, "screen area", { rectRadius: 0.06 });
    [0, 1, 2].forEach((k) => s.addShape(pres.shapes.OVAL, { x: fx + fw - 0.45 - k * 0.22, y: fy + 0.17, w: 0.13, h: 0.13, fill: { color: k === 0 ? C.accent2 : C.accent3 }, line: { type: "none" }, objectName: "window dot " + (k + 1) }));
    s.addImage({ data: await icon("LuMonitorSmartphone", HEX.accent3), x: fx + fw / 2 - 0.6, y: fy + 1.6, w: 1.2, h: 1.2, objectName: "screen placeholder icon", altText: "جای تصویر سامانه" });
    T(s, "[تصویر صفحه‌ی سامانه]", { x: fx + 0.3, y: fy + 2.95, w: fw - 0.6, h: 0.4, fontSize: 15, bold: true, color: C.accent1, align: "center", objectName: "screen placeholder text" });
    s.addNotes("(۲ دقیقه) شش قابلیت. اگر نسخه‌ی آزمایشی آماده است، اینجا یک دموی کوتاه یا اسکرین‌شات نشان بدهید. جای تصویر در سمت چپ اسلاید است.");
  }

  // =========================================================
  // 11. Innovation club — why
  {
    const s = slideIn("باشگاه نوآوری");
    head(s, "۴ · معرفی باشگاه نوآوری", "رابطه با ایده‌پردازان نباید با اعلام نتایج تمام شود");
    const qx = W - M - 6.6, qy = 1.95, qw = 6.6, qh = 4.6;
    card(s, qx, qy, qw, qh, C.text2, "club intro panel");
    await iconDisc(s, qx + qw - 0.4 - 0.9, qy + 0.4, 0.9, "LuAward", C.accent2, HEX.dk2, "club");
    T(s, "چرا باشگاه نوآوری؟", { x: qx + 0.4, y: qy + 1.5, w: qw - 0.8, h: 0.5, fontSize: 22, bold: true, color: C.accent2, objectName: "club why title" });
    T(s, "ایده‌های خوب از دل کار روزمره‌ی همکاران بیرون می‌آیند. جایزه این ایده‌ها و صاحبانشان را پیدا می‌کند و باشگاه این رابطه را ادامه می‌دهد: جایی که ایده‌پردازان دیده شوند، پشتیبانی بگیرند، مهارتشان رشد کند و ایده‌ی بعدی را با انگیزه‌ی بیشتر بیاورند.",
      { x: qx + 0.4, y: qy + 2.1, w: qw - 0.8, h: 2.3, fontSize: 16, color: C.background1, lineSpacingMultiple: 1.25, objectName: "club why body" });
    const pil = [
      ["LuEye", "دیده‌شدن", "لوح، بج سفیر نوآوری و معرفی در «گلرنگ من»"],
      ["LuHandshake", "پشتیبانی", "مزایای رفاهی و مالی از شرکت و گروه"],
      ["LuGraduationCap", "رشد", "آموزش، یک روز کنار مدیر ارشد و رویداد بین‌المللی"],
    ];
    const px = M, pw = qx - 0.4 - M, ph = (qh - 2 * 0.25) / 3;
    for (let i = 0; i < 3; i++) {
      const y = qy + i * (ph + 0.25);
      card(s, px, y, pw, ph, C.background2, "pillar " + (i + 1));
      await iconDisc(s, px + pw - 0.3 - 0.75, y + ph / 2 - 0.375, 0.75, pil[i][0], C.accent1, HEX.lt1, "pillar " + (i + 1));
      T(s, pil[i][1], { x: px + 0.3, y: y + 0.22, w: pw - 1.5, h: 0.42, fontSize: 18, bold: true, color: C.text2, objectName: "pillar title " + (i + 1) });
      T(s, pil[i][2], { x: px + 0.3, y: y + 0.66, w: pw - 1.5, h: 0.6, fontSize: 13, color: C.text1, objectName: "pillar body " + (i + 1) });
    }
    s.addNotes("(۲ دقیقه) مقدمه‌ی آیین‌نامه‌ی باشگاه. هدف: تبدیل نوآوری از یک رویداد سالانه به بخشی از فرهنگ کار در گروه. سه ستون: دیده‌شدن، پشتیبانی، رشد.");
  }

  // =========================================================
  // 12. Club benefits & who provides
  {
    const s = slideIn("باشگاه نوآوری");
    head(s, "۴ · معرفی باشگاه نوآوری", "مزایای اعضا و سهم هر طرف");
    const comp = [
      ["LuBanknote", "فوق‌العاده‌ی نوآوری", "در فیش حقوقی"],
      ["LuWallet", "وام", "تا ۳٫۵ برابر پایه‌ی حقوق"],
      ["LuCreditCard", "کارت اعتباری تارا", "سقف بالاتر از سایر همکاران"],
      ["LuGraduationCap", "آموزش", "همه‌ی آموزش‌های گروه در اختیار اعضا"],
      ["LuUsers", "یک روز کنار مدیر ارشد", "یک روز کاری همراه مدیر ارشد شرکت"],
    ];
    const gv = [
      ["LuIdCard", "لوح، بج و کارت عضویت", "بج سفیر نوآوری و کارت شماره‌دار"],
      ["LuTag", "تگ نوآوری", "در «گلرنگ من» و معرفی در پورتال"],
      ["LuFlaskConical", "تست محصول استارتاپ‌ها", "پیش از لانچ، همراه با اعتبار تست"],
      ["LuSun", "روز خانوادگی باشگاه", "یک روز در سال در مجموعه‌ی کوروش"],
      ["LuPlane", "جایزه‌ی سالانه", "اعزام به رویداد بین‌المللی فناوری، برای نفر اول"],
    ];
    const gap = 0.4, cw = (CW - gap) / 2, y0 = 1.85, hh = 0.72, rh = 0.8;
    const cols = [[W - M - cw, "سهم شرکت محل خدمت (منابع انسانی شرکت)", comp, true], [M, "سهم گلرنگ ونچرز و گروه", gv, false]];
    for (const [x, title, list, hl] of cols) {
      card(s, x, y0, cw, 0.62, hl ? C.accent2 : C.text2, "col head " + (hl ? "company" : "group"));
      T(s, title, { x: x + 0.3, y: y0, w: cw - 0.6, h: 0.62, fontSize: 16, bold: true, color: hl ? C.text2 : C.background1, valign: "middle", objectName: "col head text " + (hl ? "company" : "group") });
      card(s, x, y0 + hh, cw, list.length * rh + 0.15, C.background2, "col body " + (hl ? "company" : "group"));
      for (let i = 0; i < list.length; i++) {
        const y = y0 + hh + 0.1 + i * rh;
        await iconDisc(s, x + cw - 0.25 - 0.56, y + 0.1, 0.56, list[i][0], hl ? C.text2 : C.accent1, HEX.lt1, (hl ? "company " : "group ") + (i + 1));
        T(s, list[i][1], { x: x + 0.25, y: y + 0.07, w: cw - 1.15, h: 0.34, fontSize: 15, bold: true, color: C.text2, objectName: (hl ? "company " : "group ") + "title " + (i + 1) });
        T(s, list[i][2], { x: x + 0.25, y: y + 0.42, w: cw - 1.15, h: 0.3, fontSize: 12, color: C.accent4, objectName: (hl ? "company " : "group ") + "body " + (i + 1) });
      }
    }
    s.addNotes("(۲ دقیقه) همه‌ی اعضا این مزایا را فقط با عضویت می‌گیرند. پنج مزیت ستون راست را شرکت محل خدمت هر عضو می‌دهد و این همان جایی است که تصمیم منابع انسانی لازم است. بقیه را گلرنگ ونچرز و گروه تأمین می‌کنند. رویدادهای نمونه برای جایزه‌ی سالانه: جیتکس، وب‌سامیت یا اسلاش.");
  }

  // =========================================================
  // 13. Dashboard mockup
  {
    const s = slideIn("داشبورد");
    head(s, "۵ · داشبورد زنده‌ی جایزه نوآوری", "وضعیت جایزه در همه‌ی شرکت‌ها، در یک نگاه");
    const kpi = [
      ["ایده‌ی ثبت‌شده", "۲۴۸"],
      ["شرکت فعال", "۱۸"],
      ["نرخ مشارکت همکاران", "۶٫۴٪"],
      ["ایده در مرحله‌ی داوری", "۵۲"],
    ];
    const gap = 0.25, kw = (CW - 3 * gap) / 4, ky = 1.85, kh = 1.15;
    for (let i = 0; i < 4; i++) {
      const x = W - M - kw - i * (kw + gap);
      card(s, x, ky, kw, kh, i === 0 ? C.text2 : C.background2, "kpi " + (i + 1));
      T(s, kpi[i][1], { x: x + 0.3, y: ky + 0.12, w: kw - 0.6, h: 0.6, fontSize: 30, bold: true, color: i === 0 ? C.accent2 : C.text2, objectName: "kpi value " + (i + 1) });
      T(s, kpi[i][0], { x: x + 0.3, y: ky + 0.72, w: kw - 0.6, h: 0.32, fontSize: 13, color: i === 0 ? C.background2 : C.accent4, objectName: "kpi label " + (i + 1) });
    }
    // chart: ideas per company (sample)
    const cy = ky + kh + 0.3, chH = 6.55 - cy;
    const cwid = 7.6;
    card(s, W - M - cwid, cy, cwid, chH, C.background1, "chart frame", { line: { color: "D5DFEE", width: 1 } });
    const companies = ["شرکت الف", "شرکت ب", "شرکت ج", "شرکت د", "شرکت ه", "شرکت و"];
    const vals = [46, 39, 33, 27, 21, 14];
    s.addChart(pres.charts.BAR, [{ name: "ایده‌ی ثبت‌شده", labels: companies.slice().reverse(), values: vals.slice().reverse() }], {
      x: W - M - cwid + 0.15, y: cy + 0.1, w: cwid - 0.3, h: chH - 0.2, barDir: "bar",
      showTitle: true, title: "ایده‌های ثبت‌شده به تفکیک شرکت (داده‌ی نمونه)", titleFontSize: 13, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
      chartColors: [HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HEX.dk1, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      catAxisLineShow: false, showLegend: false, barGapWidthPct: 60, catAxisOrientation: "minMax", objectName: "ideas by company chart",
    });
    // funnel (left)
    const fx = M, fw = W - M - cwid - 0.3 - M;
    card(s, fx, cy, fw, chH, C.background2, "funnel frame");
    T(s, "قیف ایده‌ها", { x: fx + 0.3, y: cy + 0.2, w: fw - 0.6, h: 0.38, fontSize: 15, bold: true, color: C.text2, objectName: "funnel title" });
    const fun = [["ثبت‌شده", 248], ["داوری داخلی", 120], ["برتر شرکت‌ها", 36], ["فینال گروه", 12], ["پایلوت", 4]];
    const barMax = fw - 0.6 - 1.35, fy0 = cy + 0.7, fh = 0.38;
    for (let i = 0; i < fun.length; i++) {
      const y = fy0 + i * (fh + 0.1);
      const bw = Math.max(0.35, barMax * Math.sqrt(fun[i][1] / 248));
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: fx + fw - 0.3 - bw, y, w: bw, h: fh, rectRadius: 0.06, fill: { color: i === fun.length - 1 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: "funnel bar " + (i + 1) });
      T(s, fa(fun[i][1]), { x: fx + fw - 0.3 - bw, y, w: bw - 0.12, h: fh, fontSize: 13, bold: true, color: i === fun.length - 1 ? C.text2 : C.background1, valign: "middle", objectName: "funnel value " + (i + 1) });
      T(s, fun[i][0], { x: fx + 0.2, y, w: fw - 0.6 - bw - 0.05, h: fh, fontSize: 12, color: C.text1, align: "left", valign: "middle", objectName: "funnel label " + (i + 1) });
    }
    T(s, "همه‌ی اعداد این صفحه نمونه است؛ نسخه‌ی زنده در حال تکمیل است.", { x: M, y: 6.6, w: CW, h: 0.3, fontSize: 11, color: C.accent4, objectName: "sample note" });
    s.addNotes("(۲ دقیقه) نمای داشبورد. اعداد نمونه است و نسخه‌ی زنده در حال تکمیل است. اگر لینک نسخه‌ی فعلی آماده است، همین‌جا نشان بدهید. چهار شاخص بالا، ایده‌ها به تفکیک شرکت، و قیف از ثبت تا پایلوت.");
  }

  // =========================================================
  // 14. What dashboard gives HR
  {
    const s = slideIn("داشبورد");
    head(s, "۵ · داشبورد زنده‌ی جایزه نوآوری", "داشبورد برای منابع انسانی چه می‌کند");
    const it = [
      ["LuChartColumn", "جایگاه شرکت شما", "مقایسه‌ی مشارکت با شرکت‌های دیگر گروه؛ یک رقابت سالم و دیدنی"],
      ["LuFilter", "گلوگاه‌ها را نشان می‌دهد", "ایده‌ها کجا گیر می‌کنند: ثبت، داوری یا پاسخ‌دهی؟"],
      ["LuPresentation", "گزارش آماده برای مدیرعامل", "بدون جمع‌آوری دستی داده؛ هر لحظه به‌روز"],
    ];
    const gap = 0.35, cw = (CW - 2 * gap) / 3, y = 2.0, ch = 3.0;
    for (let i = 0; i < 3; i++) {
      const x = W - M - cw - i * (cw + gap);
      card(s, x, y, cw, ch, C.background2, "dash card " + (i + 1));
      await iconDisc(s, x + cw - 0.4 - 0.9, y + 0.4, 0.9, it[i][0], C.accent1, HEX.lt1, "dash " + (i + 1));
      T(s, it[i][1], { x: x + 0.4, y: y + 1.5, w: cw - 0.8, h: 0.5, fontSize: 19, bold: true, color: C.text2, objectName: "dash title " + (i + 1) });
      T(s, it[i][2], { x: x + 0.4, y: y + 2.05, w: cw - 0.8, h: 0.85, fontSize: 14, color: C.text1, objectName: "dash body " + (i + 1) });
    }
    card(s, M, 5.4, CW, 1.15, C.text2, "access band");
    await iconDisc(s, W - M - 0.35 - 0.7, 5.62, 0.7, "LuShieldCheck", C.accent2, HEX.dk2, "access");
    T(s, [
      { text: "دسترسی: ", options: { bold: true, color: C.accent2 } },
      { text: "هر شرکت داده‌ی جزئی خودش را می‌بیند و نمای کلی گروه برای همه قابل مشاهده است.", options: { color: C.background1 } },
    ], { x: M + 0.4, y: 5.4, w: CW - 1.7, h: 1.15, fontSize: 16, valign: "middle", objectName: "access text" });
    s.addNotes("(۱ دقیقه) سه کاربرد داشبورد برای منابع انسانی. سطح دسترسی پیشنهادی است و باید تأیید شود.");
  }

  // =========================================================
  // 15. The ask
  {
    const s = slideIn("جمع‌بندی");
    head(s, "جمع‌بندی", "پنج کاری که از منابع انسانی هر شرکت می‌خواهیم");
    const asks = [
      ["معرفی سفیر نوآوری", "یک نفر از منابع انسانی، تا [تاریخ]"],
      ["اجرای تقویم اطلاع‌رسانی", "با کیت آماده‌ی پیام و پوستر از گلرنگ ونچرز"],
      ["تصویب بسته‌ی تشویقی", "دیده‌شدن، رشد، زمان و پاداش، متناسب با شرکت"],
      ["راه‌اندازی سامانه در شرکت", "فعال‌سازی فراخوان و تعیین داوران داخلی"],
      ["تأمین مزایای باشگاه", "پنج مزیت سهم شرکت محل خدمت برای اعضا"],
    ];
    const lw = 7.6, x0 = W - M - lw, y0 = 1.9, rh = 0.9;
    for (let i = 0; i < asks.length; i++) {
      const y = y0 + i * rh;
      card(s, x0, y, lw, rh - 0.15, C.background2, "ask " + (i + 1));
      s.addShape(pres.shapes.OVAL, { x: W - M - 0.2 - 0.5, y: y + 0.125, w: 0.5, h: 0.5, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "ask num disc " + (i + 1) });
      T(s, fa(i + 1), { x: W - M - 0.2 - 0.5, y: y + 0.125, w: 0.5, h: 0.5, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", objectName: "ask num " + (i + 1) });
      T(s, [
        { text: asks[i][0], options: { bold: true, color: C.text2, fontSize: 16 } },
        { text: "  " + asks[i][1], options: { color: C.accent4, fontSize: 13 } },
      ], { x: x0 + 0.25, y, w: lw - 1.15, h: rh - 0.15, valign: "middle", objectName: "ask text " + (i + 1) });
    }
    const px = M, pw = x0 - 0.4 - M, py = y0, ph = 5 * rh - 0.15;
    card(s, px, py, pw, ph, C.text2, "next step panel");
    await iconDisc(s, px + pw - 0.35 - 0.8, py + 0.35, 0.8, "LuCalendarDays", C.accent2, HEX.dk2, "next step");
    T(s, "گام بعدی", { x: px + 0.35, y: py + 1.35, w: pw - 0.7, h: 0.45, fontSize: 20, bold: true, color: C.accent2, objectName: "next title" });
    T(s, "جلسه‌ی هماهنگی با سفیران نوآوری همه‌ی شرکت‌ها", { x: px + 0.35, y: py + 1.85, w: pw - 0.7, h: 0.8, fontSize: 16, color: C.background1, objectName: "next body" });
    T(s, "[تاریخ و محل جلسه]", { x: px + 0.35, y: py + 2.75, w: pw - 0.7, h: 0.4, fontSize: 14, bold: true, color: C.accent3, objectName: "next date" });
    T(s, "در این جلسه کیت اطلاع‌رسانی و دسترسی سامانه تحویل داده می‌شود.", { x: px + 0.35, y: py + 3.3, w: pw - 0.7, h: 0.8, fontSize: 13, color: C.background2, objectName: "next note" });
    s.addNotes("(۳ دقیقه) جمع‌بندی و درخواست. پنج کار مشخص. اگر فقط یک چیز در ذهنتان بماند: معرفی سفیر نوآوری تا [تاریخ]. گام بعدی جلسه‌ی هماهنگی سفیران است.");
  }

  // =========================================================
  // 16. Closing
  {
    const s = slideIn("جمع‌بندی", "COVER");
    const cx = 2.45, cy = 3.75;
    [[4.6, 70], [3.4, 45]].forEach(([d, tr], i) =>
      s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: cy - d / 2, w: d, h: d, fill: { color: C.accent1, transparency: tr }, line: { type: "none" }, objectName: "ring " + (i + 1) })
    );
    card(s, cx - 1.1, cy - 1.1, 2.2, 2.2, C.background1, "qr box", { rectRadius: 0.1 });
    s.addImage({ data: await icon("LuQrCode", HEX.dk2), x: cx - 0.55, y: cy - 0.75, w: 1.1, h: 1.1, objectName: "qr placeholder icon", altText: "جای کد QR" });
    T(s, "[QR سامانه]", { x: cx - 1.0, y: cy + 0.45, w: 2.0, h: 0.4, fontSize: 13, bold: true, color: C.accent1, align: "center", objectName: "qr label" });
    s.addText("رویداد تراست · گلرنگ ونچرز", { placeholder: "kicker" });
    s.addText("نوآوری از آدم‌ها شروع می‌شود", { placeholder: "title" });
    s.addText("و آدم‌ها با شما. سپاس از همراهی‌تان؛ آماده‌ی پرسش‌ها هستیم.", { placeholder: "body" });
    T(s, "[نام ارائه‌دهنده] · [ایمیل] · [شماره‌ی تماس]", { x: 5.2, y: 5.55, w: 7.53, h: 0.35, fontSize: 14, color: C.background1, objectName: "contact" });
    s.addNotes("پایان و پرسش و پاسخ. QR را به لینک سامانه یا فرم ثبت سفیر نوآوری وصل کنید.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
