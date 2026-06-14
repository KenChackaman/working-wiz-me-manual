import { toPng } from "https://esm.sh/html-to-image@1.11.11";
import pptxgenModule from "https://esm.sh/pptxgenjs@3.12.0";

function resolvePptxCtor(mod) {
  let cur = mod;
  for (let i = 0; i < 3; i += 1) {
    if (typeof cur === "function") return cur;
    if (cur && typeof cur.default === "function") cur = cur.default;
    else break;
  }
  return typeof cur === "function" ? cur : null;
}

const PptxGenJS = resolvePptxCtor(pptxgenModule.default ?? pptxgenModule);

const $ = (id) => /** @type {HTMLElement} */ (document.getElementById(id));

const fields = [
  "displayName",
  "department",
  "role",
  "strengths",
  "decisionStyle",
  "focusHours",
  "channels",
  "responseSLA",
  "meetingPrefs",
  "commStyle",
  "feedbackGood",
  "feedbackStress",
  "misunderstood",
  "helpMe",
  "avoidRoles",
  "energizers",
  "stressful",
  "boundaries",
  "footerNote",
];

const STORAGE_KEY = "working-with-me-form-v1";

function readForm() {
  /** @type {Record<string, string>} */
  const data = {};
  for (const id of fields) {
    const el = /** @type {HTMLInputElement | HTMLTextAreaElement | null} */ (document.getElementById(id));
    data[id] = (el?.value ?? "").trim();
  }
  return data;
}

function joinBlocks(lines) {
  return lines
    .map((x) => (typeof x === "string" ? x.trim() : ""))
    .filter(Boolean)
    .join("\n\n");
}

function setText(id, text, fallback = "（未入力）") {
  const el = $(id);
  if (!el) return;
  const v = text.trim();
  el.textContent = v || fallback;
}

function updatePreview() {
  const d = readForm();

  setText("pvName", d.displayName, "（表示名）");
  setText("pvDept", d.department, "（所属）");

  const roleEl = $("pvRole");
  if (roleEl) {
    if (d.role) {
      roleEl.style.display = "block";
      roleEl.textContent = d.role;
    } else {
      roleEl.style.display = "none";
      roleEl.textContent = "";
    }
  }

  const updated = $("pvUpdated");
  if (updated) {
    const dt = new Date();
    const s = new Intl.DateTimeFormat("ja-JP", { dateStyle: "medium" }).format(dt);
    updated.textContent = `更新日：${s}`;
  }

  setText("pvStrengths", d.strengths);

  setText(
    "pvComm",
    joinBlocks([d.channels, d.responseSLA, d.meetingPrefs, d.commStyle].filter(Boolean)),
  );

  setText("pvDecision", joinBlocks([d.decisionStyle, d.focusHours].filter(Boolean)));

  setText("pvFeedback", joinBlocks([d.feedbackGood, d.feedbackStress].filter(Boolean)));

  setText("pvTeam", joinBlocks([d.misunderstood, d.helpMe].filter(Boolean)));

  setText(
    "pvEnergy",
    joinBlocks([d.avoidRoles, d.energizers, d.stressful, d.boundaries].filter(Boolean)),
  );

  const footer = $("pvFooter");
  if (footer) {
    const base =
      "※これは「協力のためのガイド」であり、相手への要求リストではありません。更新しながら育てていきましょう。";
    footer.textContent = d.footerNote ? `${d.footerNote}\n${base}` : base;
  }
}

function saveLocal() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(readForm()));
  } catch {
    // ignore
  }
}

function loadLocal() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    for (const id of fields) {
      const el = /** @type {HTMLInputElement | HTMLTextAreaElement | null} */ (document.getElementById(id));
      if (!el) continue;
      const v = data[id];
      if (typeof v === "string") el.value = v;
    }
  } catch {
    // ignore
  }
}

function validateRequired() {
  const d = readForm();
  const missing = [];
  if (!d.displayName) missing.push("表示名");
  if (!d.department) missing.push("所属");
  if (!d.strengths) missing.push("得意な貢献・強み");
  return missing;
}

function setStatus(msg) {
  const el = $("status");
  if (el) el.textContent = msg ?? "";
}

function safeBaseName(name) {
  const base = (name || "working-with-me").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
  return base.slice(0, 80);
}

async function exportPng() {
  const missing = validateRequired();
  if (missing.length) {
    setStatus(`PNG出力の前に入力してください：${missing.join("、")}`);
    return;
  }

  const card = /** @type {HTMLElement} */ ($("card"));
  const btn = /** @type {HTMLButtonElement} */ ($("btnPng"));
  btn.disabled = true;
  setStatus("PNGを生成しています…");

  /** @type {HTMLElement | null} */
  let clone = null;
  try {
    updatePreview();
    await new Promise((r) => requestAnimationFrame(r));

    clone = /** @type {HTMLElement} */ (card.cloneNode(true));
    clone.style.position = "fixed";
    clone.style.left = "-12000px";
    clone.style.top = "0";
    clone.style.width = "1920px";
    clone.style.height = "1080px";
    clone.style.margin = "0";
    clone.style.boxSizing = "border-box";
    document.body.appendChild(clone);

    const dataUrl = await toPng(clone, {
      pixelRatio: 1,
      width: 1920,
      height: 1080,
      cacheBust: true,
      backgroundColor: "#ffffff",
    });

    const a = document.createElement("a");
    const d = readForm();
    const stamp = new Intl.DateTimeFormat("ja-JP", { dateStyle: "short" }).format(new Date()).replace(/\//g, "-");
    a.download = `${safeBaseName(d.displayName)}_working-with-me_${stamp}.png`;
    a.href = dataUrl;
    a.click();
    setStatus("PNGをダウンロードしました。");
  } catch (e) {
    console.error(e);
    setStatus("PNGの生成に失敗しました。ローカルサーバー経由で開いているか確認してください。");
  } finally {
    clone?.remove();
    btn.disabled = false;
  }
}

function addPptxBlock(slide, title, body, x, y, w, h) {
  slide.addText(title, {
    x,
    y,
    w,
    h: 0.22,
    fontSize: 12,
    bold: true,
    color: "BF0000",
    fontFace: "Meiryo",
  });
  slide.addText(body.trim() || " ", {
    x,
    y: y + 0.22,
    w,
    h: h - 0.26,
    fontSize: 11,
    color: "222222",
    fontFace: "Meiryo",
    valign: "top",
    shrinkText: true,
  });
}

function exportPptx() {
  const missing = validateRequired();
  if (missing.length) {
    setStatus(`PowerPoint出力の前に入力してください：${missing.join("、")}`);
    return;
  }

  if (!PptxGenJS) {
    setStatus("PowerPoint生成ライブラリの読み込みに失敗しました。ネットワークを確認してください。");
    return;
  }

  const d = readForm();
  updatePreview();

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  pptx.subject = "Working with me";
  pptx.title = "私の取扱説明書";

  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  const rectShape =
    pptx?.shapes?.RECTANGLE ??
    pptx?.ShapeType?.rect ??
    PptxGenJS?.ShapeType?.rect ??
    PptxGenJS?.shapes?.RECTANGLE;
  if (rectShape) {
    slide.addShape(rectShape, {
      x: 0,
      y: 0,
      w: 10,
      h: 0.14,
      fill: { color: "BF0000" },
      line: { color: "BF0000", width: 0 },
    });
  }

  slide.addText("私の取扱説明書", {
    x: 0.45,
    y: 0.22,
    w: 6.2,
    h: 0.55,
    fontSize: 28,
    bold: true,
    color: "111111",
    fontFace: "Meiryo",
  });
  slide.addText("Working with me（一緒に働くときのガイド）", {
    x: 0.45,
    y: 0.72,
    w: 6.2,
    h: 0.35,
    fontSize: 13,
    color: "5C5C5C",
    fontFace: "Meiryo",
    italic: true,
  });

  const stamp = new Intl.DateTimeFormat("ja-JP", { dateStyle: "medium" }).format(new Date());
  slide.addText(d.displayName, {
    x: 6.85,
    y: 0.22,
    w: 2.7,
    h: 0.45,
    fontSize: 20,
    bold: true,
    color: "111111",
    fontFace: "Meiryo",
    align: "right",
  });
  slide.addText(d.department, {
    x: 6.85,
    y: 0.62,
    w: 2.7,
    h: 0.55,
    fontSize: 12,
    color: "333333",
    fontFace: "Meiryo",
    align: "right",
    valign: "top",
    wrap: true,
  });
  if (d.role) {
    slide.addText(d.role, {
      x: 6.85,
      y: 1.05,
      w: 2.7,
      h: 0.25,
      fontSize: 11,
      color: "5C5C5C",
      fontFace: "Meiryo",
      align: "right",
    });
  }
  slide.addText(`更新日：${stamp}`, {
    x: 6.85,
    y: 1.28,
    w: 2.7,
    h: 0.2,
    fontSize: 10,
    color: "5C5C5C",
    fontFace: "Meiryo",
    align: "right",
  });

  const comm = joinBlocks([d.channels, d.responseSLA, d.meetingPrefs, d.commStyle].filter(Boolean));
  const decision = joinBlocks([d.decisionStyle, d.focusHours].filter(Boolean));
  const feedback = joinBlocks([d.feedbackGood, d.feedbackStress].filter(Boolean));
  const team = joinBlocks([d.misunderstood, d.helpMe].filter(Boolean));
  const energy = joinBlocks([d.avoidRoles, d.energizers, d.stressful, d.boundaries].filter(Boolean));

  const y0 = 1.62;
  const colW = 4.55;
  const gap = 0.35;
  const leftX = 0.45;
  const rightX = leftX + colW + gap;
  const blockH = 1.28;

  addPptxBlock(slide, "得意な貢献・強み", d.strengths, leftX, y0, colW, blockH);
  addPptxBlock(slide, "コミュニケーション", comm, rightX, y0, colW, blockH);

  addPptxBlock(slide, "意思決め・集中", decision, leftX, y0 + blockH + 0.12, colW, blockH);
  addPptxBlock(slide, "フィードバック", feedback, rightX, y0 + blockH + 0.12, colW, blockH);

  addPptxBlock(slide, "誤解されがち／助けてほしいこと", team, leftX, y0 + 2 * (blockH + 0.12), colW, blockH);
  addPptxBlock(slide, "エネルギー／注意・境界", energy, rightX, y0 + 2 * (blockH + 0.12), colW, blockH);

  const baseNote =
    "※これは「協力のためのガイド」であり、相手への要求リストではありません。更新しながら育てていきましょう。";
  const note = d.footerNote ? `${d.footerNote} / ${baseNote}` : baseNote;
  slide.addText(note, {
    x: 0.45,
    y: 5.05,
    w: 9.1,
    h: 0.45,
    fontSize: 9,
    color: "5C5C5C",
    fontFace: "Meiryo",
    valign: "top",
    wrap: true,
  });

  const fileStamp = new Intl.DateTimeFormat("ja-JP", { dateStyle: "short" }).format(new Date()).replace(/\//g, "-");
  pptx.writeFile({ fileName: `${safeBaseName(d.displayName)}_working-with-me_${fileStamp}.pptx` });
  setStatus("PowerPointをダウンロードしました。");
}

function resetForm() {
  for (const id of fields) {
    const el = /** @type {HTMLInputElement | HTMLTextAreaElement | null} */ (document.getElementById(id));
    if (el) el.value = "";
  }
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // ignore
  }
  updatePreview();
  setStatus("入力をクリアしました。");
}

function wire() {
  for (const id of fields) {
    const el = document.getElementById(id);
    if (!el) continue;
    el.addEventListener("input", () => {
      updatePreview();
      saveLocal();
      setStatus("");
    });
  }

  $("btnPng")?.addEventListener("click", () => void exportPng());
  $("btnPptx")?.addEventListener("click", () => {
    try {
      exportPptx();
    } catch (e) {
      console.error(e);
      setStatus("PowerPointの生成に失敗しました。ブラウザのコンソールを確認してください。");
    }
  });
  $("btnReset")?.addEventListener("click", resetForm);
}

function setupPreviewScaling() {
  const frame = /** @type {HTMLElement | null} */ (document.getElementById("previewFrame"));
  const root = /** @type {HTMLElement | null} */ (document.getElementById("exportRoot"));
  if (!frame || !root) return;

  const apply = () => {
    const w = frame.clientWidth;
    const scale = w > 0 ? w / 1920 : 0.5;
    root.style.setProperty("--card-scale", String(scale));
  };

  apply();

  if (typeof ResizeObserver !== "undefined") {
    const ro = new ResizeObserver(() => apply());
    ro.observe(frame);
  } else {
    window.addEventListener("resize", apply);
  }
}

loadLocal();
wire();
updatePreview();
setupPreviewScaling();
