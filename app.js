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

/** 各項目: 3つの候補（本文）＋下の自由記述欄を結合してプレビュー／出力に反映 */
const PRESETS = {
  strengths: {
    title: "得意な貢献・強み",
    options: [
      { v: "0", label: "構造化・合意形成", body: "曖昧な課題を構造化し、関係者の合意形成まで伴走します。" },
      { v: "1", label: "プロセス改善", body: "手戻りを減らす仕組みづくりと、測定できる改善が得意です。" },
      { v: "2", label: "ステークホルダ調整", body: "利害が分かれる状況でも、前に進む対話を設計するのが得意です。" },
    ],
  },
  decisionStyle: {
    title: "意思決めの進め方",
    options: [
      { v: "0", label: "叩き台駆動", body: "まず叩き台を早く出し、そこから議論を収束させるのが早いです。" },
      { v: "1", label: "根拠リンク派", body: "意思決めは根拠リンク付きだと安心し、判断が早くなります。" },
      { v: "2", label: "小さく試す", body: "不確実性が高いときは小さく試して学び、段階的に確度を上げます。" },
    ],
  },
  focusHours: {
    title: "集中しやすい時間帯",
    options: [
      { v: "0", label: "午前深掘り", body: "午前は深掘り作業向き。午後は会議やレビューが多くても問題ありません。" },
      { v: "1", label: "午後ブロック", body: "午後にまとまった集中ブロックを確保すると、成果が出やすいです。" },
      { v: "2", label: "非同期中心", body: "非同期で整理し、必要なときだけ短時間で同期します。" },
    ],
  },
  channels: {
    title: "連絡チャネルの優先順位",
    options: [
      { v: "0", label: "チャット優先", body: "基本はチャット→必要なら短い通話→合意事項はメールで残します。" },
      { v: "1", label: "メール優先", body: "記録が必要なものはメール、急ぎはチャットで一次連絡します。" },
      { v: "2", label: "カレンダー駆動", body: "日程調整はカレンダー、議論は会議、確定はチャット／メールです。" },
    ],
  },
  responseSLA: {
    title: "返信の目安",
    options: [
      { v: "0", label: "24h一次返信", body: "通常は24時間以内に一次返信します。緊急はチャットで@ください。" },
      { v: "1", label: "営業日内", body: "営業日ベースで返します。週末は原則翌営業日扱いです。" },
      { v: "2", label: "週次まとめ", body: "軽微なものは週次でまとめ返信、緊急のみ即日対応します。" },
    ],
  },
  meetingPrefs: {
    title: "会議の好み",
    options: [
      { v: "0", label: "アジェンダ必須", body: "アジェンダと目的が明確だと参加の質が上がります。" },
      { v: "1", label: "結論先出し", body: "結論→理由→次アクションの順が読みやすく好みです。" },
      { v: "2", label: "短時間・少人数", body: "短時間・少人数のほうが意思決めが速いです。" },
    ],
  },
  commStyle: {
    title: "文章 vs 口頭",
    options: [
      { v: "0", label: "複雑は口頭", body: "複雑系は口頭で擦り合わせ、合意事項は文章で残したいです。" },
      { v: "1", label: "文章で整理", body: "まず文章で整理してから会話に入ると、認識が揃いやすいです。" },
      { v: "2", label: "図・表が得意", body: "図・表があると理解が早いです（ホワイトボード派）。" },
    ],
  },
  feedbackGood: {
    title: "受け取りやすいフィードバック",
    options: [
      { v: "0", label: "具体例つき", body: "具体例つきだと改善が早く、行動に落とし込みやすいです。" },
      { v: "1", label: "意図の共有", body: "意図（なぜそう思うか）が分かると受け取りやすいです。" },
      { v: "2", label: "小さく頻度", body: "大きな一発より、小さく頻度があるほうが伸びます。" },
    ],
  },
  feedbackStress: {
    title: "ストレスになりやすい伝え方",
    options: [
      { v: "0", label: "人格に見える表現", body: "人格に見える表現は避けてほしいです（行動・事実ベースが助かります）。" },
      { v: "1", label: "曖昧な一言", body: "「なんとなく」だけだと改善点が掴みにくく消耗します。" },
      { v: "2", label: "公開での指摘", body: "指摘は基本1on1希望です（公開は例外）。" },
    ],
  },
  misunderstood: {
    title: "誤解されがちな点",
    options: [
      { v: "0", label: "沈黙＝不満ではない", body: "沈黙＝不満ではなく、整理・考え込みのことが多いです。" },
      { v: "1", label: "質問が多い", body: "質問が多いのは詰めではなく、手戻りを減らすためです。" },
      { v: "2", label: "スピード感", body: "速く見えることがありますが、重要論点は抜けないよう確認します。" },
    ],
  },
  helpMe: {
    title: "こうされると助かる",
    options: [
      { v: "0", label: "背景1行", body: "背景目的を1行で共有してくれると、着手が速いです。" },
      { v: "1", label: "期待アウトカム", body: "期待アウトカム（何ができればOKか）があると迷いません。" },
      { v: "2", label: "期限の明示", body: "期限と優先度が分かると、調整がしやすいです。" },
    ],
  },
  avoidRoles: {
    title: "苦手／避けたい役割",
    options: [
      { v: "0", label: "細かい手作業連続", body: "細かい手作業の連続は得意ではありません（自動化したい）。" },
      { v: "1", label: "単独常駐", body: "単独常駐のオペレーション中心は避けたいです。" },
      { v: "2", label: "目的不明の長会議", body: "目的が曖昧な長会議はエネルギー消費が大きいです。" },
    ],
  },
  energizers: {
    title: "エネルギーが上がること",
    options: [
      { v: "0", label: "ユーザー課題の改善", body: "ユーザー課題が数字で改善したときに一番やる気が出ます。" },
      { v: "1", label: "チームで越える壁", body: "チームで難しい壁を越えた瞬間が好きです。" },
      { v: "2", label: "学びの循環", body: "学びを言語化して共有できる文化があると伸びます。" },
    ],
  },
  stressful: {
    title: "消耗しやすいこと",
    options: [
      { v: "0", label: "目的なき多人数会議", body: "目的が曖昧なままの多人数会議は消耗しやすいです。" },
      { v: "1", label: "頻繁な文脈切替", body: "短時間で文脈切替が続くと疲れやすいです。" },
      { v: "2", label: "曖昧な優先度", body: "優先度が曖昧なまま並走が増えるとストレスが上がります。" },
    ],
  },
  boundaries: {
    title: "連絡・時間の境界",
    options: [
      { v: "0", label: "休日は非同期", body: "休日は原則非同期。緊急はチャットのみでお願いします。" },
      { v: "1", label: "夜は翌朝", body: "夜間の連絡は翌朝まとめて見ることが多いです（緊急は別ルール）。" },
      { v: "2", label: "家族時間を確保", body: "家族時間を確保する日は、返信が遅れることがあります。" },
    ],
  },
  footerNote: {
    title: "フッター注記",
    options: [
      { v: "0", label: "社内共有", body: "共有範囲：社内／更新：四半期ごと" },
      { v: "1", label: "チーム限定", body: "共有範囲：チーム内／更新：プロジェクト完了まで" },
      { v: "2", label: "自己管理", body: "この文書は自己管理用。配布は都度確認します。" },
    ],
  },
};

const PRESET_FIELD_IDS = /** @type {string[]} */ (Object.keys(PRESETS));

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

/** @param {string} fieldId */
function mergePresetField(fieldId, freeRaw) {
  const pr = PRESETS[fieldId];
  if (!pr) return (freeRaw || "").trim();
  const checked = /** @type {HTMLInputElement | null} */ (
    document.querySelector(`input[name="${fieldId}_preset"]:checked`)
  );
  let presetBody = "";
  if (checked) {
    const opt = pr.options.find((o) => o.v === checked.value);
    if (opt) presetBody = String(opt.body || "").trim();
  }
  const free = (freeRaw || "").trim();
  if (presetBody && free) return `${presetBody}\n\n${free}`;
  return presetBody || free;
}

function readEffectiveForm() {
  const raw = readForm();
  /** @type {Record<string, string>} */
  const d = { ...raw };
  for (const id of PRESET_FIELD_IDS) {
    d[id] = mergePresetField(id, raw[id]);
  }
  return d;
}

function collectPersistedState() {
  const data = readForm();
  for (const id of PRESET_FIELD_IDS) {
    const c = /** @type {HTMLInputElement | null} */ (document.querySelector(`input[name="${id}_preset"]:checked`));
    data[`${id}_preset`] = c ? c.value : "";
  }
  return data;
}

function injectPresets() {
  for (const el of document.querySelectorAll("[data-preset]")) {
    const id = el.getAttribute("data-preset");
    if (!id || !PRESETS[id] || el.querySelector(".preset-host")) continue;
    const pr = PRESETS[id];
    const label = el.querySelector("label");
    const control = el.querySelector("textarea, input[type=text]");
    if (!label || !control) continue;

    const host = document.createElement("div");
    host.className = "preset-host";

    const grid = document.createElement("div");
    grid.className = "preset-grid";
    grid.setAttribute("role", "radiogroup");
    grid.setAttribute("aria-label", `${pr.title}の候補から選択`);

    for (const opt of pr.options) {
      const card = document.createElement("label");
      card.className = "preset-card";
      const inp = document.createElement("input");
      inp.type = "radio";
      inp.name = `${id}_preset`;
      inp.value = opt.v;
      const t = document.createElement("span");
      t.className = "preset-card__title";
      t.textContent = opt.label;
      const b = document.createElement("span");
      b.className = "preset-card__body";
      b.textContent = opt.body;
      card.append(inp, t, b);
      grid.appendChild(card);
    }

    const hint = document.createElement("p");
    hint.className = "preset-free-hint";
    hint.textContent = "自由記述（任意）：候補に足したい具体例・例外・条件を追記できます。";

    host.append(grid, hint);
    label.insertAdjacentElement("afterend", host);
  }
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
  const d = readEffectiveForm();

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
    localStorage.setItem(STORAGE_KEY, JSON.stringify(collectPersistedState()));
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
    for (const id of PRESET_FIELD_IDS) {
      const key = `${id}_preset`;
      if (!(key in data)) continue;
      const v = data[key];
      const radio = document.querySelector(`input[name="${id}_preset"][value="${String(v)}"]`);
      if (radio) /** @type {HTMLInputElement} */ (radio).checked = true;
    }
  } catch {
    // ignore
  }
}

function validateRequired() {
  const d = readEffectiveForm();
  const missing = [];
  if (!d.displayName) missing.push("表示名");
  if (!d.department) missing.push("所属");
  if (!d.strengths) missing.push("得意な貢献・強み（候補の選択または自由記述）");
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

  const d = readEffectiveForm();
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
  document.querySelectorAll('input[type="radio"][name$="_preset"]').forEach((r) => {
    /** @type {HTMLInputElement} */ (r).checked = false;
  });
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // ignore
  }
  updatePreview();
  setStatus("入力をクリアしました。");
}

function wire() {
  const onChange = () => {
    updatePreview();
    saveLocal();
    setStatus("");
  };

  for (const id of fields) {
    const el = document.getElementById(id);
    if (!el) continue;
    el.addEventListener("input", onChange);
  }

  document.querySelectorAll('input[type="radio"][name$="_preset"]').forEach((r) => {
    r.addEventListener("change", onChange);
  });

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

injectPresets();
loadLocal();
wire();
updatePreview();
setupPreviewScaling();
