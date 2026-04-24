// server.js
// ✅ FULL CODE (FIX: tiêu đề PHẦN đúng vị trí như file Word gốc + GIỮ BẢNG trong Word)
// - Không lệch khi mỗi PHẦN reset "Câu 1."
// - Server trả thêm `blocks` đã trộn (section + question) đúng thứ tự để frontend render chuẩn.
// - ✅ NEW: Giữ được bảng <w:tbl> và nội dung trong bảng (kể cả underline + token math/img)
//
// ✅ FIX ẢNH BỊ THIẾU (Câu 7, Câu 11):
// - Bắt thêm <a:blip ...> (không tự đóng) ngoài <a:blip .../>
// - Bắt thêm cả r:link (một số doc dùng link thay vì embed)
//
// ✅ FIX MẤT CĂN THỨC (MathType OLE):
// - extractMathMLFromOleScan() bắt cả <math> và <m:math>
// - normalize MathML: strip prefix m:, menclose radical -> msqrt, mo √ -> msqrt
// - tokenize msqrt -> token, convert, rebuild \sqrt{...} (radical-safe)
// - hard wrap nếu MathML có căn mà LaTeX không có \sqrt
//
// Chạy: node server.js
// Yêu cầu: inkscape (convert emf/wmf), ruby + mt2mml_v2.rb (ưu tiên) / mt2mml.rb (fallback)
// npm i express multer unzipper cors mathml-to-latex

import express from "express";
import multer from "multer";
import unzipper from "unzipper";
import cors from "cors";
import fs from "fs";
import os from "os";
import path from "path";
import { execFile, execFileSync } from "child_process";
import { MathMLToLaTeX } from "mathml-to-latex";

const app = express();
app.use(cors());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

/* ================= Helpers ================= */

function parseRels(relsXml) {
  const map = new Map();
  const re =
    /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/>/g;
  let m;
  while ((m = re.exec(relsXml))) map.set(m[1], m[2]);
  return map;
}

function normalizeTargetToWordPath(target) {
  let t = (target || "").replace(/^(\.\.\/)+/, "");
  if (!t.startsWith("word/")) t = `word/${t}`;
  return t;
}

function extOf(p = "") {
  return p.split(".").pop()?.toLowerCase() || "";
}

function guessMimeFromFilename(filename = "") {
  const ext = extOf(filename);
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "bmp") return "image/bmp";
  if (ext === "webp") return "image/webp";
  if (ext === "svg") return "image/svg+xml";
  if (ext === "emf") return "image/emf";
  if (ext === "wmf") return "image/wmf";
  return "application/octet-stream";
}

function decodeXmlEntities(s = "") {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) =>
      String.fromCharCode(parseInt(h, 16))
    );
}

async function getZipEntryBuffer(zipFiles, p) {
  const f = zipFiles.find((x) => x.path === p);
  if (!f) return null;
  return await f.buffer();
}

/* ================= Inkscape Convert EMF/WMF -> PNG ================= */

function inkscapeConvertToPng(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    execFile(
      "inkscape",
      [
        inputPath,
        "--export-type=png",
        `--export-filename=${outputPath}`,
        "--export-area-drawing",
        "--export-background-opacity=0",
      ],
      { timeout: 30000 },
      (err, stdout, stderr) => {
        if (err) return reject(new Error(stderr || err.message));
        resolve(true);
      }
    );
  });
}

async function maybeConvertEmfWmfToPng(buf, filename) {
  const ext = extOf(filename);
  if (ext !== "emf" && ext !== "wmf") return null;

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "mtype-"));
  const inPath = path.join(tmpDir, `in.${ext}`);
  const outPath = path.join(tmpDir, "out.png");

  try {
    fs.writeFileSync(inPath, buf);
    await inkscapeConvertToPng(inPath, outPath);
    return fs.readFileSync(outPath);
  } finally {
    try {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    } catch {}
  }
}

/* ================= MathType OLE -> MathML -> LaTeX ================= */

function extractMathMLFromOleScan(buf) {
  const tryExtract = (s) => {
    if (!s) return null;

    // bắt cả <math ...> và <m:math ...>
    let i = s.indexOf("<math");
    let close = "</math>";
    if (i === -1) {
      i = s.indexOf("<m:math");
      close = "</m:math>";
    }
    if (i === -1) return null;

    const j = s.indexOf(close, i);
    if (j !== -1) return s.slice(i, j + close.length);

    // fallback: nếu open là <m:math> nhưng close lại </math> (hiếm)
    const j2 = s.indexOf("</math>", i);
    if (j2 !== -1) return s.slice(i, j2 + 7);

    return null;
  };

  // utf8
  let out = tryExtract(buf.toString("utf8"));
  if (out) return out;

  // utf16le
  out = tryExtract(buf.toString("utf16le"));
  if (out) return out;

  return null;
}

function rubyOleToMathML(oleBuf) {
  return new Promise((resolve, reject) => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "ole-"));
    const inPath = path.join(tmpDir, "oleObject.bin");
    fs.writeFileSync(inPath, oleBuf);

    // ✅ Ưu tiên mt2mml_v2.rb nếu có (MTEF→MathML thật), fallback mt2mml.rb
    const script = fs.existsSync("mt2mml_v2.rb") ? "mt2mml_v2.rb" : "mt2mml.rb";

    execFile(
      "ruby",
      [script, inPath],
      { timeout: 30000, maxBuffer: 20 * 1024 * 1024 },
      (err, stdout, stderr) => {
        try {
          fs.rmSync(tmpDir, { recursive: true, force: true });
        } catch {}
        if (err) return reject(new Error(stderr || err.message));
        resolve(String(stdout || "").trim());
      }
    );
  });
}

/* ================== LATEX POSTPROCESS ================== */

const SQRT_MATHML_RE = /(msqrt|mroot|√|&#8730;|&#x221a;|&#x221A;|&radic;)/i;

/** ✅ normalize MathML trước khi convert (cứu căn + prefix m:) */
function normalizeMathMLForConvert(mml) {
  let s = String(mml || "");

  // 1) strip prefix m: (mathml-to-latex hay fail nếu giữ m:)
  s = s.replace(/<\/?m:/g, "<");
  // strip prefix kiểu khác nếu có (hiếm)
  s = s.replace(/<\/?[a-zA-Z0-9]+:/g, (tag) =>
    tag
      .replace(/^</, "<")
      .replace(/^<\/?[a-zA-Z0-9]+:/, (x) =>
        x.replace(/^<\//, "</").replace(/^</, "<")
      )
  );

  // 2) menclose radical -> msqrt (thủ phạm “mất căn” phổ biến)
  const reRad =
    /<menclose\b[^>]*\bnotation\s*=\s*"radical"[^>]*>([\s\S]*?)<\/menclose>/gi;
  while (reRad.test(s)) s = s.replace(reRad, "<msqrt>$1</msqrt>");

  // 3) chuẩn hoá entity √ nếu có
  s = s.replace(/&radic;|&#8730;|&#x221a;|&#x221A;/g, "√");

  // 4) mo √ ... -> msqrt (nhiều file gặp dạng này)
  const reMoSqrt =
    /<mo>\s*√\s*<\/mo>\s*(<mrow>[\s\S]*?<\/mrow>|<mi>[\s\S]*?<\/mi>|<mn>[\s\S]*?<\/mn>|<mfenced[\s\S]*?<\/mfenced>)/gi;
  while (reMoSqrt.test(s)) s = s.replace(reMoSqrt, "<msqrt>$1</msqrt>");

  return s;
}

/** ✅ token hóa msqrt để converter có drop vẫn rebuild được \sqrt{...} */
function tokenizeMsqrtBlocks(mathml) {
  const s = String(mathml || "");
  const re = /<\/?msqrt\b[^>]*>/gi;

  const stack = [];
  const blocks = []; // match pairs

  let m;
  while ((m = re.exec(s)) !== null) {
    const tag = m[0];
    const isClose = tag.startsWith("</");
    if (!isClose) {
      stack.push({ openStart: m.index, openEnd: re.lastIndex });
    } else {
      const open = stack.pop();
      if (!open) continue;
      blocks.push({
        openStart: open.openStart,
        openEnd: open.openEnd,
        closeStart: m.index,
        closeEnd: re.lastIndex,
      });
    }
  }

  if (!blocks.length) return { out: s, tokens: [] };

  // replace from back to front to keep indices stable
  blocks.sort((a, b) => b.openStart - a.openStart);

  let out = s;
  const tokens = [];
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    const token = `SQRTTOKEN${i + 1}X`; // ✅ tránh underscore để ít bị bẻ
    const inner = out.slice(b.openEnd, b.closeStart);
    tokens.push({ token, inner });

    out = out.slice(0, b.openStart) + `<mi>${token}</mi>` + out.slice(b.closeEnd);
  }

  return { out, tokens };
}

function sanitizeLatexStrict(latex) {
  if (!latex) return latex;
  latex = String(latex).replace(/\s+/g, " ").trim();

  latex = latex
    .replace(
      /\\left(?!\s*(\(|\[|\\\{|\\langle|\\vert|\\\||\||\.))/g,
      ""
    )
    .replace(
      /\\right(?!\s*(\)|\]|\\\}|\\rangle|\\vert|\\\||\||\.))/g,
      ""
    );

  const tokens = latex.match(/\\left\b|\\right\b/g) || [];
  let bal = 0;
  let broken = false;
  for (const t of tokens) {
    if (t === "\\left") bal++;
    else {
      if (bal === 0) {
        broken = true;
        break;
      }
      bal--;
    }
  }
  if (bal !== 0) broken = true;

  if (broken) latex = latex.replace(/\\left\s*/g, "").replace(/\\right\s*/g, "");
  return latex;
}

function fixSetBracesHard(latex) {
  let s = String(latex || "");

  s = s.replace(
    /\\underset\s*\{([^}]*)\}\s*\{\s*l\s*i\s*m\s*\}/gi,
    "\\underset{$1}{\\lim}"
  );
  s = s.replace(/\b(l)\s+(i)\s+(m)\b/gi, "lim");
  s = s.replace(/(^|[^A-Za-z\\])lim([^A-Za-z]|$)/g, "$1\\lim$2");

  s = s.replace(/\\arrow\b/g, "\\rightarrow");
  s = s.replace(/\bxarrow\b/g, "x\\rightarrow");
  s = s.replace(/\\xarrow\b/g, "\\xrightarrow");

  s = s.replace(/\\\{\s*\./g, "\\{");
  s = s.replace(/\.\s*\\\}/g, "\\}");
  s = s.replace(/\\\}\s*\./g, "\\}");

  s = s.replace(/\\mathbb\{([A-Za-z])\\\}/g, "\\mathbb{$1}");
  s = s.replace(/\\mathbb\{([A-Za-z])\}\s*\.\s*\}/g, "\\mathbb{$1}}");

  s = s.replace(/\\backslash\s*{(?!\\)/g, "\\backslash \\{");
  s = s.replace(/\\setminus\s*{(?!\\)/g, "\\setminus \\{");

  if (
    (s.includes("\\backslash \\{") || s.includes("\\setminus \\{")) &&
    !s.includes("\\}")
  ) {
    s = s.replace(/\}\s*$/g, "").trim() + "\\}";
  }

  s = s.replace(/\\\}\s*([,.;:])/g, "\\}$1");

  s = s.replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, (m, a, b) => {
    const bb = String(b).replace(/(\d)\s+(\d)/g, "$1$2");
    return `\\frac{${a}}{${bb}}`;
  });

  s = s.replace(/\s+/g, " ").trim();
  return s;
}

function restoreArrowAndCoreCommands(latex) {
  let s = String(latex || "");
  s = s.replace(/\s+/g, " ").trim();
  s = s.replace(/\b([A-Za-z])\s+arrow\b/g, "$1 \\to");
  s = s.replace(/\brightarrow\b/g, "\\rightarrow");
  s = s.replace(/\barrow\b/g, "\\rightarrow");
  s = s.replace(/(^|[^A-Za-z\\])to([^A-Za-z]|$)/g, "$1\\to$2");
  return s.replace(/\s+/g, " ").trim();
}

function fixPiecewiseFunction(latex) {
  let s = String(latex || "");

  s = s.replace(/\(\.\s+/g, "(");
  s = s.replace(/\s+\.\)/g, ")");
  s = s.replace(/\[\.\s+/g, "[");
  s = s.replace(/\s+\.\]/g, "]");

  const piecewiseMatch = s.match(/(?<!\\)\{\.\s+/);
  if (piecewiseMatch) {
    const startIdx = piecewiseMatch.index;
    const contentStart = startIdx + piecewiseMatch[0].length;

    let braceCount = 1;
    let endIdx = contentStart;
    let foundEnd = false;

    for (let i = contentStart; i < s.length; i++) {
      const ch = s[i];
      const prevCh = i > 0 ? s[i - 1] : "";
      if (prevCh === "\\") continue;

      if (ch === "{") braceCount++;
      else if (ch === "}") {
        braceCount--;
        if (braceCount === 0) {
          endIdx = i;
          foundEnd = true;
          break;
        }
      }
    }

    if (!foundEnd) endIdx = s.length;

    let content = s.slice(contentStart, endIdx).trim();
    content = content.replace(/\s+\.\s*$/, "");
    content = content.replace(/\s+\\\s+(?=\d)/g, " \\\\ ");

    const before = s.slice(0, startIdx);
    const after = foundEnd ? s.slice(endIdx + 1) : "";
    s = before + `\\begin{cases} ${content} \\end{cases}` + after;
  }

  return s;
}

function fixSqrtLatex(latex, mathmlMaybe = "") {
  let s = String(latex || "");

  s = s.replace(/√\s*\(\s*([\s\S]*?)\s*\)/g, "\\sqrt{$1}");
  s = s.replace(/√\s*([A-Za-z0-9]+)\b/g, "\\sqrt{$1}");

  if (SQRT_MATHML_RE.test(String(mathmlMaybe || ""))) {
    const hasSqrt = /\\sqrt\b|\\root\b/.test(s);
    if (!hasSqrt && s) {
      s = s.replace(/\bradic\b/gi, "\\sqrt{}");
    }
  }

  return s;
}

function postProcessLatex(latex, mathmlMaybe = "") {
  let s = latex || "";
  s = sanitizeLatexStrict(s);
  s = fixSetBracesHard(s);
  s = restoreArrowAndCoreCommands(s);
  s = fixPiecewiseFunction(s);
  s = fixSqrtLatex(s, mathmlMaybe);
  return String(s || "").replace(/\s+/g, " ").trim();
}

/** ✅ Radical-safe: tokenize msqrt -> convert -> rebuild sqrt */
function mathmlToLatexSafe(mml, _depth = 0) {
  try {
    if (!mml) return "";
    let m = String(mml);
    if (!m.includes("<math")) return "";

    m = normalizeMathMLForConvert(m);

    const tok = tokenizeMsqrtBlocks(m);
    const mTok = tok.out;

    let latex0 = (MathMLToLaTeX.convert(mTok) || "").trim();
    latex0 = postProcessLatex(latex0, mTok);

    if (!tok.tokens.length) {
      // hard wrap nếu MathML có căn mà latex không có sqrt
      if (SQRT_MATHML_RE.test(m) && latex0 && !/\\sqrt\b|\\root\b/.test(latex0)) {
        return `\\sqrt{${latex0}}`;
      }
      return latex0;
    }

    let out = latex0;

    const depth = Number(_depth || 0);
    const canRecurse = depth < 4;

    for (const t of tok.tokens) {
      let innerLatex = "";
      const innerMath = `<math>${t.inner}</math>`;

      if (canRecurse) {
        innerLatex = mathmlToLatexSafe(innerMath, depth + 1);
      } else {
        innerLatex = (MathMLToLaTeX.convert(normalizeMathMLForConvert(innerMath)) || "").trim();
        innerLatex = postProcessLatex(innerLatex, innerMath);
      }

      innerLatex = innerLatex || "";
      const repl = `\\sqrt{${innerLatex}}`;

      const reTok = new RegExp(t.token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
      out = out.replace(reTok, repl);
    }

    out = String(out || "").replace(/\s+/g, " ").trim();

    // ✅ HARD FIX cuối: nếu MathML có căn mà latex vẫn không có \sqrt
    if (SQRT_MATHML_RE.test(m) && out && !/\\sqrt\b|\\root\b/.test(out)) {
      out = `\\sqrt{${out}}`;
    }

    return out;
  } catch {
    return "";
  }
}

/* ================= MathType FIRST ================= */

async function tokenizeMathTypeOleFirst(docXml, rels, zipFiles, images) {
  let idx = 0;
  const found = {};
  const OBJECT_RE = /<w:object[\s\S]*?<\/w:object>/g;

  docXml = docXml.replace(OBJECT_RE, (block) => {
    const ole = block.match(/<o:OLEObject\b[^>]*\br:id="([^"]+)"/);
    if (!ole) return block;

    const oleRid = ole[1];
    const oleTarget = rels.get(oleRid);
    if (!oleTarget) return block;

    const vmlRid = block.match(/<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/);
    // ✅ FIX preview: bắt cả r:embed hoặc r:link và tag có thể / > hoặc />
    const blipRid = block.match(/<a:blip\b[^>]*\br:(?:embed|link)="([^"]+)"[^>]*\/?>/);

    const previewRid = vmlRid?.[1] || blipRid?.[1] || null;

    const key = `mathtype_${++idx}`;
    found[key] = { oleTarget, previewRid };
    return `[!m:$${key}$]`;
  });

  const latexMap = {};

  await Promise.all(
    Object.entries(found).map(async ([key, info]) => {
      const oleFull = normalizeTargetToWordPath(info.oleTarget);
      const oleBuf = await getZipEntryBuffer(zipFiles, oleFull);

      let mml = "";
      if (oleBuf) mml = extractMathMLFromOleScan(oleBuf) || "";

      if (!mml && oleBuf) {
        try {
          mml = await rubyOleToMathML(oleBuf);
        } catch {
          mml = "";
        }
      }

      // ✅ normalize trước convert (giúp cả trường hợp m:msqrt)
      if (mml) mml = normalizeMathMLForConvert(mml);

      const latex = mml ? mathmlToLatexSafe(mml) : "";
      if (latex) {
        latexMap[key] = latex;
        return;
      }

      // fallback preview image
      if (info.previewRid) {
        const t = rels.get(info.previewRid);
        if (t) {
          const imgFull = normalizeTargetToWordPath(t);
          const imgBuf = await getZipEntryBuffer(zipFiles, imgFull);
          if (imgBuf) {
            const mime = guessMimeFromFilename(imgFull);
            if (mime === "image/emf" || mime === "image/wmf") {
              try {
                const pngBuf = await maybeConvertEmfWmfToPng(imgBuf, imgFull);
                if (pngBuf) {
                  images[`fallback_${key}`] = `data:image/png;base64,${pngBuf.toString(
                    "base64"
                  )}`;
                  latexMap[key] = "";
                  return;
                }
              } catch {}
            }
            images[`fallback_${key}`] = `data:${mime};base64,${imgBuf.toString(
              "base64"
            )}`;
          }
        }
      }

      latexMap[key] = "";
    })
  );

  return { outXml: docXml, latexMap };
}

/* ================= Images AFTER MathType ================= */

async function tokenizeImagesAfter(docXml, rels, zipFiles) {
  let idx = 0;
  const imgMap = {};
  const jobs = [];

  const schedule = (rid, key) => {
    const target = rels.get(rid);
    if (!target) return;
    const full = normalizeTargetToWordPath(target);

    jobs.push(
      (async () => {
        const buf = await getZipEntryBuffer(zipFiles, full);
        if (!buf) return;

        const mime = guessMimeFromFilename(full);
        if (mime === "image/emf" || mime === "image/wmf") {
          try {
            const pngBuf = await maybeConvertEmfWmfToPng(buf, full);
            if (pngBuf) {
              imgMap[key] = `data:image/png;base64,${pngBuf.toString("base64")}`;
              return;
            }
          } catch {}
        }
        imgMap[key] = `data:${mime};base64,${buf.toString("base64")}`;
      })()
    );
  };

  // ✅ FIX DUY NHẤT: bắt cả <a:blip .../> và <a:blip ...> + cả r:embed và r:link
  docXml = docXml.replace(
    /<a:blip\b[^>]*\br:(?:embed|link)="([^"]+)"[^>]*\/?>/g,
    (m, rid) => {
      const key = `img_${++idx}`;
      schedule(rid, key);
      return `[!img:$${key}$]`;
    }
  );

  docXml = docXml.replace(
    /<v:imagedata\b[^>]*\br:id="([^"]+)"[^>]*\/>/g,
    (m, rid) => {
      const key = `img_${++idx}`;
      schedule(rid, key);
      return `[!img:$${key}$]`;
    }
  );

  await Promise.all(jobs);
  return { outXml: docXml, imgMap };
}

/* ================= ✅ TABLE SUPPORT (GIỮ BẢNG + NỘI DUNG TRONG Ô) ================= */

function convertRunsToHtml(fragmentXml) {
  let frag = String(fragmentXml || "");

  frag = frag
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n");

  frag = frag.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
    const hasU =
      /<w:u\b[^>]*\/>/.test(run) &&
      !/<w:u\b[^>]*w:val="none"[^>]*\/>/.test(run);

    let inner = run.replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g, "");
    inner = inner.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (_, t) => t ?? "");
    inner = inner.replace(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g,
      (_, t) => t ?? ""
    );

    inner = inner.replace(/<[^>]+>/g, "");
    if (!inner) return "";
    return hasU ? `<u>${inner}</u>` : inner;
  });

  frag = frag.replace(/<(?!\/?u\b)[^>]+>/g, "");
  frag = decodeXmlEntities(frag);

  frag = frag.replace(/\r/g, "");
  frag = frag.replace(/[ \t]+\n/g, "\n").trim();
  return frag;
}

function convertParagraphsToHtml(parXml) {
  let p = String(parXml || "");
  p = convertRunsToHtml(p);
  return p;
}

function wordTableXmlToHtmlTable(tblXml) {
  const tbl = String(tblXml || "");
  const rows = tbl.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || [];

  let html = `<table class="doc-table">`;

  for (const tr of rows) {
    html += `<tr>`;
    const cells = tr.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || [];

    for (const tc of cells) {
      const ps = tc.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
      const parts = ps.map(convertParagraphsToHtml).filter(Boolean);
      const cellHtml = parts.join("<br/>").trim();
      html += `<td>${cellHtml || ""}</td>`;
    }

    html += `</tr>`;
  }

  html += `</table>`;
  return html;
}

/* ================= Text (GIỮ token + underline + ✅ TABLE) ================= */

function wordXmlToTextKeepTokens(docXml) {
  let x = String(docXml || "");

  x = x.replace(/\[!m:\$\$?(.*?)\$\$?\]/g, "___MATH_TOKEN___$1___END___");
  x = x.replace(/\[!img:\$\$?(.*?)\$\$?\]/g, "___IMG_TOKEN___$1___END___");

  const tableMap = {};
  let tableIdx = 0;

  x = x.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, (tblBlock) => {
    const key = `___TABLE_TOKEN___${++tableIdx}___END___`;
    tableMap[key] = wordTableXmlToHtmlTable(tblBlock);
    return key;
  });

  x = x
    .replace(/<w:tab\s*\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n");

  x = x.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
    const hasU =
      /<w:u\b[^>]*\/>/.test(run) &&
      !/<w:u\b[^>]*w:val="none"[^>]*\/>/.test(run);

    let inner = run.replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/g, "");
    inner = inner.replace(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, (_, t) => t ?? "");
    inner = inner.replace(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g,
      (_, t) => t ?? ""
    );

    inner = inner.replace(/<[^>]+>/g, "");
    if (!inner) return "";
    return hasU ? `<u>${inner}</u>` : inner;
  });

  x = x.replace(/<(?!\/?(u|table|tr|td|br)\b)[^>]+>/g, "");

  for (const [k, v] of Object.entries(tableMap)) {
    x = x.split(k).join(v);
  }

  x = x
    .replace(/___MATH_TOKEN___(.*?)___END___/g, "[!m:$$$1$$]")
    .replace(/___IMG_TOKEN___(.*?)___END___/g, "[!img:$$$1$$]");

  x = decodeXmlEntities(x)
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return x;
}

/* ================= SECTION TITLES (PHẦN ...) ================= */

function extractSectionTitles(rawText) {
  const text = String(rawText || "").replace(/\r/g, "");

  const qRe = /(^|\n)\s*Câu\s+(\d+)\./gi;
  const qAnchors = [];
  let qm;
  while ((qm = qRe.exec(text)) !== null) {
    qAnchors.push({
      idx: qm.index + (qm[1] ? qm[1].length : 0),
      no: Number(qm[2]),
    });
  }

  const sRe =
    /(^|\n)\s*(?:[-•–]\s*)?PHẦN\s+([0-9]+|[IVXLCDM]+)\s*[\.\:\-]?\s*/gi;

  const sections = [];
  let sm;
  while ((sm = sRe.exec(text)) !== null) {
    const startChar = sm.index + (sm[1] ? sm[1].length : 0);
    sections.push({
      title: "",
      order: sections.length + 1,
      startChar,
      endChar: null,
      firstQuestionNo: null,
      questionCount: 0,
      questionIndexStart: null,
      questionIndexEnd: null,
      _phanLabel: sm[2],
    });
  }

  for (let i = 0; i < sections.length; i++) {
    sections[i].endChar =
      i + 1 < sections.length ? sections[i + 1].startChar : text.length;
  }

  const normalizeTitle = (s) =>
    String(s || "")
      .replace(/\u00A0/g, " ")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{2,}/g, "\n")
      .trim()
      .replace(/\s*\n\s*/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  for (const sec of sections) {
    const startIdx = qAnchors.findIndex(
      (q) => q.idx >= sec.startChar && q.idx < sec.endChar
    );

    const firstQIdx = startIdx === -1 ? sec.endChar : qAnchors[startIdx].idx;

    let titleBlock = text.slice(sec.startChar, firstQIdx);
    titleBlock = titleBlock.replace(/^\s+/g, "");

    const cut = titleBlock.search(/(^|\n)\s*Câu\s+\d+\./i);
    if (cut >= 0) titleBlock = titleBlock.slice(0, cut);

    sec.title = normalizeTitle(titleBlock);

    if (startIdx === -1) continue;

    let endIdx = qAnchors.length;
    for (let k = startIdx; k < qAnchors.length; k++) {
      if (qAnchors[k].idx >= sec.endChar) {
        endIdx = k;
        break;
      }
    }

    sec.questionIndexStart = startIdx;
    sec.questionIndexEnd = endIdx;
    sec.questionCount = endIdx - startIdx;
    sec.firstQuestionNo = qAnchors[startIdx]?.no ?? null;
  }

  return sections;
}

/* ================== EXAM PARSER (GIỮ NGUYÊN) ================== */

function stripTagsToPlain(s) {
  return String(s || "")
    .replace(/<u[^>]*>/gi, "")
    .replace(/<\/u>/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function detectHasMCQ(plain) {
  const marks = plain.match(/\b[ABCD]\./g) || [];
  return new Set(marks).size >= 2;
}

function detectHasTF4(plain) {
  const marks = plain.match(/\b[a-d]\)/gi) || [];
  return new Set(marks.map((x) => x.toLowerCase())).size >= 2;
}

function extractUnderlinedKeys(blockText) {
  const keys = { mcq: null, tf: [] };
  const s = String(blockText || "");

  let m =
    s.match(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./i) ||
    s.match(/<u[^>]*>\s*([A-D])\.\s*<\/u>/i);
  if (m) keys.mcq = m[1].toUpperCase();

  let mm;
  const reTF1 = /<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi;
  while ((mm = reTF1.exec(s)) !== null) keys.tf.push(mm[1].toLowerCase());

  const reTF2 = /<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi;
  while ((mm = reTF2.exec(s)) !== null) keys.tf.push(mm[1].toLowerCase());

  keys.tf = [...new Set(keys.tf)];
  return keys;
}

function normalizeUnderlinedMarkersForSplit(s) {
  let x = String(s || "");
  x = x.replace(/<u[^>]*>\s*([A-D])\s*<\/u>\s*\./gi, "$1.");
  x = x.replace(/<u[^>]*>\s*([A-D])\.\s*<\/u>/gi, "$1.");
  x = x.replace(/<u[^>]*>\s*([a-d])\s*\)\s*<\/u>/gi, "$1)");
  x = x.replace(/<u[^>]*>\s*([a-d])\s*<\/u>\s*\)/gi, "$1)");
  return x;
}

function findSolutionMarkerIndex(text, fromIndex = 0) {
  const s = String(text || "");
  const re = /(Lời\s*giải|Giải\s*chi\s*tiết|Hướng\s*dẫn\s*giải)/i;
  const sub = s.slice(fromIndex);
  const m = re.exec(sub);
  if (!m) return -1;
  return fromIndex + m.index;
}

function splitSolutionSections(tailText) {
  let s = String(tailText || "").trim();
  if (!s) return { solution: "", detail: "" };

  const reCT = /(Giải\s*chi\s*tiết)/i;
  const matchCT = reCT.exec(s);
  if (matchCT) {
    const idxCT = matchCT.index;
    return {
      solution: s.slice(0, idxCT).trim(),
      detail: s.slice(idxCT).trim(),
    };
  }
  return { solution: s, detail: "" };
}

function cleanStemFromQuestionNo(s) {
  return String(s || "").replace(/^Câu\s+\d+\.?\s*/i, "").trim();
}

function splitChoicesTextABCD(blockText) {
  let s = normalizeUnderlinedMarkersForSplit(blockText);
  s = s.replace(/\r/g, "");

  const solIdx = findSolutionMarkerIndex(s, 0);
  const main = solIdx >= 0 ? s.slice(0, solIdx) : s;
  const tail = solIdx >= 0 ? s.slice(solIdx) : "";

  const re = /(^|\n)\s*(\*?)([A-D])\.\s*/g;

  const hits = [];
  let m;
  while ((m = re.exec(main)) !== null) {
    hits.push({ idx: m.index + m[1].length, star: m[2] === "*", key: m[3] });
  }
  if (hits.length < 2) return null;

  const out = {
    stem: main.slice(0, hits[0].idx).trim(),
    choices: { A: "", B: "", C: "", D: "" },
    starredCorrect: null,
    tail,
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : main.length;
    let seg = main.slice(start, end).trim();
    seg = seg.replace(/^(\*?)([A-D])\.\s*/i, "");
    out.choices[key] = seg.trim();
    if (hits[i].star) out.starredCorrect = key;
  }
  return out;
}

function splitStatementsTextabcd(blockText) {
  let s = normalizeUnderlinedMarkersForSplit(blockText);
  s = s.replace(/\r/g, "");

  const solIdx = findSolutionMarkerIndex(s, 0);
  const main = solIdx >= 0 ? s.slice(0, solIdx) : s;
  const tail = solIdx >= 0 ? s.slice(solIdx) : "";

  const re = /(^|\n)\s*([a-d])\)\s*/gi;
  const hits = [];
  let m;
  while ((m = re.exec(main)) !== null) {
    hits.push({ idx: m.index + m[1].length, key: m[2].toLowerCase() });
  }
  if (hits.length < 2) return null;

  const out = {
    stem: main.slice(0, hits[0].idx).trim(),
    statements: { a: "", b: "", c: "", d: "" },
    tail,
  };

  for (let i = 0; i < hits.length; i++) {
    const key = hits[i].key;
    const start = hits[i].idx;
    const end = i + 1 < hits.length ? hits[i + 1].idx : main.length;
    let seg = main.slice(start, end).trim();
    seg = seg.replace(/^([a-d])\)\s*/i, "");
    out.statements[key] = seg.trim();
  }
  return out;
}

function parseExamFromText(text) {
  const blocks = String(text || "").split(/(?=Câu\s+\d+\.)/);
  const exam = { version: 9, questions: [] };

  for (const block of blocks) {
    if (!/^Câu\s+\d+\./i.test(block)) continue;

    const qnoMatch = block.match(/^Câu\s+(\d+)\./i);
    const no = qnoMatch ? Number(qnoMatch[1]) : null;

    const under = extractUnderlinedKeys(block);
    const plain = stripTagsToPlain(block);

    const isMCQ = detectHasMCQ(plain);
    const isTF4 = !isMCQ && detectHasTF4(plain);

    if (isMCQ) {
      const parts = splitChoicesTextABCD(block);
      const tail = parts?.tail || "";
      const solParts = splitSolutionSections(tail);

      const answer = parts?.starredCorrect || under.mcq || null;

      exam.questions.push({
        no,
        type: "mcq",
        stem: cleanStemFromQuestionNo(parts?.stem || block),
        choices: {
          A: parts?.choices?.A || "",
          B: parts?.choices?.B || "",
          C: parts?.choices?.C || "",
          D: parts?.choices?.D || "",
        },
        answer,
        solution: solParts.solution || "",
        detail: solParts.detail || "",
        _plain: plain,
      });
      continue;
    }

    if (isTF4) {
      const parts = splitStatementsTextabcd(block);
      const tail = parts?.tail || "";
      const solParts = splitSolutionSections(tail);

      const ans = { a: null, b: null, c: null, d: null };
      for (const k of ["a", "b", "c", "d"]) {
        if (under.tf.includes(k)) ans[k] = true;
      }

      exam.questions.push({
        no,
        type: "tf4",
        stem: cleanStemFromQuestionNo(parts?.stem || block),
        statements: {
          a: parts?.statements?.a || "",
          b: parts?.statements?.b || "",
          c: parts?.statements?.c || "",
          d: parts?.statements?.d || "",
        },
        answer: ans,
        solution: solParts.solution || "",
        detail: solParts.detail || "",
        _plain: plain,
      });
      continue;
    }

    const solIdx = findSolutionMarkerIndex(block, 0);
    const stemPart = solIdx >= 0 ? block.slice(0, solIdx).trim() : block.trim();
    const tailPart = solIdx >= 0 ? block.slice(solIdx).trim() : "";

    const solParts = splitSolutionSections(tailPart);

    exam.questions.push({
      no,
      type: "short",
      stem: cleanStemFromQuestionNo(stemPart),
      boxes: 4,
      solution: solParts.solution || tailPart || "",
      detail: solParts.detail || "",
      _plain: plain,
    });
  }

  return exam;
}

function legacyQuestionsFromExam(exam) {
  const out = [];
  for (const q of exam.questions) {
    if (q.type !== "mcq") continue;
    out.push({
      type: "multiple_choice",
      content: q.stem,
      choices: [
        { label: "A", text: q.choices.A },
        { label: "B", text: q.choices.B },
        { label: "C", text: q.choices.C },
        { label: "D", text: q.choices.D },
      ],
      correct: q.answer,
      solution: [q.solution, q.detail].filter(Boolean).join("\n").trim(),
    });
  }
  return out;
}

/* ================= helper: gán sectionOrder cho từng question ================= */

function attachSectionOrderToQuestions(exam, sections) {
  if (!exam?.questions?.length || !Array.isArray(sections)) return;

  for (const q of exam.questions) {
    q.sectionOrder = null;
    q.sectionTitle = null;
  }

  for (const sec of sections) {
    if (
      typeof sec.questionIndexStart !== "number" ||
      typeof sec.questionIndexEnd !== "number"
    ) {
      continue;
    }
    const a = Math.max(0, sec.questionIndexStart);
    const b = Math.min(exam.questions.length, sec.questionIndexEnd);
    for (let i = a; i < b; i++) {
      exam.questions[i].sectionOrder = sec.order;
      exam.questions[i].sectionTitle = sec.title;
    }
  }
}

/* ================= ✅ FIX UI: BUILD BLOCKS (SECTION + QUESTION) đúng thứ tự ================= */

function buildOrderedBlocks(exam) {
  const blocks = [];
  let lastSec = null;

  for (const q of exam?.questions || []) {
    const sec = q.sectionOrder || null;
    if (sec && sec !== lastSec) {
      blocks.push({
        type: "section",
        order: sec,
        title: q.sectionTitle || `PHẦN ${sec}`,
      });
      lastSec = sec;
    }
    blocks.push({ type: "question", data: q });
  }
  return blocks;
}

/* ================= API ================= */

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    if (!req.file?.buffer) throw new Error("No file uploaded");

    const zip = await unzipper.Open.buffer(req.file.buffer);

    const docEntry = zip.files.find((f) => f.path === "word/document.xml");
    const relEntry = zip.files.find(
      (f) => f.path === "word/_rels/document.xml.rels"
    );
    if (!docEntry || !relEntry)
      throw new Error("Missing document.xml or document.xml.rels");

    let docXml = (await docEntry.buffer()).toString("utf8");
    const relsXml = (await relEntry.buffer()).toString("utf8");
    const rels = parseRels(relsXml);

    // 1) MathType -> LaTeX (and fallback images)
    const images = {};
    const mt = await tokenizeMathTypeOleFirst(docXml, rels, zip.files, images);
    docXml = mt.outXml;
    const latexMap = mt.latexMap;

    // 2) normal images
    const imgTok = await tokenizeImagesAfter(docXml, rels, zip.files);
    docXml = imgTok.outXml;
    Object.assign(images, imgTok.imgMap);

    // 3) text (giữ token + underline + ✅ TABLE)
    const text = wordXmlToTextKeepTokens(docXml);

    // 4) parse exam output (GIỮ NGUYÊN)
    const exam = parseExamFromText(text);

    // sections theo vị trí + index câu toàn cục
    const sections = extractSectionTitles(text);

    exam.sections = sections;

    attachSectionOrderToQuestions(exam, sections);

    const blocks = buildOrderedBlocks(exam);

    const questions = legacyQuestionsFromExam(exam);

    res.json({
      ok: true,
      total: exam.questions.length,
      sections,
      blocks,
      exam,
      questions,
      latex: latexMap,
      images,
      rawText: text,
      debug: {
        latexCount: Object.keys(latexMap).length,
        imagesCount: Object.keys(images).length,
        exam: {
          questions: exam.questions.length,
          mcq: exam.questions.filter((x) => x.type === "mcq").length,
          tf4: exam.questions.filter((x) => x.type === "tf4").length,
          short: exam.questions.filter((x) => x.type === "short").length,
        },
      },
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: err?.message || String(err) });
  }
});

app.get("/ping", (_, res) => res.send("ok"));

app.get("/debug-inkscape", (_, res) => {
  try {
    const v = execFileSync("inkscape", ["--version"]).toString();
    res.type("text/plain").send(v);
  } catch {
    res.status(500).type("text/plain").send("NO INKSCAPE");
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server running on", PORT));
