/* ============================================================================
 * 통PDF (TongPDF) - 100% Client-Side Document & PDF Utility
 * Supports: .hwpx, .hwp, .pptx, .docx, .txt/.md, .show (best-effort + PC guide),
 *           Images <-> PDF, PDF Merge / Split
 * ============================================================================ */

(function () {
  'use strict';

  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  }

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  // -------------------- Tab Navigation --------------------
  function switchTab(tabId) {
    $$('.tab-btn').forEach((btn) => {
      const active = btn.dataset.tab === tabId;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-selected', String(active));
    });
    $$('.panel').forEach((panel) => {
      panel.classList.toggle('active', panel.id === 'panel-' + tabId);
    });
  }

  $$('.tab-btn').forEach((btn) => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // Support deep-linking tab via ?tab=doc2pdf|imgPdf|mergeSplit|pcTool
  const initialTab = new URLSearchParams(window.location.search).get('tab');
  if (initialTab && ['doc2pdf', 'imgPdf', 'mergeSplit', 'pcTool'].includes(initialTab)) {
    switchTab(initialTab);
  }

  $$('[data-switch-tab]').forEach((btn) => {
    btn.addEventListener('click', () => {
      switchTab(btn.dataset.switchTab);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });

  // -------------------- Common Helpers --------------------
  function formatBytes(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
  }

  function getExt(filename) {
    const idx = filename.lastIndexOf('.');
    return idx >= 0 ? filename.slice(idx + 1).toLowerCase() : '';
  }

  function baseName(filename) {
    const idx = filename.lastIndexOf('.');
    return idx >= 0 ? filename.slice(0, idx) : filename;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 15000);
  }

  function wireDropzone(dropzoneEl, inputEl, onFiles) {
    dropzoneEl.addEventListener('click', () => inputEl.click());
    inputEl.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length) {
        onFiles(Array.from(e.target.files));
        inputEl.value = '';
      }
    });
    ['dragenter', 'dragover'].forEach((ev) => {
      dropzoneEl.addEventListener(ev, (e) => {
        e.preventDefault();
        dropzoneEl.classList.add('dragover');
      });
    });
    ['dragleave', 'drop'].forEach((ev) => {
      dropzoneEl.addEventListener(ev, (e) => {
        e.preventDefault();
        dropzoneEl.classList.remove('dragover');
      });
    });
    dropzoneEl.addEventListener('drop', (e) => {
      const files = Array.from(e.dataTransfer.files || []);
      if (files.length) onFiles(files);
    });
  }

  function loadImageElement(blob) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => {
        URL.revokeObjectURL(url);
        resolve(img);
      };
      img.onerror = (err) => {
        URL.revokeObjectURL(url);
        reject(err);
      };
      img.src = url;
    });
  }

  // -------------------- Pure-JS OLE2/CFB Parser for legacy .hwp --------------------
  async function inflateRaw(u8) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('이 브라우저는 압축 해제를 지원하지 않습니다.');
    }
    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    writer.write(u8);
    writer.close();
    const out = await new Response(ds.readable).arrayBuffer();
    return new Uint8Array(out);
  }

  async function extractHwpText(arrayBuffer) {
    const dv = new DataView(arrayBuffer);
    const u8 = new Uint8Array(arrayBuffer);
    // Verify OLE2 header: D0 CF 11 E0 A1 B1 1A E1
    const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    for (let i = 0; i < 8; i++) {
      if (u8[i] !== sig[i]) {
        // Fallback: maybe it's actually an HWPX renamed to .hwp or plain text
        if (u8[0] === 0x50 && u8[1] === 0x4b) {
          return { isZip: true, blocks: [] };
        }
        throw new Error('유효한 한글(.hwp) 파일 형식이 아닙니다.');
      }
    }

    const sectorSize = 1 << dv.getUint16(30, true);
    const miniSectorSize = 1 << dv.getUint16(32, true);
    const fatCount = dv.getUint32(44, true);
    const dirStart = dv.getUint32(48, true);
    const miniCutoff = dv.getUint32(56, true);
    const miniFatStart = dv.getUint32(60, true);
    const difatStart = dv.getUint32(68, true);
    const difatCount = dv.getUint32(72, true);

    const sectorOffset = (sec) => (sec + 1) * sectorSize;

    // Collect FAT sector numbers from header DIFAT (109 entries) + extra DIFAT chain
    const fatSectors = [];
    for (let i = 0; i < 109 && fatSectors.length < fatCount; i++) {
      const s = dv.getUint32(76 + i * 4, true);
      if (s < 0xfffffffe) fatSectors.push(s);
    }
    let nextDifat = difatStart;
    for (let c = 0; c < difatCount && nextDifat < 0xfffffffe; c++) {
      const off = sectorOffset(nextDifat);
      const count = sectorSize / 4 - 1;
      for (let i = 0; i < count && fatSectors.length < fatCount; i++) {
        const s = dv.getUint32(off + i * 4, true);
        if (s < 0xfffffffe) fatSectors.push(s);
      }
      nextDifat = dv.getUint32(off + count * 4, true);
    }

    // Build FAT table
    const fat = [];
    for (const sec of fatSectors) {
      const off = sectorOffset(sec);
      for (let i = 0; i < sectorSize / 4; i++) {
        fat.push(dv.getUint32(off + i * 4, true));
      }
    }

    function readChain(startSec, size) {
      const chunks = [];
      let sec = startSec;
      const visited = new Set();
      while (sec < 0xfffffffe && !visited.has(sec)) {
        visited.add(sec);
        const off = sectorOffset(sec);
        if (off + sectorSize > u8.length) break;
        chunks.push(u8.subarray(off, off + sectorSize));
        sec = fat[sec];
      }
      const total = chunks.reduce((a, b) => a + b.length, 0);
      const res = new Uint8Array(size != null ? Math.min(size, total) : total);
      let pos = 0;
      for (const ch of chunks) {
        const len = Math.min(ch.length, res.length - pos);
        if (len <= 0) break;
        res.set(ch.subarray(0, len), pos);
        pos += len;
      }
      return res;
    }

    // Read Directory Stream
    const dirBytes = readChain(dirStart, null);
    const dirDv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
    const entries = [];
    for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
      const nameLen = dirDv.getUint16(off + 64, true);
      if (nameLen < 2) continue;
      let name = '';
      for (let j = 0; j < nameLen - 2; j += 2) {
        name += String.fromCharCode(dirDv.getUint16(off + j, true));
      }
      const type = dirBytes[off + 66];
      const startSec = dirDv.getUint32(off + 116, true);
      const size = dirDv.getUint32(off + 120, true);
      entries.push({ name, type, startSec, size });
    }

    // Mini stream & Mini FAT
    const rootEntry = entries[0];
    const miniStream = rootEntry ? readChain(rootEntry.startSec, rootEntry.size) : new Uint8Array(0);
    const miniFatBytes = miniFatStart < 0xfffffffe ? readChain(miniFatStart, null) : new Uint8Array(0);
    const miniFat = [];
    const mfDv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
    for (let i = 0; i + 4 <= miniFatBytes.length; i += 4) {
      miniFat.push(mfDv.getUint32(i, true));
    }

    function readEntryData(entry) {
      if (entry.size < miniCutoff && entry.type === 2) {
        const res = new Uint8Array(entry.size);
        let sec = entry.startSec;
        let pos = 0;
        const visited = new Set();
        while (sec < 0xfffffffe && pos < entry.size && !visited.has(sec)) {
          visited.add(sec);
          const off = sec * miniSectorSize;
          const len = Math.min(miniSectorSize, entry.size - pos, miniStream.length - off);
          if (len <= 0) break;
          res.set(miniStream.subarray(off, off + len), pos);
          pos += len;
          sec = miniFat[sec];
        }
        return res;
      }
      return readChain(entry.startSec, entry.size);
    }

    // Check FileHeader flags (bit 0: compressed, bit 1: encrypted)
    const headerEntry = entries.find((e) => e.name === 'FileHeader');
    let compressed = true;
    let encrypted = false;
    if (headerEntry) {
      const hData = readEntryData(headerEntry);
      if (hData.length >= 40) {
        const flags = new DataView(hData.buffer, hData.byteOffset, hData.byteLength).getUint32(36, true);
        compressed = (flags & 1) !== 0;
        encrypted = (flags & 2) !== 0;
      }
    }
    if (encrypted) {
      throw new Error('암호가 걸린 .hwp 파일은 브라우저에서 읽을 수 없습니다. PC용 변환기(.exe)를 이용해 주세요.');
    }

    const sectionEntries = entries
      .filter((e) => /^Section\d+$/i.test(e.name))
      .sort((a, b) => parseInt(a.name.replace(/\D/g, ''), 10) - parseInt(b.name.replace(/\D/g, ''), 10));

    const blocks = [];
    for (const sec of sectionEntries) {
      let data = readEntryData(sec);
      if (compressed) {
        try {
          data = await inflateRaw(data);
        } catch (_) {
          // If already uncompressed, continue
        }
      }
      const sDv = new DataView(data.buffer, data.byteOffset, data.byteLength);
      let pos = 0;
      while (pos + 4 <= data.length) {
        const hdr = sDv.getUint32(pos, true);
        pos += 4;
        const tagId = hdr & 0x3ff;
        let size = (hdr >>> 20) & 0xfff;
        if (size === 0xfff) {
          if (pos + 4 > data.length) break;
          size = sDv.getUint32(pos, true);
          pos += 4;
        }
        if (pos + size > data.length) break;
        // HWPTAG_PARA_TEXT = 67 (0x43)
        if (tagId === 67) {
          let text = '';
          let i = 0;
          while (i + 2 <= size) {
            const code = sDv.getUint16(pos + i, true);
            i += 2;
            if (code < 32) {
              // Extended/inline controls consume extra 12 bytes (6 UTF-16 units)
              if ([1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23].includes(code)) {
                i += 12;
              } else if (code === 4 || code === 5 || code === 6 || code === 7 || code === 8 || code === 9) {
                i += 12;
                if (code === 9) text += '    ';
              } else if (code === 10 || code === 13) {
                text += '\n';
              }
            } else {
              text += String.fromCharCode(code);
            }
          }
          const cleaned = text.replace(/\u0000/g, '').trim();
          if (cleaned) {
            blocks.push({ type: 'p', text: cleaned });
          }
        }
        pos += size;
      }
    }

    // Fallback to PrvText if BodyText produced nothing
    if (!blocks.length) {
      const prv = entries.find((e) => e.name === 'PrvText');
      if (prv) {
        const pData = readEntryData(prv);
        const text = new TextDecoder('utf-16le').decode(pData).replace(/\u0000/g, '').trim();
        text.split(/\r?\n/).forEach((line) => {
          if (line.trim()) blocks.push({ type: 'p', text: line.trim() });
        });
      }
    }

    return { isZip: false, blocks };
  }

  // -------------------- HWPX Parser --------------------
  async function parseHwpx(arrayBuffer) {
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();

    // Load BinData images map (id -> blob)
    const binImages = new Map();
    const binFiles = Object.keys(zip.files).filter((k) => /^BinData\//i.test(k) && !zip.files[k].dir);
    for (const bf of binFiles) {
      const ext = getExt(bf);
      if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp'].includes(ext)) {
        const key = baseName(bf.split('/').pop());
        const blob = await zip.files[bf].async('blob');
        binImages.set(key, blob);
      }
    }

    const sectionFiles = Object.keys(zip.files)
      .filter((k) => /^Contents\/section\d+\.xml$/i.test(k))
      .sort((a, b) => {
        const na = parseInt(a.replace(/\D/g, ''), 10) || 0;
        const nb = parseInt(b.replace(/\D/g, ''), 10) || 0;
        return na - nb;
      });

    const blocks = [];

    function extractTextFromNode(node) {
      const parts = [];
      const walk = (n) => {
        if (!n) return;
        const local = n.localName || n.nodeName.split(':').pop();
        if (local === 't') {
          parts.push(n.textContent || '');
          return;
        }
        if (local === 'tab') {
          parts.push('    ');
          return;
        }
        if (local === 'lineBreak') {
          parts.push('\n');
          return;
        }
        if (local === 'tbl') return; // handled separately
        for (const ch of n.childNodes || []) walk(ch);
      };
      walk(node);
      return parts.join('').replace(/\s+\n/g, '\n').trim();
    }

    for (const sf of sectionFiles) {
      const xmlStr = await zip.files[sf].async('string');
      const doc = parser.parseFromString(xmlStr, 'application/xml');
      const root = doc.documentElement;

      for (const child of Array.from(root.children)) {
        const local = child.localName || child.nodeName.split(':').pop();
        if (local !== 'p') continue;

        // Check for embedded tables inside this paragraph
        const tables = Array.from(child.getElementsByTagName('*')).filter(
          (el) => (el.localName || el.nodeName.split(':').pop()) === 'tbl'
        );

        // Check for embedded pictures inside this paragraph
        const imgs = Array.from(child.getElementsByTagName('*')).filter(
          (el) => (el.localName || el.nodeName.split(':').pop()) === 'img'
        );

        const paraText = extractTextFromNode(child);
        if (paraText) {
          blocks.push({ type: 'p', text: paraText });
        }

        for (const imgEl of imgs) {
          const ref = imgEl.getAttribute('binaryItemIDRef') || imgEl.getAttribute('hc:binaryItemIDRef');
          if (ref && binImages.has(ref)) {
            blocks.push({ type: 'image', blob: binImages.get(ref) });
          }
        }

        for (const tbl of tables) {
          const rows = [];
          const trEls = Array.from(tbl.children).filter(
            (el) => (el.localName || el.nodeName.split(':').pop()) === 'tr'
          );
          for (const tr of trEls) {
            const cells = [];
            const tcEls = Array.from(tr.children).filter(
              (el) => (el.localName || el.nodeName.split(':').pop()) === 'tc'
            );
            for (const tc of tcEls) {
              const cellTexts = [];
              const tEls = Array.from(tc.getElementsByTagName('*')).filter(
                (el) => (el.localName || el.nodeName.split(':').pop()) === 't'
              );
              tEls.forEach((t) => {
                if (t.textContent) cellTexts.push(t.textContent);
              });
              cells.push(cellTexts.join(' ').trim());
            }
            if (cells.length) rows.push(cells);
          }
          if (rows.length) {
            blocks.push({ type: 'table', rows });
          }
        }
      }
    }

    return blocks;
  }

  // -------------------- DOCX Parser --------------------
  async function parseDocx(arrayBuffer) {
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const docFile = zip.files['word/document.xml'];
    if (!docFile) throw new Error('word/document.xml을 찾을 수 없습니다.');
    const xmlStr = await docFile.async('string');
    const doc = new DOMParser().parseFromString(xmlStr, 'application/xml');
    const body = doc.getElementsByTagName('*');
    let bodyEl = null;
    for (const el of body) {
      if ((el.localName || el.nodeName.split(':').pop()) === 'body') {
        bodyEl = el;
        break;
      }
    }
    if (!bodyEl) return [];

    const blocks = [];
    for (const ch of Array.from(bodyEl.children)) {
      const local = ch.localName || ch.nodeName.split(':').pop();
      if (local === 'p') {
        const texts = Array.from(ch.getElementsByTagName('*'))
          .filter((el) => (el.localName || el.nodeName.split(':').pop()) === 't')
          .map((el) => el.textContent || '');
        const line = texts.join('').trim();
        if (line) {
          const pStyle = Array.from(ch.getElementsByTagName('*')).find(
            (el) => (el.localName || el.nodeName.split(':').pop()) === 'pStyle'
          );
          const isHeading = pStyle && /heading|제목/i.test(pStyle.getAttribute('w:val') || '');
          blocks.push({ type: isHeading ? 'heading' : 'p', text: line });
        }
      } else if (local === 'tbl') {
        const rows = [];
        const trs = Array.from(ch.children).filter((e) => (e.localName || e.nodeName.split(':').pop()) === 'tr');
        for (const tr of trs) {
          const cells = [];
          const tcs = Array.from(tr.children).filter((e) => (e.localName || e.nodeName.split(':').pop()) === 'tc');
          for (const tc of tcs) {
            const ts = Array.from(tc.getElementsByTagName('*'))
              .filter((e) => (e.localName || e.nodeName.split(':').pop()) === 't')
              .map((e) => e.textContent || '');
            cells.push(ts.join(' ').trim());
          }
          if (cells.length) rows.push(cells);
        }
        if (rows.length) blocks.push({ type: 'table', rows });
      }
    }
    return blocks;
  }

  // -------------------- PPTX Parser (Slides) --------------------
  async function parsePptx(arrayBuffer) {
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();

    const slideFiles = Object.keys(zip.files)
      .filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k))
      .sort((a, b) => {
        const na = parseInt(a.replace(/\D/g, ''), 10) || 0;
        const nb = parseInt(b.replace(/\D/g, ''), 10) || 0;
        return na - nb;
      });

    const slides = [];
    for (let i = 0; i < slideFiles.length; i++) {
      const sf = slideFiles[i];
      const xmlStr = await zip.files[sf].async('string');
      const doc = parser.parseFromString(xmlStr, 'application/xml');

      // Extract paragraphs grouped by shape
      const paragraphs = [];
      const pEls = Array.from(doc.getElementsByTagName('*')).filter(
        (el) => (el.localName || el.nodeName.split(':').pop()) === 'p'
      );
      for (const pEl of pEls) {
        const tEls = Array.from(pEl.getElementsByTagName('*')).filter(
          (el) => (el.localName || el.nodeName.split(':').pop()) === 't'
        );
        const line = tEls.map((t) => t.textContent || '').join('').trim();
        if (line) paragraphs.push(line);
      }

      // Extract slide images via _rels
      const images = [];
      const relPath = sf.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels';
      if (zip.files[relPath]) {
        const relXml = await zip.files[relPath].async('string');
        const relDoc = parser.parseFromString(relXml, 'application/xml');
        const rels = Array.from(relDoc.getElementsByTagName('Relationship'));
        for (const r of rels) {
          const target = r.getAttribute('Target') || '';
          if (/\.(png|jpe?g|webp|gif)$/i.test(target)) {
            const norm = 'ppt/' + target.replace(/^\.\.\//, '').replace(/^\/+/, '');
            if (zip.files[norm]) {
              images.push(await zip.files[norm].async('blob'));
            }
          }
        }
      }

      slides.push({
        number: i + 1,
        title: paragraphs[0] || `슬라이드 ${i + 1}`,
        lines: paragraphs.slice(1),
        images
      });
    }
    return slides;
  }

  // -------------------- Best-effort .show (Hancom Show) Parser --------------------
  async function parseShowBestEffort(arrayBuffer) {
    const u8 = new Uint8Array(arrayBuffer);
    if (u8[0] !== 0x50 || u8[1] !== 0x4b) {
      throw new Error(
        '구형 바이너리 한쇼(.show) 파일은 브라우저에서 읽을 수 없습니다. 상단의 [PC용 무설치 변환기(.exe)]를 이용해 주세요.'
      );
    }
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();
    const xmlFiles = Object.keys(zip.files)
      .filter((k) => /\.xml$/i.test(k) && /slide|section|content/i.test(k))
      .sort();

    const slides = [];
    let idx = 1;
    for (const xf of xmlFiles) {
      const xmlStr = await zip.files[xf].async('string');
      const doc = parser.parseFromString(xmlStr, 'application/xml');
      const texts = Array.from(doc.getElementsByTagName('*'))
        .filter((el) => {
          const ln = el.localName || el.nodeName.split(':').pop();
          return ln === 't' || ln === 'text';
        })
        .map((el) => (el.textContent || '').trim())
        .filter(Boolean);
      if (texts.length) {
        slides.push({
          number: idx++,
          title: texts[0],
          lines: texts.slice(1),
          images: []
        });
      }
    }
    if (!slides.length) {
      throw new Error(
        '한쇼(.show) 내부 슬라이드를 브라우저에서 해석할 수 없습니다. 상단의 [PC용 무설치 변환기(.exe)]로 원본 100% 변환해 주세요.'
      );
    }
    return slides;
  }

  // -------------------- Canvas Pagination -> PDF Builder --------------------
  function getPageSpec(orientation, isSlide) {
    // High-DPI canvas pixels & PDF points (72 pt/in)
    if (orientation === 'slide169' || (orientation === 'auto' && isSlide)) {
      return { cw: 1600, ch: 900, pw: 960, ph: 540, isSlide: true };
    }
    if (orientation === 'landscape') {
      return { cw: 1684, ch: 1190, pw: 841.89, ph: 595.28, isSlide: false };
    }
    // A4 portrait
    return { cw: 1240, ch: 1754, pw: 595.28, ph: 841.89, isSlide: false };
  }

  function wrapLines(ctx, text, maxWidth) {
    const out = [];
    const paragraphs = String(text || '').split(/\r?\n/);
    for (const para of paragraphs) {
      if (!para) {
        out.push('');
        continue;
      }
      let current = '';
      for (const ch of para) {
        const next = current + ch;
        if (ctx.measureText(next).width > maxWidth && current) {
          out.push(current);
          current = ch;
        } else {
          current = next;
        }
      }
      if (current) out.push(current);
    }
    return out;
  }

  async function renderDocumentBlocksToPdf(title, blocks, orientation, fontSizeOpt) {
    const spec = getPageSpec(orientation, false);
    const fontPx = fontSizeOpt === 'small' ? 22 : fontSizeOpt === 'large' ? 30 : 25;
    const lineH = Math.round(fontPx * 1.55);
    const marginX = 88;
    const marginTop = 88;
    const marginBottom = 86;
    const contentW = spec.cw - marginX * 2;
    const fontFamily = '"Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';

    const canvases = [];
    let canvas, ctx, y;

    function newPage(isFirst) {
      canvas = document.createElement('canvas');
      canvas.width = spec.cw;
      canvas.height = spec.ch;
      ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, spec.cw, spec.ch);
      canvases.push(canvas);
      y = marginTop;

      if (isFirst && title) {
        ctx.fillStyle = '#0f172a';
        ctx.font = `800 ${fontPx + 9}px ${fontFamily}`;
        const tLines = wrapLines(ctx, title, contentW);
        for (const tl of tLines) {
          ctx.fillText(tl, marginX, y + fontPx);
          y += Math.round((fontPx + 9) * 1.35);
        }
        ctx.strokeStyle = '#cbd5e1';
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.moveTo(marginX, y + 6);
        ctx.lineTo(spec.cw - marginX, y + 6);
        ctx.stroke();
        y += 28;
      }
    }

    newPage(true);

    for (const blk of blocks) {
      if (blk.type === 'heading' || blk.type === 'p') {
        const isHead = blk.type === 'heading';
        const px = isHead ? fontPx + 5 : fontPx;
        const lh = Math.round(px * 1.55);
        ctx.font = `${isHead ? '700' : '400'} ${px}px ${fontFamily}`;
        ctx.fillStyle = isHead ? '#0f172a' : '#1e293b';
        const lines = wrapLines(ctx, blk.text, contentW);
        if (isHead) y += 10;
        for (const line of lines) {
          if (y + lh > spec.ch - marginBottom) {
            newPage(false);
            ctx.font = `${isHead ? '700' : '400'} ${px}px ${fontFamily}`;
            ctx.fillStyle = isHead ? '#0f172a' : '#1e293b';
          }
          ctx.fillText(line, marginX, y + px);
          y += lh;
        }
        y += Math.round(px * 0.45);
      } else if (blk.type === 'table' && blk.rows && blk.rows.length) {
        const cellPx = Math.max(18, fontPx - 3);
        const cellLh = Math.round(cellPx * 1.45);
        const pad = 10;
        for (let rIdx = 0; rIdx < blk.rows.length; rIdx++) {
          const row = blk.rows[rIdx];
          const cols = Math.max(1, row.length);
          const colW = contentW / cols;
          ctx.font = `${rIdx === 0 ? '700' : '400'} ${cellPx}px ${fontFamily}`;
          const wrappedCells = row.map((cell) => wrapLines(ctx, cell || '', colW - pad * 2));
          const maxLines = Math.max(1, ...wrappedCells.map((l) => l.length));
          const rowH = maxLines * cellLh + pad * 2;

          if (y + rowH > spec.ch - marginBottom) {
            newPage(false);
          }

          for (let cIdx = 0; cIdx < cols; cIdx++) {
            const cx = marginX + cIdx * colW;
            ctx.fillStyle = rIdx === 0 ? '#f1f5f9' : '#ffffff';
            ctx.fillRect(cx, y, colW, rowH);
            ctx.strokeStyle = '#94a3b8';
            ctx.lineWidth = 1.5;
            ctx.strokeRect(cx, y, colW, rowH);

            ctx.fillStyle = '#0f172a';
            ctx.font = `${rIdx === 0 ? '700' : '400'} ${cellPx}px ${fontFamily}`;
            const lines = wrappedCells[cIdx];
            lines.forEach((ln, li) => {
              ctx.fillText(ln, cx + pad, y + pad + cellPx + li * cellLh - 3);
            });
          }
          y += rowH;
        }
        y += 18;
      } else if (blk.type === 'image' && blk.blob) {
        try {
          const img = await loadImageElement(blk.blob);
          const maxW = contentW;
          const maxH = spec.ch - marginTop - marginBottom - 40;
          let drawW = img.width;
          let drawH = img.height;
          if (drawW > maxW) {
            const ratio = maxW / drawW;
            drawW *= ratio;
            drawH *= ratio;
          }
          if (drawH > maxH) {
            const ratio = maxH / drawH;
            drawW *= ratio;
            drawH *= ratio;
          }
          if (y + drawH > spec.ch - marginBottom) {
            newPage(false);
          }
          const dx = marginX + (contentW - drawW) / 2;
          ctx.drawImage(img, dx, y, drawW, drawH);
          y += drawH + 20;
        } catch (_) {
          // Ignore unsupported image format
        }
      }
    }

    // Add page footers
    const totalPages = canvases.length;
    canvases.forEach((cv, idx) => {
      const c = cv.getContext('2d');
      c.fillStyle = '#64748b';
      c.font = `500 18px ${fontFamily}`;
      const footerText = `${idx + 1} / ${totalPages}`;
      const tw = c.measureText(footerText).width;
      c.fillText(footerText, (spec.cw - tw) / 2, spec.ch - 34);
    });

    return canvasesToPdfBytes(canvases, spec.pw, spec.ph);
  }

  async function renderSlidesToPdf(slides, orientation, fontSizeOpt) {
    const spec = getPageSpec(orientation, true);
    const fontFamily = '"Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';
    const basePx = fontSizeOpt === 'small' ? 24 : fontSizeOpt === 'large' ? 34 : 28;
    const canvases = [];

    for (const slide of slides) {
      const canvas = document.createElement('canvas');
      canvas.width = spec.cw;
      canvas.height = spec.ch;
      const ctx = canvas.getContext('2d');

      // Clean presentation card background
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, spec.cw, spec.ch);

      // Top accent bar
      ctx.fillStyle = '#1d4ed8';
      ctx.fillRect(0, 0, spec.cw, 14);

      const marginX = 76;
      const contentW = spec.cw - marginX * 2;
      let y = 76;

      // Slide title
      ctx.fillStyle = '#0f172a';
      ctx.font = `800 ${basePx + 14}px ${fontFamily}`;
      const titleLines = wrapLines(ctx, slide.title, contentW);
      for (const tl of titleLines.slice(0, 2)) {
        ctx.fillText(tl, marginX, y + basePx + 8);
        y += Math.round((basePx + 14) * 1.3);
      }

      ctx.strokeStyle = '#e2e8f0';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(marginX, y + 6);
      ctx.lineTo(spec.cw - marginX, y + 6);
      ctx.stroke();
      y += 32;

      const hasImages = slide.images && slide.images.length > 0;
      const textColW = hasImages ? Math.round(contentW * 0.56) : contentW;

      // Slide body text
      ctx.fillStyle = '#1e293b';
      ctx.font = `400 ${basePx}px ${fontFamily}`;
      const lh = Math.round(basePx * 1.48);
      for (const line of slide.lines) {
        const wrapped = wrapLines(ctx, '• ' + line, textColW);
        for (const wl of wrapped) {
          if (y + lh > spec.ch - 70) break;
          ctx.fillText(wl, marginX, y + basePx);
          y += lh;
        }
        y += 8;
        if (y + lh > spec.ch - 70) break;
      }

      // First slide image on the right column if available
      if (hasImages) {
        try {
          const img = await loadImageElement(slide.images[0]);
          const boxX = marginX + textColW + 36;
          const boxY = 180;
          const boxW = contentW - textColW - 36;
          const boxH = spec.ch - boxY - 80;
          const ratio = Math.min(boxW / img.width, boxH / img.height, 1.5);
          const dw = img.width * ratio;
          const dh = img.height * ratio;
          ctx.drawImage(img, boxX + (boxW - dw) / 2, boxY + (boxH - dh) / 2, dw, dh);
        } catch (_) {}
      }

      // Slide number footer
      ctx.fillStyle = '#64748b';
      ctx.font = `600 20px ${fontFamily}`;
      ctx.fillText(`${slide.number} / ${slides.length}`, spec.cw - marginX - 70, spec.ch - 32);

      canvases.push(canvas);
    }

    return canvasesToPdfBytes(canvases, spec.pw, spec.ph);
  }

  async function canvasesToPdfBytes(canvases, pageW, pageH) {
    const pdfDoc = await window.PDFLib.PDFDocument.create();
    for (const canvas of canvases) {
      const dataUrl = canvas.toDataURL('image/jpeg', 0.9);
      const jpgBytes = await fetch(dataUrl).then((r) => r.arrayBuffer());
      const jpgImg = await pdfDoc.embedJpg(jpgBytes);
      const page = pdfDoc.addPage([pageW, pageH]);
      page.drawImage(jpgImg, { x: 0, y: 0, width: pageW, height: pageH });
    }
    return pdfDoc.save();
  }

  // -------------------- TAB 1: Doc -> PDF --------------------
  const docQueue = [];
  const docFileList = $('#docFileList');
  const docStatus = $('#docStatus');
  const convertAllDocsBtn = $('#convertAllDocsBtn');
  const downloadAllZipBtn = $('#downloadAllZipBtn');

  function renderDocQueue() {
    docFileList.innerHTML = '';
    const hasPending = docQueue.length > 0;
    const hasDone = docQueue.some((item) => item.pdfBlob);
    convertAllDocsBtn.disabled = !hasPending;
    downloadAllZipBtn.disabled = !hasDone;

    docQueue.forEach((item, idx) => {
      const div = document.createElement('div');
      div.className = 'file-item';
      const extClass = ['pptx', 'ppt'].includes(item.ext)
        ? 'pptx'
        : item.ext === 'show'
        ? 'show'
        : '';

      div.innerHTML = `
        <div class="file-meta">
          <span class="file-ext ${extClass}">.${item.ext || '?'}</span>
          <div>
            <div class="file-name">${escapeHtml(item.file.name)}</div>
            <div class="file-sub">${formatBytes(item.file.size)}${item.note ? ' · ' + escapeHtml(item.note) : ''}</div>
          </div>
        </div>
        <span class="file-status ${item.statusClass}">${escapeHtml(item.statusText)}</span>
        <div class="file-actions">
          ${
            item.pdfBlob
              ? `<button type="button" class="btn btn-outline btn-sm" data-preview="${idx}">미리보기</button>
                 <button type="button" class="btn btn-primary btn-sm" data-download="${idx}">PDF 저장</button>`
              : `<button type="button" class="btn btn-outline btn-sm" data-convert="${idx}">변환</button>`
          }
          <button type="button" class="btn btn-danger-outline btn-sm" data-remove="${idx}">삭제</button>
        </div>
      `;
      docFileList.appendChild(div);
    });

    docFileList.querySelectorAll('[data-convert]').forEach((btn) => {
      btn.addEventListener('click', () => convertOneDoc(Number(btn.dataset.convert)));
    });
    docFileList.querySelectorAll('[data-download]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const it = docQueue[Number(btn.dataset.download)];
        if (it && it.pdfBlob) downloadBlob(it.pdfBlob, baseName(it.file.name) + '.pdf');
      });
    });
    docFileList.querySelectorAll('[data-preview]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const it = docQueue[Number(btn.dataset.preview)];
        if (it && it.pdfBlob) {
          const url = URL.createObjectURL(it.pdfBlob);
          window.open(url, '_blank', 'noopener,noreferrer');
        }
      });
    });
    docFileList.querySelectorAll('[data-remove]').forEach((btn) => {
      btn.addEventListener('click', () => {
        docQueue.splice(Number(btn.dataset.remove), 1);
        renderDocQueue();
      });
    });
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  wireDropzone($('#docDropzone'), $('#docFileInput'), (files) => {
    for (const file of files) {
      const ext = getExt(file.name);
      const item = {
        file,
        ext,
        pdfBlob: null,
        statusText: '대기 중',
        statusClass: '',
        note: ''
      };
      if (ext === 'ppt') {
        item.statusText = 'PC용 .exe 권장';
        item.statusClass = 'warn';
        item.note = '구형 .ppt는 .pptx로 저장하거나 PC용 .exe 사용';
      } else if (ext === 'show') {
        item.statusText = '제한 지원 (PC용 .exe 권장)';
        item.statusClass = 'warn';
        item.note = '원본 슬라이드 100% 변환은 PC용 .exe 이용';
      }
      docQueue.push(item);
    }
    docStatus.textContent = `${files.length}개 파일이 추가되었습니다. [전체 PDF 변환]을 눌러주세요.`;
    renderDocQueue();
  });

  async function convertOneDoc(index) {
    const item = docQueue[index];
    if (!item) return;

    const orientation = $('#docOrientation').value;
    const fontSizeOpt = $('#docFontSize').value;

    item.statusText = '변환 중...';
    item.statusClass = 'working';
    renderDocQueue();

    try {
      const buf = await item.file.arrayBuffer();
      const title = baseName(item.file.name);
      let pdfBytes;

      if (item.ext === 'hwpx') {
        const blocks = await parseHwpx(buf);
        if (!blocks.length) throw new Error('문서에서 읽을 본문이 없습니다.');
        pdfBytes = await renderDocumentBlocksToPdf(title, blocks, orientation, fontSizeOpt);
      } else if (item.ext === 'hwp') {
        const parsed = await extractHwpText(buf);
        const blocks = parsed.isZip ? await parseHwpx(buf) : parsed.blocks;
        if (!blocks.length) throw new Error('한글 본문을 추출할 수 없습니다. PC용 변환기(.exe)를 이용해 주세요.');
        pdfBytes = await renderDocumentBlocksToPdf(title, blocks, orientation, fontSizeOpt);
      } else if (item.ext === 'docx') {
        const blocks = await parseDocx(buf);
        if (!blocks.length) throw new Error('워드 문서에서 본문을 찾을 수 없습니다.');
        pdfBytes = await renderDocumentBlocksToPdf(title, blocks, orientation, fontSizeOpt);
      } else if (item.ext === 'pptx') {
        const slides = await parsePptx(buf);
        if (!slides.length) throw new Error('슬라이드를 찾을 수 없습니다.');
        pdfBytes = await renderSlidesToPdf(slides, orientation, fontSizeOpt);
      } else if (item.ext === 'show') {
        const slides = await parseShowBestEffort(buf);
        pdfBytes = await renderSlidesToPdf(slides, orientation, fontSizeOpt);
        item.note = '텍스트 요약 PDF (원본 모양은 PC용 .exe 권장)';
      } else if (item.ext === 'ppt') {
        throw new Error('구형 바이너리 .ppt는 브라우저 변환이 불가합니다. 파워포인트에서 .pptx로 저장하거나 PC용 .exe를 사용해 주세요.');
      } else if (item.ext === 'txt' || item.ext === 'md') {
        const text = new TextDecoder('utf-8').decode(buf);
        const blocks = text.split(/\r?\n/).map((l) => ({ type: l.startsWith('#') ? 'heading' : 'p', text: l.replace(/^#+\s*/, '') }));
        pdfBytes = await renderDocumentBlocksToPdf(title, blocks, orientation, fontSizeOpt);
      } else {
        throw new Error(`지원하지 않는 확장자(.${item.ext})입니다.`);
      }

      item.pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
      item.statusText = `완료 (${formatBytes(item.pdfBlob.size)})`;
      item.statusClass = 'done';
    } catch (err) {
      item.statusText = '변환 실패';
      item.statusClass = 'error';
      item.note = err.message || '오류가 발생했습니다.';
    }
    renderDocQueue();
  }

  convertAllDocsBtn.addEventListener('click', async () => {
    convertAllDocsBtn.disabled = true;
    docStatus.textContent = '일괄 변환을 시작합니다...';
    let ok = 0;
    for (let i = 0; i < docQueue.length; i++) {
      if (!docQueue[i].pdfBlob) {
        await convertOneDoc(i);
      }
      if (docQueue[i].pdfBlob) ok++;
    }
    docStatus.textContent = `변환 완료: 총 ${docQueue.length}개 중 ${ok}개 성공`;
    renderDocQueue();
  });

  downloadAllZipBtn.addEventListener('click', async () => {
    const doneItems = docQueue.filter((it) => it.pdfBlob);
    if (!doneItems.length) return;
    const zip = new window.JSZip();
    doneItems.forEach((it) => {
      zip.file(baseName(it.file.name) + '.pdf', it.pdfBlob);
    });
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(zipBlob, `TongPDF_일괄변환_${new Date().toISOString().slice(0, 10)}.zip`);
  });

  $('#clearDocsBtn').addEventListener('click', () => {
    docQueue.length = 0;
    docStatus.textContent = '';
    renderDocQueue();
  });

  // -------------------- TAB 2: Images <-> PDF --------------------
  const imgQueue = [];
  const img2pdfList = $('#img2pdfList');
  const img2pdfRunBtn = $('#img2pdfRunBtn');
  const img2pdfStatus = $('#img2pdfStatus');

  function renderImgQueue() {
    img2pdfList.innerHTML = '';
    img2pdfRunBtn.disabled = imgQueue.length === 0;
    imgQueue.forEach((file, idx) => {
      const div = document.createElement('div');
      div.className = 'file-item';
      div.innerHTML = `
        <div class="file-meta">
          <span class="file-ext img">IMG</span>
          <div>
            <div class="file-name">${idx + 1}. ${escapeHtml(file.name)}</div>
            <div class="file-sub">${formatBytes(file.size)}</div>
          </div>
        </div>
        <div class="file-actions">
          <button type="button" class="btn btn-outline btn-sm" data-up="${idx}" ${idx === 0 ? 'disabled' : ''}>▲</button>
          <button type="button" class="btn btn-outline btn-sm" data-down="${idx}" ${idx === imgQueue.length - 1 ? 'disabled' : ''}>▼</button>
          <button type="button" class="btn btn-danger-outline btn-sm" data-del="${idx}">삭제</button>
        </div>
      `;
      img2pdfList.appendChild(div);
    });
    img2pdfList.querySelectorAll('[data-up]').forEach((b) =>
      b.addEventListener('click', () => {
        const i = Number(b.dataset.up);
        [imgQueue[i - 1], imgQueue[i]] = [imgQueue[i], imgQueue[i - 1]];
        renderImgQueue();
      })
    );
    img2pdfList.querySelectorAll('[data-down]').forEach((b) =>
      b.addEventListener('click', () => {
        const i = Number(b.dataset.down);
        [imgQueue[i + 1], imgQueue[i]] = [imgQueue[i], imgQueue[i + 1]];
        renderImgQueue();
      })
    );
    img2pdfList.querySelectorAll('[data-del]').forEach((b) =>
      b.addEventListener('click', () => {
        imgQueue.splice(Number(b.dataset.del), 1);
        renderImgQueue();
      })
    );
  }

  wireDropzone($('#img2pdfDropzone'), $('#img2pdfInput'), (files) => {
    files.forEach((f) => imgQueue.push(f));
    img2pdfStatus.textContent = `이미지 ${imgQueue.length}장 준비 완료`;
    renderImgQueue();
  });

  img2pdfRunBtn.addEventListener('click', async () => {
    img2pdfRunBtn.disabled = true;
    img2pdfStatus.textContent = 'PDF 생성 중...';
    try {
      const mode = $('#imgPageMode').value;
      const pdfDoc = await window.PDFLib.PDFDocument.create();
      for (const file of imgQueue) {
        const imgEl = await loadImageElement(file);
        const cv = document.createElement('canvas');
        cv.width = imgEl.width;
        cv.height = imgEl.height;
        cv.getContext('2d').drawImage(imgEl, 0, 0);
        const jpgBytes = await fetch(cv.toDataURL('image/jpeg', 0.92)).then((r) => r.arrayBuffer());
        const embedded = await pdfDoc.embedJpg(jpgBytes);

        if (mode === 'fit') {
          const page = pdfDoc.addPage([embedded.width, embedded.height]);
          page.drawImage(embedded, { x: 0, y: 0, width: embedded.width, height: embedded.height });
        } else {
          const a4W = 595.28;
          const a4H = 841.89;
          const margin = 24;
          const page = pdfDoc.addPage([a4W, a4H]);
          const ratio = Math.min((a4W - margin * 2) / embedded.width, (a4H - margin * 2) / embedded.height);
          const w = embedded.width * ratio;
          const h = embedded.height * ratio;
          page.drawImage(embedded, { x: (a4W - w) / 2, y: (a4H - h) / 2, width: w, height: h });
        }
      }
      const bytes = await pdfDoc.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), `이미지모음_${new Date().toISOString().slice(0, 10)}.pdf`);
      img2pdfStatus.textContent = 'PDF 생성 및 다운로드 완료!';
    } catch (err) {
      img2pdfStatus.textContent = '오류: ' + err.message;
    }
    img2pdfRunBtn.disabled = false;
  });

  // PDF -> Images
  let pdf2imgFile = null;
  const pdf2imgRunBtn = $('#pdf2imgRunBtn');
  const pdf2imgStatus = $('#pdf2imgStatus');

  wireDropzone($('#pdf2imgDropzone'), $('#pdf2imgInput'), (files) => {
    pdf2imgFile = files[0];
    pdf2imgRunBtn.disabled = !pdf2imgFile;
    pdf2imgStatus.textContent = pdf2imgFile ? `선택됨: ${pdf2imgFile.name} (${formatBytes(pdf2imgFile.size)})` : '';
  });

  pdf2imgRunBtn.addEventListener('click', async () => {
    if (!pdf2imgFile) return;
    pdf2imgRunBtn.disabled = true;
    pdf2imgStatus.textContent = 'PDF 페이지를 이미지로 렌더링 중...';
    try {
      const scale = parseFloat($('#pdf2imgScale').value) || 2.0;
      const buf = await pdf2imgFile.arrayBuffer();
      const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
      const zip = new window.JSZip();
      const prefix = baseName(pdf2imgFile.name);

      for (let p = 1; p <= pdf.numPages; p++) {
        pdf2imgStatus.textContent = `페이지 변환 중... (${p} / ${pdf.numPages})`;
        const page = await pdf.getPage(p);
        const vp = page.getViewport({ scale });
        const cv = document.createElement('canvas');
        cv.width = vp.width;
        cv.height = vp.height;
        await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
        const blob = await new Promise((res) => cv.toBlob(res, 'image/png'));
        zip.file(`${prefix}_p${String(p).padStart(3, '0')}.png`, blob);
      }
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      downloadBlob(zipBlob, `${prefix}_이미지.zip`);
      pdf2imgStatus.textContent = `완료! 총 ${pdf.numPages}페이지 이미지 ZIP 다운로드됨.`;
    } catch (err) {
      pdf2imgStatus.textContent = '오류: ' + err.message;
    }
    pdf2imgRunBtn.disabled = false;
  });

  // -------------------- TAB 3: PDF Merge & Split --------------------
  const mergeQueue = [];
  const mergeList = $('#mergeList');
  const mergeRunBtn = $('#mergeRunBtn');
  const mergeStatus = $('#mergeStatus');

  function renderMergeQueue() {
    mergeList.innerHTML = '';
    mergeRunBtn.disabled = mergeQueue.length < 2;
    mergeQueue.forEach((file, idx) => {
      const div = document.createElement('div');
      div.className = 'file-item';
      div.innerHTML = `
        <div class="file-meta">
          <span class="file-ext pdf">PDF</span>
          <div>
            <div class="file-name">${idx + 1}. ${escapeHtml(file.name)}</div>
            <div class="file-sub">${formatBytes(file.size)}</div>
          </div>
        </div>
        <div class="file-actions">
          <button type="button" class="btn btn-outline btn-sm" data-mup="${idx}" ${idx === 0 ? 'disabled' : ''}>▲</button>
          <button type="button" class="btn btn-outline btn-sm" data-mdown="${idx}" ${idx === mergeQueue.length - 1 ? 'disabled' : ''}>▼</button>
          <button type="button" class="btn btn-danger-outline btn-sm" data-mdel="${idx}">삭제</button>
        </div>
      `;
      mergeList.appendChild(div);
    });
    mergeList.querySelectorAll('[data-mup]').forEach((b) =>
      b.addEventListener('click', () => {
        const i = Number(b.dataset.mup);
        [mergeQueue[i - 1], mergeQueue[i]] = [mergeQueue[i], mergeQueue[i - 1]];
        renderMergeQueue();
      })
    );
    mergeList.querySelectorAll('[data-mdown]').forEach((b) =>
      b.addEventListener('click', () => {
        const i = Number(b.dataset.mdown);
        [mergeQueue[i + 1], mergeQueue[i]] = [mergeQueue[i], mergeQueue[i + 1]];
        renderMergeQueue();
      })
    );
    mergeList.querySelectorAll('[data-mdel]').forEach((b) =>
      b.addEventListener('click', () => {
        mergeQueue.splice(Number(b.dataset.mdel), 1);
        renderMergeQueue();
      })
    );
  }

  wireDropzone($('#mergeDropzone'), $('#mergeInput'), (files) => {
    files.forEach((f) => mergeQueue.push(f));
    mergeStatus.textContent = `PDF ${mergeQueue.length}개 선택됨`;
    renderMergeQueue();
  });

  $('#mergeClearBtn').addEventListener('click', () => {
    mergeQueue.length = 0;
    mergeStatus.textContent = '';
    renderMergeQueue();
  });

  mergeRunBtn.addEventListener('click', async () => {
    mergeRunBtn.disabled = true;
    mergeStatus.textContent = 'PDF 합치는 중...';
    try {
      const merged = await window.PDFLib.PDFDocument.create();
      for (const file of mergeQueue) {
        const src = await window.PDFLib.PDFDocument.load(await file.arrayBuffer());
        const copied = await merged.copyPages(src, src.getPageIndices());
        copied.forEach((p) => merged.addPage(p));
      }
      const bytes = await merged.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), `병합문서_${new Date().toISOString().slice(0, 10)}.pdf`);
      mergeStatus.textContent = `완료! ${mergeQueue.length}개 PDF를 하나로 합쳤습니다.`;
    } catch (err) {
      mergeStatus.textContent = '오류: ' + err.message;
    }
    mergeRunBtn.disabled = false;
  });

  // PDF Split
  let splitFile = null;
  let splitTotalPages = 0;
  const splitRangeBtn = $('#splitRangeBtn');
  const splitEachBtn = $('#splitEachBtn');
  const splitStatus = $('#splitStatus');

  wireDropzone($('#splitDropzone'), $('#splitInput'), async (files) => {
    splitFile = files[0];
    if (!splitFile) return;
    try {
      const doc = await window.PDFLib.PDFDocument.load(await splitFile.arrayBuffer());
      splitTotalPages = doc.getPageCount();
      $('#splitFileInfo').textContent = `${splitFile.name} · 총 ${splitTotalPages}페이지`;
      splitRangeBtn.disabled = false;
      splitEachBtn.disabled = false;
      splitStatus.textContent = `총 ${splitTotalPages}페이지 로드 완료`;
    } catch (err) {
      splitStatus.textContent = 'PDF 읽기 오류: ' + err.message;
    }
  });

  function parsePageRanges(input, maxPages) {
    const pages = new Set();
    const parts = input.split(',').map((s) => s.trim()).filter(Boolean);
    for (const part of parts) {
      if (part.includes('-')) {
        const [a, b] = part.split('-').map((n) => parseInt(n.trim(), 10));
        if (!isNaN(a) && !isNaN(b)) {
          for (let p = Math.max(1, a); p <= Math.min(maxPages, b); p++) pages.add(p - 1);
        }
      } else {
        const n = parseInt(part, 10);
        if (!isNaN(n) && n >= 1 && n <= maxPages) pages.add(n - 1);
      }
    }
    return Array.from(pages).sort((a, b) => a - b);
  }

  splitRangeBtn.addEventListener('click', async () => {
    if (!splitFile) return;
    const raw = $('#splitRangeInput').value.trim();
    if (!raw) {
      splitStatus.textContent = '추출할 페이지 범위를 입력하세요 (예: 1-3, 5). 전체 분리는 오른쪽 버튼을 눌러주세요.';
      return;
    }
    const indices = parsePageRanges(raw, splitTotalPages);
    if (!indices.length) {
      splitStatus.textContent = `유효한 페이지 번호가 없습니다 (1~${splitTotalPages}).`;
      return;
    }
    try {
      const src = await window.PDFLib.PDFDocument.load(await splitFile.arrayBuffer());
      const out = await window.PDFLib.PDFDocument.create();
      const pages = await out.copyPages(src, indices);
      pages.forEach((p) => out.addPage(p));
      const bytes = await out.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), `${baseName(splitFile.name)}_추출.pdf`);
      splitStatus.textContent = `${indices.length}페이지 추출 PDF 저장 완료!`;
    } catch (err) {
      splitStatus.textContent = '오류: ' + err.message;
    }
  });

  splitEachBtn.addEventListener('click', async () => {
    if (!splitFile) return;
    splitEachBtn.disabled = true;
    splitStatus.textContent = '1페이지씩 분리 중...';
    try {
      const src = await window.PDFLib.PDFDocument.load(await splitFile.arrayBuffer());
      const zip = new window.JSZip();
      const prefix = baseName(splitFile.name);
      for (let i = 0; i < src.getPageCount(); i++) {
        const single = await window.PDFLib.PDFDocument.create();
        const [p] = await single.copyPages(src, [i]);
        single.addPage(p);
        const bytes = await single.save();
        zip.file(`${prefix}_p${String(i + 1).padStart(3, '0')}.pdf`, bytes);
      }
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      downloadBlob(zipBlob, `${prefix}_페이지분리.zip`);
      splitStatus.textContent = `총 ${src.getPageCount()}개 PDF로 분리하여 ZIP 저장 완료!`;
    } catch (err) {
      splitStatus.textContent = '오류: ' + err.message;
    }
    splitEachBtn.disabled = false;
  });

  // -------------------- Keyboard Shortcuts --------------------
  const tabOrder = ['doc2pdf', 'imgPdf', 'mergeSplit', 'pcTool'];
  let currentTabId = 'doc2pdf';

  const origSwitchTab = switchTab;
  switchTab = function (tabId) {
    currentTabId = tabId;
    origSwitchTab(tabId);
  };

  document.addEventListener('keydown', (e) => {
    const tag = (e.target && e.target.tagName) || '';
    const isTyping = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';

    // F1 -> Open Guide
    if (e.key === 'F1') {
      e.preventDefault();
      window.open('./guide.html', '_blank', 'noopener,noreferrer');
      return;
    }

    // Alt + 1..4 -> Switch tabs
    if (e.altKey && !e.ctrlKey && !e.metaKey && ['1', '2', '3', '4'].includes(e.key)) {
      e.preventDefault();
      const idx = Number(e.key) - 1;
      if (tabOrder[idx]) switchTab(tabOrder[idx]);
      return;
    }

    // Ctrl/Cmd + O -> Open file picker for active tab
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'o') {
      e.preventDefault();
      if (currentTabId === 'doc2pdf') $('#docFileInput').click();
      else if (currentTabId === 'imgPdf') $('#img2pdfInput').click();
      else if (currentTabId === 'mergeSplit') $('#mergeInput').click();
      return;
    }

    // Ctrl/Cmd + Enter -> Trigger primary convert/action on active tab
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && !isTyping) {
      e.preventDefault();
      if (currentTabId === 'doc2pdf' && !convertAllDocsBtn.disabled) convertAllDocsBtn.click();
      else if (currentTabId === 'imgPdf' && !img2pdfRunBtn.disabled) img2pdfRunBtn.click();
      else if (currentTabId === 'mergeSplit' && !mergeRunBtn.disabled) mergeRunBtn.click();
      return;
    }

    // Ctrl/Cmd + Shift + D -> Download all converted PDFs as ZIP
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'd') {
      if (currentTabId === 'doc2pdf' && !downloadAllZipBtn.disabled) {
        e.preventDefault();
        downloadAllZipBtn.click();
      }
    }
  });
})();
