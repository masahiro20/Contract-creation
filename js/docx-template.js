/*
 * Word(.docx) ひな形エンジン
 *
 * 方針：元の Word ファイルをそのまま使い、差し込み箇所の「文字」だけを書き換える。
 *  - 書式（フォント・サイズ・行間・字下げ・表・余白など）には一切手を加えない
 *  - 差し込んだ文字は、その場所にもともとあった文字の書式（run）をそのまま引き継ぐ
 *  - 差し込みのない部品（画像・スタイル・ヘッダー等）は ZIP のバイト列ごとそのまま戻す
 */
(function () {
  'use strict';

  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const TEXT_PART_RE = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
  const PH_RE = /\{\{\s*([^{}:]+?)\s*(?::\s*([a-z]+)\s*)?\}\}/g;

  // ---------- DOM ヘルパー ----------
  const isW = (el, name) => el && el.nodeType === 1 && el.namespaceURI === W && el.localName === name;
  const child = (el, name) => {
    if (!el) return null;
    for (const c of el.childNodes) if (isW(c, name)) return c;
    return null;
  };
  const children = (el, name) => (el ? Array.from(el.childNodes).filter((c) => isW(c, name)) : []);
  const attr = (el, name) => (el ? el.getAttributeNS(W, name) || el.getAttribute('w:' + name) || '' : '');
  const closest = (node, name) => {
    for (let n = node.parentNode; n && n.nodeType === 1; n = n.parentNode) if (isW(n, name)) return n;
    return null;
  };

  /** 段落内の要素（文字・タブ・改行）を文書順に集める。入れ子の段落（テキストボックス）は含めない */
  function paraItems(p) {
    const items = [];
    const walk = (node) => {
      for (const c of node.childNodes) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === MC && c.localName === 'Fallback') continue;
        if (c.namespaceURI === W) {
          const n = c.localName;
          if (n === 'p' || n === 'pPr' || n === 'rPr' || n === 'delText' || n === 'instrText') continue;
          if (n === 't') {
            items.push({ type: 't', el: c });
            continue;
          }
          if (n === 'tab' && isW(c.parentNode, 'r')) {
            items.push({ type: 'tab', el: c });
            continue;
          }
          if (n === 'br' || n === 'cr') {
            items.push({ type: attr(c, 'type') === 'page' ? 'page' : 'br', el: c });
            continue;
          }
        }
        walk(c);
      }
    };
    walk(p);
    return items;
  }

  const textNodes = (p) => paraItems(p).filter((i) => i.type === 't').map((i) => i.el);
  const paraText = (p) => textNodes(p).map((t) => t.textContent).join('');

  function setText(t, text) {
    t.textContent = text;
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  }

  /** w:t に文字を入れる。改行を含む場合は同じ run の中に w:br を挟む（書式は同じ run のまま） */
  function writeText(t, text) {
    const lines = text.split('\n');
    setText(t, lines[0]);
    let after = t;
    for (let i = 1; i < lines.length; i++) {
      const br = t.ownerDocument.createElementNS(W, 'w:br');
      const nt = t.ownerDocument.createElementNS(W, 'w:t');
      setText(nt, lines[i]);
      after.after(br, nt);
      after = nt;
    }
  }

  /**
   * 段落の文字列の [s, e) を text に置き換える。
   * 置換後の文字は、範囲の先頭文字があった run の書式をそのまま使う。
   */
  function replaceRange(p, s, e, text) {
    const ts = textNodes(p);
    let pos = 0;
    let si = -1;
    let so = 0;
    let ei = -1;
    let eo = 0;
    for (let i = 0; i < ts.length; i++) {
      const len = ts[i].textContent.length;
      if (si < 0 && s < pos + len) (si = i), (so = s - pos);
      if (si >= 0 && ei < 0 && e <= pos + len) (ei = i), (eo = e - pos);
      pos += len;
      if (ei >= 0) break;
    }
    if (si < 0 || ei < 0) throw new Error('置換範囲が不正です');
    if (si === ei) {
      const v = ts[si].textContent;
      writeText(ts[si], v.slice(0, so) + text + v.slice(eo));
      return;
    }
    const tail = ts[ei].textContent.slice(eo);
    for (let k = si + 1; k <= ei; k++) setText(ts[k], '');
    setText(ts[ei], tail);
    writeText(ts[si], ts[si].textContent.slice(0, so) + text);
  }

  function findPlaceholders(text) {
    const out = [];
    let m;
    PH_RE.lastIndex = 0;
    while ((m = PH_RE.exec(text))) out.push({ key: m[1].trim(), type: m[2] || '', start: m.index, end: m.index + m[0].length });
    return out;
  }

  // ---------- スタイル（プレビュー用） ----------
  function parseStyles(doc) {
    const styles = new Map();
    const defaults = { pPr: null, rPr: null };
    if (!doc) return { styles, defaults, defaultPara: null };
    const dd = doc.getElementsByTagNameNS(W, 'docDefaults')[0];
    if (dd) {
      defaults.rPr = child(child(dd, 'rPrDefault'), 'rPr');
      defaults.pPr = child(child(dd, 'pPrDefault'), 'pPr');
    }
    let defaultPara = null;
    for (const s of doc.getElementsByTagNameNS(W, 'style')) {
      const id = attr(s, 'styleId');
      styles.set(id, s);
      if (attr(s, 'type') === 'paragraph' && ['1', 'true'].includes(attr(s, 'default'))) defaultPara = id;
    }
    return { styles, defaults, defaultPara };
  }

  // ---------- 本体 ----------
  class DocxTemplate {
    static async load(buffer) {
      const t = new DocxTemplate();
      t.entries = Zip.read(buffer);
      t.parts = new Map();
      const parser = new DOMParser();
      const decoder = new TextDecoder();
      const docEntry = t.entries.find((e) => e.name === 'word/document.xml');
      if (!docEntry) throw new Error('Word(.docx) ファイルではありません');
      for (const e of t.entries) {
        if (!TEXT_PART_RE.test(e.name) && e.name !== 'word/styles.xml') continue;
        const xml = decoder.decode(await Zip.inflate(e));
        const doc = parser.parseFromString(xml, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) throw new Error(e.name + ' を読み込めませんでした');
        t.parts.set(e.name, { doc, dirty: false, xml });
      }
      t.styleInfo = parseStyles(t.parts.get('word/styles.xml') && t.parts.get('word/styles.xml').doc);
      return t;
    }

    /** 差し込み対象になり得る段落（本文 → ヘッダー/フッター等の順） */
    paragraphs() {
      const names = Array.from(this.parts.keys()).filter((n) => TEXT_PART_RE.test(n));
      names.sort((a, b) => (a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a.localeCompare(b)));
      const out = [];
      for (const name of names) {
        for (const p of this.parts.get(name).doc.getElementsByTagNameNS(W, 'p')) out.push({ part: name, p });
      }
      return out;
    }

    /** 文書中の {{項目}} を出現順に返す */
    fields() {
      const seen = new Map();
      for (const { p } of this.paragraphs()) {
        for (const ph of findPlaceholders(paraText(p))) {
          if (!seen.has(ph.key)) seen.set(ph.key, { key: ph.key, type: ph.type });
          else if (ph.type && !seen.get(ph.key).type) seen.get(ph.key).type = ph.type;
        }
      }
      return Array.from(seen.values());
    }

    /** 全段落の文字列（検証用） */
    texts() {
      return this.paragraphs().map(({ p }) => paraText(p));
    }

    markDirty(p) {
      for (const [, part] of this.parts) if (part.doc === p.ownerDocument) part.dirty = true;
    }

    /** 段落 p の [s,e) を {{key}} に置き換える（ひな形づくり用） */
    markRange(p, s, e, key) {
      replaceRange(p, s, e, `{{${key}}}`);
      this.markDirty(p);
    }

    /** 文書中の text（既存の {{…}} の中は除く）をすべて {{key}} に置き換え、件数を返す */
    markAll(text, key) {
      let n = 0;
      for (const { p } of this.paragraphs()) {
        const s = paraText(p);
        const phs = findPlaceholders(s);
        const hits = [];
        let i = s.indexOf(text);
        while (i >= 0) {
          const end = i + text.length;
          if (!phs.some((ph) => i < ph.end && end > ph.start)) hits.push(i);
          i = s.indexOf(text, end);
        }
        for (const h of hits.reverse()) replaceRange(p, h, h + text.length, `{{${key}}}`);
        if (hits.length) this.markDirty(p);
        n += hits.length;
      }
      return n;
    }

    countText(text) {
      let n = 0;
      for (const { p } of this.paragraphs()) {
        const s = paraText(p);
        const phs = findPlaceholders(s);
        let i = s.indexOf(text);
        while (i >= 0) {
          const end = i + text.length;
          if (!phs.some((ph) => i < ph.end && end > ph.start)) n++;
          i = s.indexOf(text, end);
        }
      }
      return n;
    }

    /** {{key}} を別の文字に置き換える（valueOf(key) が文字列を返す）。差し込み・項目解除の両方で使う */
    fill(valueOf) {
      for (const { p } of this.paragraphs()) {
        const phs = findPlaceholders(paraText(p));
        let changed = false;
        for (const ph of phs.reverse()) {
          const v = valueOf(ph.key);
          if (v == null) continue;
          replaceRange(p, ph.start, ph.end, String(v));
          changed = true;
        }
        if (changed) this.markDirty(p);
      }
    }

    /** 元に戻す用に、文字部品の状態を保存・復元 */
    snapshot() {
      const s = new XMLSerializer();
      const snap = {};
      for (const [name, part] of this.parts) if (part.dirty) snap[name] = s.serializeToString(part.doc);
      return snap;
    }

    restore(snap) {
      const parser = new DOMParser();
      for (const [name, part] of this.parts) {
        if (!TEXT_PART_RE.test(name)) continue;
        const xml = snap[name] || part.xml;
        part.doc = parser.parseFromString(xml, 'application/xml');
        part.dirty = !!snap[name];
      }
    }

    /** 変更した部品だけを書き出し、それ以外は元のバイト列のまま ZIP にする */
    async toBlob() {
      const s = new XMLSerializer();
      const enc = new TextEncoder();
      const out = [];
      for (const e of this.entries) {
        const part = this.parts.get(e.name);
        if (part && part.dirty) {
          let xml = s.serializeToString(part.doc).replace(/^<\?xml[^>]*\?>\s*/, '');
          xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
          out.push(await Zip.replaceData(e, enc.encode(xml)));
        } else {
          out.push(e);
        }
      }
      return Zip.write(out, DOCX_MIME);
    }

    async toArrayBuffer() {
      return (await this.toBlob()).arrayBuffer();
    }

    // =================================================================
    // プレビュー（確認用の簡易表示。出力される Word には影響しない）
    // =================================================================
    /**
     * opts.mode: 'raw'  … ひな形づくり用。文字を選択できるよう位置情報を付ける
     *            'fill' … 入力内容を差し込んだ状態で表示
     * opts.value(key) → { text, empty, label }   ('fill' のとき)
     */
    renderHtml(opts) {
      this._pmap = [];
      const doc = this.parts.get('word/document.xml').doc;
      const body = doc.getElementsByTagNameNS(W, 'body')[0];
      const sect = body ? child(body, 'sectPr') : null;
      const pg = child(sect, 'pgSz');
      const mar = child(sect, 'pgMar');
      const tw = (v, d) => (v ? +v / 20 : d);
      const page = {
        width: tw(attr(pg, 'w'), 595),
        top: tw(attr(mar, 'top'), 72),
        right: tw(attr(mar, 'right'), 72),
        bottom: tw(attr(mar, 'bottom'), 72),
        left: tw(attr(mar, 'left'), 72),
      };
      const defaultRun = this._runCss([]);
      const extras = (re, label) => {
        let h = '';
        for (const [name, part] of this.parts) {
          if (!re.test(name)) continue;
          const html = this._renderBlock(part.doc.documentElement, opts);
          if (html.replace(/<[^>]+>|&nbsp;|\s/g, '')) h += `<div class="dx-hf"><span class="dx-hf-label">${label}</span>${html}</div>`;
        }
        return h;
      };
      return {
        page,
        html:
          `<div class="dx-page" style="width:${page.width}pt;padding:${page.top}pt ${page.right}pt ${page.bottom}pt ${page.left}pt;${defaultRun}">` +
          extras(/^word\/header\d*\.xml$/, 'ヘッダー') +
          this._renderBlock(body, opts) +
          extras(/^word\/footer\d*\.xml$/, 'フッター') +
          `</div>`,
      };
    }

    paragraphById(pid) {
      return this._pmap ? this._pmap[pid] : null;
    }

    _renderBlock(container, opts) {
      let h = '';
      for (const c of container.childNodes) {
        if (isW(c, 'p')) h += this._renderPara(c, opts);
        else if (isW(c, 'tbl')) h += this._renderTable(c, opts);
        else if (isW(c, 'sdt')) h += this._renderBlock(child(c, 'sdtContent') || c, opts);
        else if (c.nodeType === 1 && (c.localName === 'customXml' || c.localName === 'sdtContent')) h += this._renderBlock(c, opts);
      }
      return h;
    }

    _renderTable(tbl, opts) {
      let h = '<table class="dx-tbl">';
      for (const tr of children(tbl, 'tr')) {
        h += '<tr>';
        for (const tc of children(tr, 'tc')) {
          const pr = child(tc, 'tcPr');
          const span = attr(child(pr, 'gridSpan'), 'val');
          const vm = child(pr, 'vMerge');
          const cont = vm && attr(vm, 'val') !== 'restart';
          const w = child(pr, 'tcW');
          const width = w && attr(w, 'type') === 'dxa' && attr(w, 'w') ? `width:${+attr(w, 'w') / 20}pt;` : '';
          const cls = cont ? ' class="dx-vcont"' : vm ? ' class="dx-vstart"' : '';
          h += `<td${span ? ` colspan="${span}"` : ''}${cls} style="${width}">${cont ? '' : this._renderBlock(tc, opts)}</td>`;
        }
        h += '</tr>';
      }
      return h + '</table>';
    }

    _styleChain(id, kind) {
      const out = [];
      const { styles } = this.styleInfo;
      let cur = id;
      for (let i = 0; cur && i < 10; i++) {
        const s = styles.get(cur);
        if (!s) break;
        const pr = child(s, kind);
        if (pr) out.push(pr);
        cur = attr(child(s, 'basedOn'), 'val');
      }
      return out;
    }

    _pick(list, name) {
      for (const pr of list) {
        const c = child(pr, name);
        if (c) return c;
      }
      return null;
    }

    _paraPrList(p) {
      const pPr = child(p, 'pPr');
      const sid = attr(child(pPr, 'pStyle'), 'val') || this.styleInfo.defaultPara;
      const list = [];
      if (pPr) list.push(pPr);
      list.push(...this._styleChain(sid, 'pPr'));
      if (this.styleInfo.defaults.pPr) list.push(this.styleInfo.defaults.pPr);
      return { list, sid };
    }

    _runCss(rPrList) {
      const all = [...rPrList];
      if (this.styleInfo.defaults.rPr) all.push(this.styleInfo.defaults.rPr);
      const on = (name) => {
        const el = this._pick(all, name);
        return el && !['0', 'false', 'off'].includes(attr(el, 'val'));
      };
      let css = '';
      const sz = attr(this._pick(all, 'sz'), 'val');
      css += `font-size:${sz ? +sz / 2 : 10.5}pt;`;
      const fonts = this._pick(all, 'rFonts');
      const ea = attr(fonts, 'eastAsia');
      const ascii = attr(fonts, 'ascii');
      css += `font-family:${fontStack(ascii, ea)};`;
      if (on('b')) css += 'font-weight:bold;';
      if (on('i')) css += 'font-style:italic;';
      const u = this._pick(all, 'u');
      const deco = [];
      if (u && attr(u, 'val') !== 'none') deco.push('underline');
      if (on('strike') || on('dstrike')) deco.push('line-through');
      if (deco.length) css += `text-decoration:${deco.join(' ')};`;
      const color = attr(this._pick(all, 'color'), 'val');
      if (color && color !== 'auto') css += `color:#${color};`;
      const va = attr(this._pick(all, 'vertAlign'), 'val');
      if (va === 'superscript') css += 'vertical-align:super;font-size:smaller;';
      if (va === 'subscript') css += 'vertical-align:sub;font-size:smaller;';
      return css;
    }

    _renderPara(p, opts) {
      const { list, sid } = this._paraPrList(p);
      const pPr = child(p, 'pPr');
      let css = '';
      const jc = attr(this._pick(list, 'jc'), 'val');
      if (jc === 'center') css += 'text-align:center;';
      else if (jc === 'right' || jc === 'end') css += 'text-align:right;';
      else if (jc === 'both') css += 'text-align:justify;';
      else if (jc === 'distribute') css += 'text-align:justify;text-align-last:justify;';
      const ind = this._pick(list, 'ind');
      if (ind) {
        const chars = (n) => (attr(ind, n) ? +attr(ind, n) / 100 : null);
        const twip = (n) => (attr(ind, n) ? +attr(ind, n) / 20 : null);
        const left = chars('leftChars') != null ? chars('leftChars') + 'em' : twip('left') != null ? twip('left') + 'pt' : twip('start') != null ? twip('start') + 'pt' : null;
        const right = chars('rightChars') != null ? chars('rightChars') + 'em' : twip('right') != null ? twip('right') + 'pt' : null;
        if (left) css += `padding-left:${left};`;
        if (right) css += `padding-right:${right};`;
        if (chars('hangingChars') != null) css += `text-indent:-${chars('hangingChars')}em;`;
        else if (twip('hanging') != null) css += `text-indent:-${twip('hanging')}pt;`;
        else if (chars('firstLineChars') != null) css += `text-indent:${chars('firstLineChars')}em;`;
        else if (twip('firstLine') != null) css += `text-indent:${twip('firstLine')}pt;`;
      }
      const sp = this._pick(list, 'spacing');
      if (sp) {
        if (attr(sp, 'before')) css += `margin-top:${+attr(sp, 'before') / 20}pt;`;
        if (attr(sp, 'after')) css += `margin-bottom:${+attr(sp, 'after') / 20}pt;`;
        const line = +attr(sp, 'line');
        const rule = attr(sp, 'lineRule');
        if (line && (!rule || rule === 'auto')) css += `line-height:${(line / 240) * 1.5};`;
        else if (line) css += `line-height:${line / 20}pt;`;
      }
      const pageBefore = child(pPr, 'pageBreakBefore');
      const paraRPr = this._styleChain(sid, 'rPr');

      const pid = this._pmap.push(p) - 1;
      const items = paraItems(p);
      const text = items.filter((i) => i.type === 't').map((i) => i.el.textContent).join('');
      const phs = findPlaceholders(text);
      const phAt = (g) => phs.find((ph) => g >= ph.start && g < ph.end);

      let h = '';
      let g = 0;
      let tIndex = 0;
      for (const it of items) {
        const run = closest(it.el, 'r');
        const rcss = this._runCss([...(run ? [child(run, 'rPr')].filter(Boolean) : []), ...(run ? this._styleChain(attr(child(child(run, 'rPr'), 'rStyle'), 'val'), 'rPr') : []), ...paraRPr]);
        if (it.type === 'tab') {
          h += `<span class="dx-tab" style="${rcss}">\t</span>`;
          continue;
        }
        if (it.type === 'br') {
          h += '<br>';
          continue;
        }
        if (it.type === 'page') {
          h += '<span class="dx-pagebreak"></span>';
          continue;
        }
        const s = it.el.textContent;
        const ti = tIndex++;
        // 差し込み位置の境界で分割
        let i = 0;
        while (i < s.length) {
          const ph = phAt(g + i);
          let j;
          if (ph) j = Math.min(s.length, ph.end - g);
          else {
            const next = phs.find((x) => x.start > g + i);
            j = next ? Math.min(s.length, next.start - g) : s.length;
          }
          const seg = s.slice(i, j);
          if (opts.mode === 'raw') {
            h += `<span data-t="${ti}" data-o="${i}"${ph ? ` class="tok" data-key="${esc(ph.key)}"` : ''} style="${rcss}">${esc(seg)}</span>`;
          } else if (ph) {
            if (g + i === ph.start) {
              const v = opts.value(ph.key);
              h += `<span class="ph ${v.empty ? 'empty' : 'filled'}" data-key="${esc(ph.key)}" style="${rcss}">${esc(v.empty ? `【${v.label}】` : v.text).replace(/\n/g, '<br>')}</span>`;
            }
          } else {
            h += `<span style="${rcss}">${esc(seg)}</span>`;
          }
          i = j;
        }
        g += s.length;
      }
      const rcssEmpty = this._runCss(paraRPr);
      return `${pageBefore ? '<div class="dx-pagebreak"></div>' : ''}<p class="dx-p" data-pid="${pid}" style="${css}">${h || `<span style="${rcssEmpty}">&#8203;</span>`}</p>`;
    }
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  const FONT_ALIASES = {
    'ＭＳ 明朝': '"MS Mincho","ＭＳ 明朝"',
    'ＭＳ Ｐ明朝': '"MS PMincho","ＭＳ Ｐ明朝"',
    'ＭＳ ゴシック': '"MS Gothic","ＭＳ ゴシック"',
    'ＭＳ Ｐゴシック': '"MS PGothic","ＭＳ Ｐゴシック"',
    游明朝: '"Yu Mincho","YuMincho","游明朝"',
    游ゴシック: '"Yu Gothic","YuGothic","游ゴシック"',
  };
  function fontStack(ascii, ea) {
    const names = [];
    for (const f of [ascii, ea]) if (f) names.push(FONT_ALIASES[f] || `"${f.replace(/"/g, '')}"`);
    const serif = /明朝|Mincho|Century|Times|serif/i.test(`${ea} ${ascii}`) || (!ea && !ascii);
    names.push(serif ? '"Hiragino Mincho ProN","Yu Mincho","Noto Serif JP",serif' : '"Hiragino Sans","Yu Gothic","Noto Sans JP",sans-serif');
    return names.join(',').replace(/"/g, "'");
  }

  window.DocxTemplate = DocxTemplate;
  DocxTemplate.paraText = paraText;
  DocxTemplate.findPlaceholders = findPlaceholders;
})();
