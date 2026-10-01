/*
 * 契約書メーカー 本体
 *
 * 元の Word ファイルをひな形として登録し、変わる文字を {{項目名}} にしておくと、
 * 入力フォームの内容をその位置に差し込んだ Word を出力する。
 * 書式は元ファイルのまま（docx-template.js 参照）。
 */
(function () {
  'use strict';

  // =====================================================================
  // 定数
  // =====================================================================
  const TYPE_LABELS = {
    text: 'テキスト（1行）',
    textarea: 'テキスト（複数行）',
    date: '日付',
    money: '金額',
    number: '数値',
    select: '選択肢',
  };
  const NUMERIC_TYPES = ['date', 'money', 'number'];
  const DEFAULT_GROUP = '入力項目';
  const PRESETS = window.FIELD_PRESETS || [];
  const PRESET_MAP = new Map(PRESETS.map((p) => [p.key, p]));

  // =====================================================================
  // 保存
  // =====================================================================
  const store = {
    get(k, fallback) {
      try {
        const v = localStorage.getItem('cc.' + k);
        return v == null ? fallback : JSON.parse(v);
      } catch (e) {
        return fallback;
      }
    },
    set(k, v) {
      try {
        localStorage.setItem('cc.' + k, JSON.stringify(v));
      } catch (e) {
        /* noop */
      }
    },
    del(k) {
      try {
        localStorage.removeItem('cc.' + k);
      } catch (e) {
        /* noop */
      }
    },
  };

  // Word ファイル本体は容量が大きいので IndexedDB に保存する
  const idb = {
    db: null,
    open() {
      if (this.db) return Promise.resolve(this.db);
      return new Promise((resolve, reject) => {
        const req = indexedDB.open('contract-maker', 1);
        req.onupgradeneeded = () => req.result.createObjectStore('templates', { keyPath: 'id' });
        req.onsuccess = () => resolve((this.db = req.result));
        req.onerror = () => reject(req.error);
      });
    },
    async tx(mode, fn) {
      const db = await this.open();
      return new Promise((resolve, reject) => {
        const t = db.transaction('templates', mode);
        const r = fn(t.objectStore('templates'));
        t.oncomplete = () => resolve(r && r.result);
        t.onerror = () => reject(t.error);
      });
    },
    all() {
      return this.tx('readonly', (s) => s.getAll());
    },
    put(rec) {
      return this.tx('readwrite', (s) => s.put(rec));
    },
    del(id) {
      return this.tx('readwrite', (s) => s.delete(id));
    },
  };

  // =====================================================================
  // ひな形の登録
  // =====================================================================
  const builtinRecords = [];
  let customRecords = [];

  function b64ToBuffer(b64) {
    const bin = atob(b64);
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8.buffer;
  }

  function bufferToB64(buf) {
    const u8 = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    return btoa(bin);
  }

  function recordFromBundle(t, builtin) {
    return {
      id: String(t.id || 'tpl-' + Date.now().toString(36)),
      name: t.name || '無題のひな形',
      description: t.description || '',
      fileName: t.fileName || '',
      dateStyle: t.dateStyle === 'seireki' ? 'seireki' : 'wareki',
      fields: Array.isArray(t.fields) ? t.fields : [],
      bytes: b64ToBuffer(t.docx),
      builtin,
    };
  }

  /** templates/*.js（配布ひな形）から呼ばれる */
  function registerBundle(bundle) {
    for (const t of bundle.templates || []) {
      if (!t || !t.docx) continue;
      const rec = recordFromBundle(t, true);
      const i = builtinRecords.findIndex((r) => r.id === rec.id);
      if (i >= 0) builtinRecords[i] = rec;
      else builtinRecords.push(rec);
    }
  }

  function bundleOf(recs) {
    return {
      format: 'contract-maker',
      version: 2,
      templates: recs.map((r) => ({
        id: r.id,
        name: r.name,
        description: r.description,
        fileName: r.fileName,
        dateStyle: r.dateStyle,
        fields: r.fields,
        docx: bufferToB64(r.bytes),
      })),
    };
  }

  const allRecords = () => [...builtinRecords, ...customRecords];
  const findRecord = (id) => allRecords().find((r) => r.id === id) || null;

  async function reloadCustoms() {
    try {
      customRecords = ((await idb.all()) || []).map((r) => ({ ...r, builtin: false }));
    } catch (e) {
      customRecords = customRecords || [];
      toast('このブラウザではひな形を保存できません（プライベートモード等）');
    }
  }

  // =====================================================================
  // 入力項目
  // =====================================================================
  function makeField(key, detectedType, settings) {
    const f = { ...(PRESET_MAP.get(key) || {}), ...(settings || {}) };
    let type = f.type || detectedType || 'text';
    if (!TYPE_LABELS[type]) type = 'text';
    const options = Array.isArray(f.options)
      ? f.options
      : typeof f.options === 'string'
        ? f.options.split(/[,、\n]/).map((s) => s.trim()).filter(Boolean)
        : [];
    return {
      key,
      label: f.label || key,
      type,
      group: f.group || DEFAULT_GROUP,
      required: typeof f.compute === 'function' ? false : f.required !== undefined ? !!f.required : true,
      remember: !!f.remember,
      options,
      placeholder: f.placeholder || '',
      hint: f.hint || '',
      default: f.default,
      zenkaku: !!f.zenkaku,
      blank: typeof f.blank === 'string' ? f.blank : '',
      dropEmptyPara: !!f.dropEmptyPara,
      compute: typeof f.compute === 'function' ? f.compute : null,
      sample: f.sample || '',
    };
  }

  function buildFields(detected, settingsList) {
    const settings = settingsList instanceof Map ? settingsList : new Map((settingsList || []).map((f) => [f.key, f]));
    const fields = detected.map((d) => makeField(d.key, d.type, settings.get(d.key)));
    const count = {};
    for (const f of fields) count[f.label] = (count[f.label] || 0) + 1;
    for (const f of fields) f.markerLabel = count[f.label] > 1 ? `${f.group}・${f.label}` : f.label;
    return fields;
  }

  function serializeField(f) {
    const o = { key: f.key, label: f.label, type: f.type, group: f.group, required: f.required };
    if (f.remember) o.remember = true;
    if (f.options && f.options.length) o.options = f.options;
    if (f.placeholder) o.placeholder = f.placeholder;
    if (f.hint) o.hint = f.hint;
    if (f.default !== undefined) o.default = f.default;
    if (f.zenkaku) o.zenkaku = true;
    if (f.blank) o.blank = f.blank;
    if (f.dropEmptyPara) o.dropEmptyPara = true;
    if (f.sample) o.sample = f.sample;
    return o;
  }

  // =====================================================================
  // 値の整形（差し込む文字は、ここで作ったものがそのまま入る）
  // =====================================================================
  const ERAS = [
    { name: '令和', start: [2019, 5, 1] },
    { name: '平成', start: [1989, 1, 8] },
    { name: '昭和', start: [1926, 12, 25] },
  ];

  function formatDate(iso, style) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso));
    if (!m) return String(iso);
    const y = +m[1];
    const mo = +m[2];
    const d = +m[3];
    if (style === 'seireki') return `${y}年${mo}月${d}日`;
    const n = y * 10000 + mo * 100 + d;
    for (const era of ERAS) {
      if (n >= era.start[0] * 10000 + era.start[1] * 100 + era.start[2]) {
        const ey = y - era.start[0] + 1;
        return `${era.name}${ey === 1 ? '元' : ey}年${mo}月${d}日`;
      }
    }
    return `${y}年${mo}月${d}日`;
  }

  function digitsOnly(v) {
    return String(v == null ? '' : v)
      .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
      .replace(/[^\d.-]/g, '');
  }

  function formatMoney(v) {
    const s = digitsOnly(v);
    if (!s) return '';
    const [int, dec] = s.split('.');
    const withComma = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return dec ? `${withComma}.${dec}` : withComma;
  }

  function toZenkaku(s) {
    return s
      .replace(/[0-9]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0))
      .replace(/,/g, '，')
      .replace(/\./g, '．');
  }

  const isEmpty = (v) => v == null || v === '';

  /** 項目の値（自動計算の項目はほかの項目から計算する） */
  function rawValue(f, values) {
    if (!f.compute) return values[f.key];
    try {
      return f.compute(values);
    } catch (e) {
      return '';
    }
  }

  const dropKeysOf = (fields) => new Set(fields.filter((f) => f.dropEmptyPara).map((f) => f.key));

  function formatValue(f, raw, dateStyle) {
    if (isEmpty(raw)) return f.blank || ''; // 未入力時の文字（ひな形ごとに設定）
    let s;
    if (f.type === 'date') s = formatDate(raw, dateStyle);
    else if (f.type === 'money') s = formatMoney(raw);
    else s = String(raw);
    if (f.zenkaku && NUMERIC_TYPES.includes(f.type)) s = toZenkaku(s);
    return s;
  }

  function todayIso() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }

  // =====================================================================
  // 画面
  // =====================================================================
  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
  const esc = (s) =>
    String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  const state = {
    rec: null,
    doc: null,
    fields: [],
    fieldMap: new Map(),
    values: store.get('case', {}),
    dateStyle: 'wareki',
  };

  function toast(msg) {
    const el = $('#toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => el.classList.remove('show'), 2600);
  }

  // ---------- サイドバー ----------
  function renderTemplateList() {
    const item = (r) =>
      `<li><button type="button" class="tpl-item${state.rec && state.rec.id === r.id ? ' active' : ''}" data-id="${esc(r.id)}">` +
      `<span class="tpl-name">${esc(r.name)}</span>` +
      (r.description ? `<span class="tpl-desc">${esc(r.description)}</span>` : '') +
      `</button></li>`;
    let html = '';
    if (builtinRecords.length) html += `<li class="tpl-sep">配布ひな形</li>${builtinRecords.map(item).join('')}`;
    html += '<li class="tpl-sep">マイひな形</li>';
    html += customRecords.length
      ? customRecords.map(item).join('')
      : '<li class="tpl-none">まだありません。<br>「＋ Wordからひな形を作る」で追加できます。</li>';
    $('#tpl-list').innerHTML = html;
  }

  // ---------- ひな形の選択 ----------
  async function selectTemplate(id) {
    const rec = findRecord(id);
    if (!rec) return;
    let doc;
    try {
      doc = await DocxTemplate.load(rec.bytes);
    } catch (e) {
      toast('ひな形を読み込めませんでした：' + e.message);
      return;
    }
    state.rec = rec;
    state.doc = doc;
    state.fields = buildFields(doc.fields(), rec.fields);
    state.fieldMap = new Map(state.fields.map((f) => [f.key, f]));
    state.dateStyle = store.get('dateStyle.' + rec.id, rec.dateStyle || 'wareki');
    applyDefaults();
    store.set('lastTemplate', rec.id);

    $('#empty-state').hidden = true;
    $('#form-pane').hidden = false;
    $('#actions').hidden = false;
    $('#tpl-title').textContent = rec.name;
    $('#tpl-desc').textContent = rec.description;
    $$('input[name="dateStyle"]').forEach((r) => (r.checked = r.value === state.dateStyle));
    renderTemplateList();
    renderForm();
    update();
    document.body.classList.remove('show-sidebar');
  }

  function applyDefaults() {
    const mem = store.get('remember', {});
    for (const f of state.fields) {
      if (!isEmpty(state.values[f.key])) continue;
      if (f.remember && !isEmpty(mem[f.key])) state.values[f.key] = mem[f.key];
      else if (f.default === 'today' && f.type === 'date') state.values[f.key] = todayIso();
      else if (f.default !== undefined && f.default !== 'today') state.values[f.key] = f.default;
    }
    store.set('case', state.values);
  }

  function showEmpty() {
    state.rec = null;
    state.doc = null;
    $('#form-pane').hidden = true;
    $('#actions').hidden = true;
    $('#empty-state').hidden = false;
    $('#paper').innerHTML = '';
    renderTemplateList();
  }

  // ---------- フォーム ----------
  const fieldId = (key) => 'f-' + Array.from(key).map((c) => c.charCodeAt(0).toString(36)).join('-');

  function renderForm() {
    const form = $('#form');
    form.innerHTML = '';
    if (!state.fields.length) {
      form.innerHTML = '<p class="muted">このひな形にはまだ入力項目がありません。「ひな形を編集」から設定してください。</p>';
      return;
    }
    const groups = new Map();
    for (const f of state.fields) {
      if (!groups.has(f.group)) groups.set(f.group, []);
      groups.get(f.group).push(f);
    }
    for (const [group, fields] of groups) {
      const fs = document.createElement('fieldset');
      fs.innerHTML = `<legend>${esc(group)}</legend>`;
      for (const f of fields) fs.appendChild(buildField(f));
      form.appendChild(fs);
    }
  }

  function buildField(f) {
    const id = fieldId(f.key);
    const v = state.values[f.key];
    const wrap = document.createElement('div');
    wrap.className = 'field field-' + f.type;
    wrap.dataset.key = f.key;
    const badges =
      (f.required ? '' : '<span class="badge">任意</span>') +
      (f.remember ? '<span class="badge badge-mem" title="入力内容を記憶して、次回から自動で入力します">記憶</span>' : '');
    const ph = esc(f.placeholder);
    const val = esc(v == null ? '' : v);
    let control;
    if (f.compute) {
      wrap.classList.add('field-computed');
      wrap.innerHTML =
        `<label for="${id}" class="field-label">${esc(f.label)}<span class="badge badge-calc">自動計算</span></label>` +
        `<input type="text" id="${id}" readonly tabindex="-1">` +
        (f.hint ? `<div class="hint">${esc(f.hint)}</div>` : '');
      return wrap;
    }
    switch (f.type) {
      case 'textarea':
        control = `<textarea id="${id}" rows="3" placeholder="${ph}">${val}</textarea>`;
        break;
      case 'date':
        control =
          `<div class="input-row"><input type="date" id="${id}" value="${val}">` +
          `<button type="button" class="mini" data-today="${esc(f.key)}">今日</button></div>`;
        break;
      case 'money':
        control =
          `<div class="input-row"><input type="text" inputmode="numeric" id="${id}" value="${esc(formatMoney(v))}" placeholder="${ph || '例：25000000'}">` +
          `<span class="suffix">円</span></div>`;
        break;
      case 'number':
        control = `<input type="text" inputmode="decimal" id="${id}" value="${val}" placeholder="${ph}">`;
        break;
      case 'select':
        control =
          `<select id="${id}"><option value="">選択してください</option>` +
          f.options.map((o) => `<option${o === v ? ' selected' : ''}>${esc(o)}</option>`).join('') +
          `</select>`;
        break;
      default:
        control = `<input type="text" id="${id}" value="${val}" placeholder="${ph}">`;
    }
    wrap.innerHTML =
      `<label for="${id}" class="field-label">${esc(f.label)}${badges}</label>` +
      control +
      `<div class="out-preview" aria-live="polite"></div>` +
      (f.hint ? `<div class="hint">${esc(f.hint)}</div>` : '');
    return wrap;
  }

  function onFieldInput(e) {
    const wrap = e.target.closest('.field');
    if (!wrap || !state.rec) return;
    const f = state.fieldMap.get(wrap.dataset.key);
    if (!f || f.compute) return;
    setValue(f, f.type === 'money' ? digitsOnly(e.target.value) : e.target.value);
  }

  function setValue(f, v) {
    state.values[f.key] = v;
    store.set('case', state.values);
    if (f.remember) {
      const mem = store.get('remember', {});
      mem[f.key] = v;
      store.set('remember', mem);
    }
    scheduleUpdate();
  }

  // ---------- プレビュー ----------
  function valueOf(key) {
    const f = state.fieldMap.get(key) || makeField(key);
    const raw = rawValue(f, state.values);
    const text = formatValue(f, raw, state.dateStyle);
    return { text, empty: isEmpty(raw) && f.required, label: f.markerLabel || f.label };
  }

  function scheduleUpdate() {
    clearTimeout(scheduleUpdate.t);
    scheduleUpdate.t = setTimeout(update, 60);
  }

  function update() {
    if (!state.doc) return;
    const scroll = $('.preview-pane').scrollTop;
    $('#paper').innerHTML = state.doc.renderHtml({ mode: 'fill', value: valueOf, dropKeys: dropKeysOf(state.fields) }).html;
    fitPage($('#paper'));
    $('.preview-pane').scrollTop = scroll;

    let total = 0;
    let done = 0;
    for (const f of state.fields) {
      const wrap = $(`.field[data-key="${CSS.escape(f.key)}"]`);
      const raw = rawValue(f, state.values);
      const text = formatValue(f, raw, state.dateStyle);
      if (wrap && f.compute) {
        const input = $('input', wrap);
        input.value = isEmpty(raw) ? '' : text;
      }
      if (wrap) {
        // 実際に差し込まれる文字をフォームにも表示（日付・金額など整形されるもの）
        const out = $('.out-preview', wrap) || { textContent: '' };
        const shown = f.type === 'date' || f.zenkaku || f.type === 'money';
        out.textContent = shown && !f.compute && !isEmpty(raw) ? `差し込まれる文字：${text}` : '';
      }
      if (!f.required) continue;
      total++;
      const filled = !isEmpty(state.values[f.key]);
      if (filled) done++;
      if (wrap) wrap.classList.toggle('missing', !filled);
    }
    const pct = total ? Math.round((done / total) * 100) : 100;
    $('#progress-bar').style.width = pct + '%';
    $('#progress-text').textContent = total === done ? `すべて入力済み（${total}項目）` : `${done} / ${total} 項目入力済み`;
    $('#progress').classList.toggle('complete', total === done);
    highlightActive();
  }

  /** 画面が狭いときは、プレビューの用紙を幅に合わせて縮小表示する */
  function fitPage(container) {
    const page = container.querySelector('.dx-page');
    if (!page) return;
    page.style.zoom = '';
    const avail = container.clientWidth;
    if (avail && page.offsetWidth > avail) page.style.zoom = (avail / page.offsetWidth).toFixed(3);
  }

  function missingLabels() {
    return state.fields.filter((f) => f.required && isEmpty(rawValue(f, state.values))).map((f) => f.markerLabel);
  }

  function highlightActive() {
    const active = document.activeElement && document.activeElement.closest && document.activeElement.closest('.field');
    const key = active ? active.dataset.key : null;
    $$('#paper .ph').forEach((s) => s.classList.toggle('active', s.dataset.key === key));
    return key;
  }

  function onFieldFocus() {
    const key = highlightActive();
    if (!key) return;
    const target = $(`#paper .ph[data-key="${CSS.escape(key)}"]`);
    if (!target) return;
    const pane = $('.preview-pane');
    const pr = pane.getBoundingClientRect();
    const tr = target.getBoundingClientRect();
    if (tr.top < pr.top + 60 || tr.bottom > pr.bottom - 40) {
      pane.scrollTo({ top: pane.scrollTop + tr.top - pr.top - pr.height / 3, behavior: 'smooth' });
    }
  }

  function focusField(key) {
    const wrap = $(`.field[data-key="${CSS.escape(key)}"]`);
    if (!wrap) return;
    document.body.classList.remove('show-preview');
    wrap.scrollIntoView({ behavior: 'smooth', block: 'center' });
    $('input, textarea, select', wrap).focus({ preventScroll: true });
    wrap.classList.add('flash');
    setTimeout(() => wrap.classList.remove('flash'), 900);
  }

  // =====================================================================
  // Word 出力
  // =====================================================================
  function currentFileName() {
    let name = '';
    if (state.rec.fileName) {
      name = state.rec.fileName.replace(/\{\{\s*([^{}:]+?)\s*(?::[a-z]+)?\s*\}\}/g, (all, key) => valueOf(key.trim()).text);
    }
    if (!name.replace(/[_\s-]/g, '')) name = `${state.rec.name}_${todayIso().replace(/-/g, '')}`;
    return name.replace(/[\\/:*?"<>|]/g, '_').replace(/_+$/, '').trim();
  }

  function downloadBlob(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }

  /** 差し込み済みの Word を作り、差し込み箇所以外の文字が変わっていないか検証する */
  async function buildFilledDocx(rec, values, dateStyle) {
    const src = await DocxTemplate.load(rec.bytes);
    const fields = new Map(buildFields(src.fields(), rec.fields).map((f) => [f.key, f]));
    const text = (key) => {
      const f = fields.get(key) || makeField(key);
      return formatValue(f, rawValue(f, values), dateStyle);
    };
    const dropKeys = dropKeysOf(Array.from(fields.values()));

    // 差し込み後に残るべき文字（差し込み箇所以外は元のまま）を先に計算しておく
    const expected = src.expectedTexts(text, dropKeys);

    src.fill(text, dropKeys);
    const blob = await src.toBlob();

    const check = await DocxTemplate.load(await blob.arrayBuffer());
    const actual = check.texts();
    const same = actual.length === expected.length && actual.every((t, i) => t === expected[i]);
    if (!same) throw new Error('出力内容の検証に失敗しました（差し込み以外の文字が変わる可能性があるため中止しました）');
    return blob;
  }

  async function exportDocx() {
    if (!state.rec) return;
    const miss = missingLabels();
    if (
      miss.length &&
      !confirm(
        `未入力の項目が ${miss.length} 件あります。\n\n・${miss.slice(0, 8).join('\n・')}${miss.length > 8 ? '\n…ほか' : ''}\n\n未入力の箇所は空欄（または設定した文字）で出力されます。よろしいですか？`
      )
    )
      return;
    try {
      const blob = await buildFilledDocx(state.rec, state.values, state.dateStyle);
      downloadBlob(blob, currentFileName() + '.docx');
      addHistory();
      toast('Wordファイルを保存しました');
    } catch (e) {
      alert(e.message);
    }
  }

  // =====================================================================
  // 履歴・案件
  // =====================================================================
  function addHistory() {
    const list = store.get('history', []);
    const entry = {
      id: Date.now().toString(36),
      templateId: state.rec.id,
      templateName: state.rec.name,
      title: currentFileName(),
      at: new Date().toISOString(),
      dateStyle: state.dateStyle,
      values: { ...state.values },
    };
    if (list[0] && list[0].templateId === entry.templateId && JSON.stringify(list[0].values) === JSON.stringify(entry.values)) list.shift();
    list.unshift(entry);
    store.set('history', list.slice(0, 200));
  }

  function openHistory() {
    const list = store.get('history', []);
    const box = $('#history-list');
    box.innerHTML = list.length
      ? list
          .map((h) => {
            const d = new Date(h.at);
            const when = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
            return (
              `<div class="hist-item"><div class="hist-main"><div class="hist-title">${esc(h.title)}</div>` +
              `<div class="hist-meta">${esc(h.templateName)} ・ ${when}</div></div>` +
              `<button type="button" class="btn small" data-hist-open="${h.id}">この内容で開く</button>` +
              `<button type="button" class="btn small ghost" data-hist-del="${h.id}">削除</button></div>`
            );
          })
          .join('')
      : '<p class="muted">まだ履歴はありません。Word出力をすると自動で記録されます。</p>';
    $('#history-dialog').showModal();
  }

  function onHistoryClick(e) {
    const openId = e.target.dataset.histOpen;
    const delId = e.target.dataset.histDel;
    const list = store.get('history', []);
    if (openId) {
      const h = list.find((x) => x.id === openId);
      if (!h) return;
      if (!findRecord(h.templateId)) return toast('このひな形は削除されています');
      state.values = { ...h.values };
      store.set('case', state.values);
      store.set('dateStyle.' + h.templateId, h.dateStyle);
      $('#history-dialog').close();
      selectTemplate(h.templateId).then(() => toast('履歴の内容を読み込みました'));
    } else if (delId) {
      store.set('history', list.filter((x) => x.id !== delId));
      openHistory();
    }
  }

  function newCase() {
    if (!confirm('入力内容をクリアして新しい案件を始めますか？\n（自社情報など「記憶」の項目は残ります。これまでの内容は履歴から呼び出せます）')) return;
    state.values = {};
    store.set('case', {});
    if (state.rec) selectTemplate(state.rec.id);
  }

  // =====================================================================
  // 案件データ（お客様ごとの入力内容をまとめて読み書き）
  // =====================================================================
  function parseCase(text) {
    const data = JSON.parse(text);
    const values = data && typeof data.values === 'object' ? data.values : data;
    if (!values || typeof values !== 'object' || Array.isArray(values)) throw new Error('形式が正しくありません');
    const clean = {};
    for (const [k, v] of Object.entries(values)) if (v != null && typeof v !== 'object') clean[k] = String(v);
    return { name: data.name || '', dateStyle: data.dateStyle, values: clean };
  }

  function previewCase() {
    const box = $('#case-summary');
    const text = $('#case-text').value.trim();
    if (!text) return (box.innerHTML = '');
    try {
      const c = parseCase(text);
      const label = (k) => (PRESET_MAP.get(k) || state.fieldMap.get(k) || {}).label || k;
      box.innerHTML =
        `<p><b>${esc(c.name || '案件')}</b>：${Object.keys(c.values).length} 項目</p><table>` +
        Object.entries(c.values)
          .map(([k, v]) => `<tr><td>${esc(label(k))}</td><td>${esc(v)}</td></tr>`)
          .join('') +
        '</table>';
    } catch (e) {
      box.innerHTML = `<p class="err">読み込めません：${esc(e.message)}</p>`;
    }
  }

  function openCaseDialog() {
    $('#case-text').value = '';
    $('#case-summary').innerHTML = '';
    $('#case-dialog').showModal();
  }

  async function loadCase() {
    let c;
    try {
      c = parseCase($('#case-text').value.trim());
    } catch (e) {
      return toast('案件データを読み込めませんでした');
    }
    if (Object.keys(state.values).some((k) => !isEmpty(state.values[k])) && !confirm('今の入力内容を、読み込んだ案件データで置き換えますか？\n（今の内容は履歴に残していない場合、元に戻せません）')) return;
    state.values = { ...c.values };
    store.set('case', state.values);
    if (c.dateStyle === 'wareki' || c.dateStyle === 'seireki') {
      for (const r of allRecords()) store.set('dateStyle.' + r.id, c.dateStyle);
    }
    $('#case-dialog').close();
    if (state.rec) await selectTemplate(state.rec.id);
    toast(`${c.name || '案件データ'}を読み込みました`);
  }

  function exportCase() {
    const name = (state.values['発注者_姓'] ? state.values['発注者_姓'] + '様' : '案件') + '_' + todayIso().replace(/-/g, '');
    const data = { format: 'contract-maker-case', version: 1, name, dateStyle: state.dateStyle, values: state.values };
    downloadBlob(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), name + '.json');
  }

  // =====================================================================
  // ひな形エディタ
  // =====================================================================
  const ed = { id: null, builtinSource: false, doc: null, fields: new Map(), undo: [], sel: null };

  async function openEditorFromFile(file) {
    if (!/\.docx$/i.test(file.name)) {
      toast('.docx 形式の Word ファイルを選んでください（.doc の場合は Word で .docx に保存し直してください）');
      return;
    }
    try {
      const buf = await file.arrayBuffer();
      const doc = await DocxTemplate.load(buf);
      openEditor({ doc, name: file.name.replace(/\.docx$/i, '') });
    } catch (e) {
      toast('Wordファイルを読み込めませんでした：' + e.message);
    }
  }

  async function openEditorFromRecord(rec) {
    const doc = await DocxTemplate.load(rec.bytes);
    openEditor({ doc, rec });
  }

  function openEditor({ doc, rec, name }) {
    ed.id = rec && !rec.builtin ? rec.id : null;
    ed.builtinSource = !!(rec && rec.builtin);
    ed.doc = doc;
    ed.undo = [];
    ed.fields = new Map(buildFields(doc.fields(), rec ? rec.fields : []).map((f) => [f.key, f]));
    $('#te-heading').textContent = !rec ? 'Wordからひな形を作る' : rec.builtin ? '配布ひな形をコピーして編集' : 'ひな形の編集';
    $('#te-note').hidden = !ed.builtinSource;
    $('#te-name').value = rec ? (rec.builtin ? rec.name + '（コピー）' : rec.name) : name || '';
    $('#te-desc').value = rec ? rec.description : '';
    $('#te-filename').value = rec ? rec.fileName : '';
    $('#te-delete').hidden = !ed.id;
    renderEditor();
    $('#tpl-editor').showModal();
  }

  function renderEditor() {
    const scroll = $('.te-doc-wrap').scrollTop;
    $('#te-doc').innerHTML = ed.doc.renderHtml({ mode: 'raw' }).html;
    fitPage($('#te-doc'));
    $('.te-doc-wrap').scrollTop = scroll;
    $('#te-undo').disabled = !ed.undo.length;
    $('#te-pop').hidden = true;
    renderEditorFields();
  }

  function renderEditorFields() {
    const detected = ed.doc.fields();
    $('#te-count').textContent = detected.length ? `${detected.length} 項目` : '';
    const box = $('#te-fields');
    if (!detected.length) {
      box.innerHTML =
        '<p class="muted">左の文書で、毎回変わる文字（お客様名・金額・日付など）をドラッグで選択してください。<br><br>' +
        'Word 上であらかじめ <code>{{注文者_氏名}}</code> のように書いておいた箇所も自動で入力項目になります。</p>';
      return;
    }
    const groups = new Set([DEFAULT_GROUP, ...PRESETS.map((p) => p.group)]);
    ed.fields.forEach((f) => groups.add(f.group));
    $('#te-groups').innerHTML = Array.from(groups).map((g) => `<option value="${esc(g)}">`).join('');

    box.innerHTML = detected
      .map((d) => {
        let f = ed.fields.get(d.key);
        if (!f) ed.fields.set(d.key, (f = makeField(d.key, d.type)));
        const k = esc(d.key);
        const numeric = NUMERIC_TYPES.includes(f.type);
        return (
          `<div class="te-field" data-key="${k}">` +
          `<div class="te-field-head"><code>{{${k}}}</code>` +
          (f.sample ? `<span class="te-sample" title="元の文字">元：${esc(f.sample)}</span>` : '') +
          `<button type="button" class="link-btn te-unmark" data-unmark="${k}">解除</button></div>` +
          `<div class="te-grid">` +
          `<label>表示名<input data-prop="label" value="${esc(f.label)}"></label>` +
          `<label>種類<select data-prop="type">${Object.entries(TYPE_LABELS)
            .map(([v, l]) => `<option value="${v}"${v === f.type ? ' selected' : ''}>${l}</option>`)
            .join('')}</select></label>` +
          `<label>グループ<input data-prop="group" list="te-groups" value="${esc(f.group)}"></label>` +
          `<label class="te-wide"${f.type === 'select' ? '' : ' hidden'}>選択肢（カンマ区切り）<input data-prop="options" value="${esc(f.options.join(','))}"></label>` +
          `<label class="te-wide">補足説明<input data-prop="hint" value="${esc(f.hint)}" placeholder="入力欄の下に表示されます"></label>` +
          `<label class="te-wide">未入力のときに入れる文字<input data-prop="blank" value="${esc(f.blank)}" placeholder="空欄のまま（例：全角スペース、―）"></label>` +
          `</div><div class="te-checks">` +
          `<label><input type="checkbox" data-prop="required"${f.required ? ' checked' : ''}> 必須</label>` +
          `<label><input type="checkbox" data-prop="remember"${f.remember ? ' checked' : ''}> 入力を記憶（自社情報など）</label>` +
          (numeric ? `<label><input type="checkbox" data-prop="zenkaku"${f.zenkaku ? ' checked' : ''}> 数字を全角にする</label>` : '') +
          (f.type === 'date' ? `<label><input type="checkbox" data-prop="today"${f.default === 'today' ? ' checked' : ''}> 初期値を今日にする</label>` : '') +
          `<label title="例：「( ポーチ・バルコニー ○○㎡ )」の行を、面積が無いときは行ごと消す"><input type="checkbox" data-prop="dropEmptyPara"${f.dropEmptyPara ? ' checked' : ''}> 未入力のとき行ごと削除</label>` +
          (f.compute ? '<span class="badge badge-calc">自動計算</span>' : '') +
          `</div></div>`
        );
      })
      .join('');
  }

  function onEditorFieldChange(e) {
    const el = e.target;
    const prop = el.dataset.prop;
    const card = el.closest('.te-field');
    if (!prop || !card) return;
    const f = ed.fields.get(card.dataset.key);
    if (!f) return;
    if (prop === 'required' || prop === 'remember' || prop === 'zenkaku' || prop === 'dropEmptyPara') f[prop] = el.checked;
    else if (prop === 'today') f.default = el.checked ? 'today' : undefined;
    else if (prop === 'options') f.options = el.value.split(/[,、]/).map((s) => s.trim()).filter(Boolean);
    else f[prop] = el.value;
    if (prop === 'type') renderEditorFields();
  }

  // ---------- 文字の選択 → 入力項目にする ----------
  function offsetInPara(pEl, node, offset) {
    const r = document.createRange();
    r.setStart(pEl, 0);
    r.setEnd(node, offset);
    let n = 0;
    r.cloneContents()
      .querySelectorAll('[data-t]')
      .forEach((s) => (n += s.textContent.length));
    return n;
  }

  function readSelection() {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return null;
    const range = sel.getRangeAt(0);
    const root = $('#te-doc');
    if (!root.contains(range.commonAncestorContainer)) return null;
    const elOf = (n) => (n.nodeType === 1 ? n : n.parentElement);
    const p1 = elOf(range.startContainer).closest('[data-pid]');
    const p2 = elOf(range.endContainer).closest('[data-pid]');
    if (!p1 || p1 !== p2) return { error: '1つの段落（行のまとまり）の中で選択してください' };
    const s = offsetInPara(p1, range.startContainer, range.startOffset);
    const e = offsetInPara(p1, range.endContainer, range.endOffset);
    if (e <= s) return null;
    const p = ed.doc.paragraphById(+p1.dataset.pid);
    const text = DocxTemplate.paraText(p);
    const phs = DocxTemplate.findPlaceholders(text);
    if (phs.some((ph) => s < ph.end && e > ph.start)) return { error: 'すでに入力項目になっている部分が含まれています' };
    return { p, s, e, text: text.slice(s, e), rect: range.getBoundingClientRect() };
  }

  function onDocMouseUp() {
    setTimeout(() => {
      const sel = readSelection();
      const pop = $('#te-pop');
      if (!sel || sel.error) {
        pop.hidden = true;
        ed.sel = null;
        if (sel && sel.error) toast(sel.error);
        return;
      }
      ed.sel = sel;
      const wrap = $('.te-doc-wrap').getBoundingClientRect();
      pop.style.top = sel.rect.bottom - wrap.top + $('.te-doc-wrap').scrollTop + 6 + 'px';
      pop.style.left = Math.max(8, sel.rect.left - wrap.left) + 'px';
      pop.hidden = false;
    }, 0);
  }

  function openMarkDialog() {
    const sel = ed.sel;
    if (!sel) return;
    $('#mark-text').textContent = sel.text;
    const count = ed.doc.countText(sel.text);
    $('#mark-all-count').textContent = count;
    $('#mark-all-wrap').hidden = count < 2;
    $('#mark-all').checked = true;
    const used = new Set(ed.doc.fields().map((f) => f.key));
    // 既に同じ元の文字で作った項目があれば、それを初期値にする
    const same = Array.from(ed.fields.values()).find((f) => f.sample === sel.text && used.has(f.key));
    $('#mark-key').value = same ? same.key : '';
    const keys = new Set([...used, ...PRESETS.map((p) => p.key)]);
    $('#mark-keys').innerHTML = Array.from(keys)
      .map((k) => `<option value="${esc(k)}">${esc((PRESET_MAP.get(k) || ed.fields.get(k) || {}).label || '')}</option>`)
      .join('');
    const groups = new Map();
    if (used.size) groups.set('このひな形の項目', Array.from(used).map((k) => ({ key: k, label: (ed.fields.get(k) || {}).label || k })));
    for (const p of PRESETS) {
      if (!groups.has(p.group)) groups.set(p.group, []);
      groups.get(p.group).push(p);
    }
    $('#mark-presets').innerHTML =
      '<p class="muted">よく使う項目（クリックで選択）</p>' +
      Array.from(groups)
        .map(
          ([g, list]) =>
            `<div class="chip-group"><span class="chip-group-name">${esc(g)}</span>` +
            list.map((p) => `<button type="button" class="chip" data-key="${esc(p.key)}" title="${esc(p.key)}">${esc(p.label)}</button>`).join('') +
            `</div>`
        )
        .join('');
    $('#te-pop').hidden = true;
    $('#mark-dialog').showModal();
    $('#mark-key').focus();
  }

  function confirmMark() {
    const sel = ed.sel;
    const key = $('#mark-key').value.trim().replace(/[{}:]/g, '');
    if (!sel) return;
    if (!key) {
      $('#mark-key').focus();
      return toast('項目名を入力してください');
    }
    ed.undo.push(ed.doc.snapshot());
    if ($('#mark-all').checked && !$('#mark-all-wrap').hidden) ed.doc.markAll(sel.text, key);
    else ed.doc.markRange(sel.p, sel.s, sel.e, key);
    if (!ed.fields.has(key)) ed.fields.set(key, makeField(key, '', { sample: sel.text }));
    else if (!ed.fields.get(key).sample) ed.fields.get(key).sample = sel.text;
    ed.sel = null;
    window.getSelection().removeAllRanges();
    $('#mark-dialog').close();
    renderEditor();
    const card = $(`.te-field[data-key="${CSS.escape(key)}"]`);
    if (card) {
      card.scrollIntoView({ block: 'nearest' });
      card.classList.add('flash');
      setTimeout(() => card.classList.remove('flash'), 900);
    }
  }

  function unmark(key) {
    const f = ed.fields.get(key);
    const back = f && f.sample ? f.sample : '';
    if (!back && !confirm(`{{${key}}} の元の文字が分からないため、この箇所は空欄になります。解除しますか？`)) return;
    ed.undo.push(ed.doc.snapshot());
    ed.doc.fill((k) => (k === key ? back : null));
    renderEditor();
  }

  function undo() {
    const snap = ed.undo.pop();
    if (!snap) return;
    ed.doc.restore(snap);
    renderEditor();
  }

  async function editorRecord() {
    const detected = ed.doc.fields();
    const fields = buildFields(detected, ed.fields).map(serializeField);
    return {
      id: ed.id || 'custom-' + Date.now().toString(36),
      name: $('#te-name').value.trim() || '無題のひな形',
      description: $('#te-desc').value.trim(),
      fileName: $('#te-filename').value.trim(),
      dateStyle: 'wareki',
      fields,
      bytes: await ed.doc.toArrayBuffer(),
      updatedAt: new Date().toISOString(),
    };
  }

  async function saveEditor() {
    if (!$('#te-name').value.trim()) {
      $('#te-name').focus();
      return toast('ひな形の名前を入力してください');
    }
    const rec = await editorRecord();
    try {
      await idb.put(rec);
    } catch (e) {
      return alert('保存できませんでした：' + e.message);
    }
    await reloadCustoms();
    $('#tpl-editor').close();
    await selectTemplate(rec.id);
    toast('ひな形を保存しました');
  }

  // ---------- 書き出し・読み込み ----------
  function exportJson(recs, name) {
    downloadBlob(new Blob([JSON.stringify(bundleOf(recs))], { type: 'application/json' }), name + '.json');
  }

  function exportJs(recs, name) {
    const js =
      '/* 契約書メーカー 配布ひな形（このファイルを templates/ に置き、index.html に <script> を1行追加） */\n' +
      'ContractApp.registerBundle(' +
      JSON.stringify(bundleOf(recs)) +
      ');\n';
    downloadBlob(new Blob([js], { type: 'text/javascript' }), name + '.js');
  }

  async function importFile(file) {
    if (/\.docx$/i.test(file.name)) return openEditorFromFile(file);
    try {
      let text = await file.text();
      text = text.replace(/^\s*\/\*[\s\S]*?\*\/\s*/, '').replace(/^\s*ContractApp\.registerBundle\(/, '').replace(/\);?\s*$/, '');
      const data = JSON.parse(text);
      let n = 0;
      for (const t of data.templates || []) {
        if (!t || !t.docx) continue;
        const rec = recordFromBundle(t, false);
        if (builtinRecords.some((b) => b.id === rec.id)) rec.id = 'custom-' + Date.now().toString(36) + n;
        await DocxTemplate.load(rec.bytes); // 壊れていないか確認
        await idb.put({ ...rec, builtin: undefined, updatedAt: new Date().toISOString() });
        n++;
      }
      await reloadCustoms();
      renderTemplateList();
      toast(n ? `${n} 件のひな形を読み込みました` : '読み込めるひな形がありませんでした');
    } catch (e) {
      toast('ファイルを読み込めませんでした');
    }
  }

  // =====================================================================
  // 初期化
  // =====================================================================
  async function init() {
    await reloadCustoms();
    renderTemplateList();

    $('#tpl-list').addEventListener('click', (e) => {
      const btn = e.target.closest('.tpl-item');
      if (btn) selectTemplate(btn.dataset.id);
    });

    const form = $('#form');
    form.addEventListener('input', onFieldInput);
    form.addEventListener('change', onFieldInput);
    form.addEventListener('focusin', onFieldFocus);
    form.addEventListener('focusout', (e) => {
      setTimeout(highlightActive, 0);
      if (e.target.closest('.field-money')) e.target.value = formatMoney(e.target.value);
    });
    form.addEventListener('submit', (e) => e.preventDefault());
    form.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.isComposing || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'BUTTON') return;
      e.preventDefault();
      const inputs = $$('input, select, textarea', form).filter((el) => el.offsetParent !== null);
      const i = inputs.indexOf(e.target);
      if (i >= 0 && inputs[i + 1]) inputs[i + 1].focus();
    });
    form.addEventListener('click', (e) => {
      const key = e.target.dataset.today;
      if (!key) return;
      const input = $('input', e.target.closest('.field'));
      input.value = todayIso();
      setValue(state.fieldMap.get(key), input.value);
    });

    $('#paper').addEventListener('click', (e) => {
      const ph = e.target.closest('.ph');
      if (ph) focusField(ph.dataset.key);
    });

    $$('input[name="dateStyle"]').forEach((r) =>
      r.addEventListener('change', () => {
        state.dateStyle = r.value;
        store.set('dateStyle.' + state.rec.id, r.value);
        update();
      })
    );

    $('#btn-docx').addEventListener('click', exportDocx);
    $('#btn-history').addEventListener('click', openHistory);
    $('#btn-new-case').addEventListener('click', newCase);
    $('#btn-case').addEventListener('click', openCaseDialog);
    $('#case-text').addEventListener('input', previewCase);
    $('#case-load').addEventListener('click', loadCase);
    $('#case-export').addEventListener('click', exportCase);
    $('#case-file-btn').addEventListener('click', () => $('#case-file').click());
    $('#case-file').addEventListener('change', async (e) => {
      const file = e.target.files[0];
      e.target.value = '';
      if (!file) return;
      $('#case-text').value = await file.text();
      previewCase();
    });
    $('#history-list').addEventListener('click', onHistoryClick);
    $('#btn-toggle-preview').addEventListener('click', () => document.body.classList.toggle('show-preview'));
    $('#btn-menu').addEventListener('click', () => document.body.classList.toggle('show-sidebar'));

    $('#btn-new-tpl').addEventListener('click', () => $('#file-docx').click());
    $('#file-docx').addEventListener('change', (e) => {
      if (e.target.files[0]) openEditorFromFile(e.target.files[0]);
      e.target.value = '';
    });
    $('#btn-edit-tpl').addEventListener('click', () => openEditorFromRecord(state.rec));
    $('#btn-import').addEventListener('click', () => $('#file-import').click());
    $('#file-import').addEventListener('change', (e) => {
      if (e.target.files[0]) importFile(e.target.files[0]);
      e.target.value = '';
    });
    $('#btn-export-all').addEventListener('click', () => {
      if (!customRecords.length) return toast('書き出すマイひな形がありません');
      exportJson(customRecords, 'マイひな形_' + todayIso().replace(/-/g, ''));
    });

    // エディタ
    $('#te-doc').addEventListener('mouseup', onDocMouseUp);
    $('#te-doc').addEventListener('keyup', onDocMouseUp);
    $('#te-doc').addEventListener('click', (e) => {
      const tok = e.target.closest('.tok');
      if (!tok || !window.getSelection().isCollapsed) return;
      const card = $(`.te-field[data-key="${CSS.escape(tok.dataset.key)}"]`);
      if (card) {
        card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        card.classList.add('flash');
        setTimeout(() => card.classList.remove('flash'), 900);
      }
    });
    $('#te-pop').addEventListener('mousedown', (e) => e.preventDefault());
    $('#te-pop').addEventListener('click', openMarkDialog);
    $('#te-undo').addEventListener('click', undo);
    $('#te-fields').addEventListener('input', onEditorFieldChange);
    $('#te-fields').addEventListener('change', onEditorFieldChange);
    $('#te-fields').addEventListener('click', (e) => {
      if (e.target.dataset.unmark) unmark(e.target.dataset.unmark);
    });
    $('#te-save').addEventListener('click', saveEditor);
    $('#te-delete').addEventListener('click', async () => {
      if (!ed.id || !confirm('このひな形を削除しますか？（元に戻せません）')) return;
      await idb.del(ed.id);
      await reloadCustoms();
      $('#tpl-editor').close();
      if (state.rec && state.rec.id === ed.id) showEmpty();
      else renderTemplateList();
      toast('ひな形を削除しました');
    });
    $('#te-export').addEventListener('click', async () => {
      const rec = await editorRecord();
      exportJson([rec], rec.name);
    });
    $('#te-export-js').addEventListener('click', async () => {
      const rec = await editorRecord();
      exportJs([rec], rec.name);
    });

    $('#mark-ok').addEventListener('click', confirmMark);
    $('#mark-key').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.isComposing) {
        e.preventDefault();
        confirmMark();
      }
    });
    $('#mark-presets').addEventListener('click', (e) => {
      const chip = e.target.closest('.chip');
      if (!chip) return;
      $('#mark-key').value = chip.dataset.key;
      $$('.chip', $('#mark-presets')).forEach((c) => c.classList.toggle('selected', c === chip));
    });
    $$('[data-close]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close()));
    window.addEventListener('resize', () => {
      fitPage($('#paper'));
      if ($('#tpl-editor').open) fitPage($('#te-doc'));
    });
    $('#btn-toggle-preview').addEventListener('click', () => setTimeout(() => fitPage($('#paper')), 0));

    const last = store.get('lastTemplate', null);
    if (last && findRecord(last)) await selectTemplate(last);
    else if (allRecords().length) await selectTemplate(allRecords()[0].id);
  }

  window.ContractApp = {
    registerBundle,
    _internal: { formatDate, formatMoney, toZenkaku, buildFilledDocx, makeField, buildFields, rawValue, formatValue, parseCase },
  };

  document.addEventListener('DOMContentLoaded', init);
})();
