/*
 * 契約書メーカー 本体
 *
 * ひな形の本文に {{項目名}} を書くと、入力フォームが自動で作られ、
 * 入力した内容がリアルタイムに契約書へ差し込まれる。
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
    checkbox: 'ON / OFF',
  };
  const DEFAULT_GROUP = '入力項目';
  const RESERVED = new Set(['else', '条']);

  // {{項目名}} または {{項目名:型}}
  const PH_RE = /\{\{\s*([^{}#\/:\s][^{}:]*?)\s*(?::\s*([a-z]+)\s*)?\}\}/g;
  // {{#if 項目}}...{{else}}...{{/if}} / {{#unless 項目}}...{{/unless}}
  const COND_RE = /\{\{#(if|unless)\s+([^{}]+?)\s*\}\}([\s\S]*?)\{\{\/\1\s*\}\}/g;
  const COND_KEY_RE = /\{\{#(?:if|unless)\s+([^{}]+?)\s*\}\}/g;
  const ARTICLE_RE = /\{\{\s*条\s*\}\}/g;

  // =====================================================================
  // 保存 (ブラウザの localStorage)
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
        /* 保存できなくても動作は続ける */
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

  // =====================================================================
  // ひな形の登録・正規化
  // =====================================================================
  const builtinTemplates = [];

  function detectFields(body) {
    const found = [];
    const seen = new Set();
    const add = (key, type, index) => {
      key = key.trim();
      if (!key || RESERVED.has(key) || seen.has(key)) return;
      seen.add(key);
      found.push({ key, type, index });
    };
    let m;
    const all = [];
    PH_RE.lastIndex = 0;
    while ((m = PH_RE.exec(body))) all.push({ key: m[1], type: m[2] || '', index: m.index });
    COND_KEY_RE.lastIndex = 0;
    while ((m = COND_KEY_RE.exec(body))) all.push({ key: m[1], type: 'checkbox', index: m.index });
    all.sort((a, b) => a.index - b.index);
    // 同じ項目に型指定付きの出現があればそれを優先
    const typed = {};
    for (const a of all) if (a.type && !typed[a.key.trim()]) typed[a.key.trim()] = a.type;
    for (const a of all) add(a.key, typed[a.key.trim()] || 'text', a.index);
    return found;
  }

  function normalizeField(f, detectedType) {
    let type = f.type || detectedType || 'text';
    if (!TYPE_LABELS[type]) type = 'text';
    const options = Array.isArray(f.options)
      ? f.options
      : typeof f.options === 'string'
        ? f.options.split(/[,、\n]/).map((s) => s.trim()).filter(Boolean)
        : [];
    return {
      key: f.key,
      label: f.label || f.key,
      type,
      group: f.group || DEFAULT_GROUP,
      required: f.required !== undefined ? !!f.required : type !== 'checkbox',
      remember: !!f.remember,
      options,
      placeholder: f.placeholder || '',
      hint: f.hint || '',
      default: f.default,
    };
  }

  function normalizeTemplate(t, builtin) {
    const body = String(t.body || '').replace(/\r\n?/g, '\n');
    const detected = detectFields(body);
    const detectedType = Object.fromEntries(detected.map((d) => [d.key, d.type]));
    const fields = [];
    const map = new Map();
    for (const f of t.fields || []) {
      if (!f || !f.key || map.has(f.key)) continue;
      const nf = normalizeField(f, detectedType[f.key]);
      fields.push(nf);
      map.set(nf.key, nf);
    }
    for (const d of detected) {
      if (map.has(d.key)) continue;
      const nf = normalizeField({ key: d.key, group: t.defaultGroup }, d.type);
      fields.push(nf);
      map.set(nf.key, nf);
    }
    // 同じ表示名が複数あるとき（甲・乙の「会社名」など）は、グループ名を付けて区別する
    const labelCount = {};
    for (const f of fields) labelCount[f.label] = (labelCount[f.label] || 0) + 1;
    for (const f of fields) f.markerLabel = labelCount[f.label] > 1 ? `${f.group}・${f.label}` : f.label;
    return {
      id: String(t.id || 'tpl-' + Date.now().toString(36)),
      name: t.name || '無題のひな形',
      description: t.description || '',
      category: t.category || '',
      fileName: t.fileName || '',
      dateStyle: t.dateStyle === 'seireki' ? 'seireki' : 'wareki',
      body,
      fields,
      fieldMap: map,
      builtin: !!builtin,
    };
  }

  /** templates/*.js から呼ばれる */
  function register(t) {
    const nt = normalizeTemplate(t, true);
    const i = builtinTemplates.findIndex((x) => x.id === nt.id);
    if (i >= 0) builtinTemplates[i] = nt;
    else builtinTemplates.push(nt);
  }

  function serializeTemplate(t) {
    return {
      id: t.id,
      name: t.name,
      description: t.description,
      category: t.category,
      fileName: t.fileName,
      dateStyle: t.dateStyle,
      fields: t.fields.map((f) => {
        const o = { key: f.key, label: f.label, type: f.type, group: f.group };
        if (f.required !== (f.type !== 'checkbox')) o.required = f.required;
        if (f.remember) o.remember = true;
        if (f.options.length) o.options = f.options;
        if (f.placeholder) o.placeholder = f.placeholder;
        if (f.hint) o.hint = f.hint;
        if (f.default !== undefined) o.default = f.default;
        return o;
      }),
      body: t.body,
    };
  }

  function customTemplates() {
    return store.get('customTemplates', []).map((t) => normalizeTemplate(t, false));
  }

  function saveCustomTemplate(t) {
    const list = store.get('customTemplates', []);
    const data = serializeTemplate(t);
    const i = list.findIndex((x) => x.id === t.id);
    if (i >= 0) list[i] = data;
    else list.push(data);
    store.set('customTemplates', list);
  }

  function deleteCustomTemplate(id) {
    store.set(
      'customTemplates',
      store.get('customTemplates', []).filter((x) => x.id !== id)
    );
  }

  function allTemplates() {
    return [...builtinTemplates, ...customTemplates()];
  }

  function findTemplate(id) {
    return allTemplates().find((t) => t.id === id) || null;
  }

  // =====================================================================
  // 値の整形
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
      const s = era.start[0] * 10000 + era.start[1] * 100 + era.start[2];
      if (n >= s) {
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

  function isEmpty(v) {
    return v == null || v === '' || v === false;
  }

  function formatValue(f, raw, dateStyle) {
    if (f.type === 'checkbox') return raw ? 'あり' : 'なし';
    if (isEmpty(raw)) return '';
    switch (f.type) {
      case 'date':
        return formatDate(raw, dateStyle);
      case 'money':
        return formatMoney(raw);
      default:
        return String(raw);
    }
  }

  function todayIso() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }

  // =====================================================================
  // 差し込み処理: ひな形 + 入力値 → ブロック配列
  // =====================================================================
  const DROP = '\u0000';

  function applyConditions(body, values) {
    for (let i = 0; i < 5; i++) {
      const next = body.replace(COND_RE, (all, kind, key, inner) => {
        const parts = inner.split(/\{\{\s*else\s*\}\}/);
        const truthy = !isEmpty(values[key.trim()]);
        let pick = (kind === 'if') === truthy ? parts[0] : parts[1] || '';
        pick = pick.replace(/^\n/, '').replace(/\n$/, '');
        return pick === '' ? DROP : pick;
      });
      if (next === body) break;
      body = next;
    }
    // 条件で消えた結果だけが残る行は、行ごと削除する
    return body
      .split('\n')
      .filter((line) => !(line.includes(DROP) && line.split(DROP).join('').trim() === ''))
      .join('\n')
      .split(DROP)
      .join('');
  }

  function evaluate(tpl, values, dateStyle) {
    let body = applyConditions(tpl.body, values);
    let article = 0;
    body = body.replace(ARTICLE_RE, () => String(++article));

    const used = new Set();
    const substitute = (text) => {
      const segs = [];
      let last = 0;
      let m;
      PH_RE.lastIndex = 0;
      while ((m = PH_RE.exec(text))) {
        const key = m[1].trim();
        if (RESERVED.has(key)) continue;
        if (m.index > last) segs.push({ t: 'text', v: text.slice(last, m.index) });
        const f = tpl.fieldMap.get(key) || normalizeField({ key }, m[2]);
        const v = formatValue(f, values[key], dateStyle);
        const empty = v === '' && f.required;
        used.add(key);
        const label = f.markerLabel || f.label;
        segs.push({ t: 'field', key, label, v: empty ? `【${label}】` : v, empty });
        last = m.index + m[0].length;
      }
      if (last < text.length) segs.push({ t: 'text', v: text.slice(last) });
      return segs;
    };

    const blocks = body.split('\n').map((line) => {
      if (/^={3,}\s*$/.test(line)) return { kind: 'pagebreak', segs: [] };
      let kind = 'para';
      let text = line;
      if (line.startsWith('# ')) (kind = 'title'), (text = line.slice(2));
      else if (line.startsWith('## ')) (kind = 'heading'), (text = line.slice(3));
      else if (line.startsWith('>> ')) (kind = 'right'), (text = line.slice(3));
      else if (line.startsWith('> ') || line === '>') (kind = 'indent'), (text = line.slice(2));
      else if (line.startsWith('^ ')) (kind = 'center'), (text = line.slice(2));
      else if (line.trim() === '') kind = 'blank';
      return { kind, segs: substitute(text) };
    });

    // 条件で使われていない項目も「使われている」扱いにする（ON/OFF項目）
    COND_KEY_RE.lastIndex = 0;
    let m;
    while ((m = COND_KEY_RE.exec(tpl.body))) used.add(m[1].trim());

    return { blocks, used };
  }

  function blocksToText(blocks) {
    return blocks
      .map((b) => {
        if (b.kind === 'pagebreak') return '\n';
        const t = b.segs.map((s) => s.v).join('');
        if (b.kind === 'indent' || b.kind === 'right') return '　　　　　　　　　　' + t;
        if (b.kind === 'title' || b.kind === 'center') return '　　　　　　　　' + t;
        return t;
      })
      .join('\n');
  }

  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function blocksToHtml(blocks) {
    return blocks
      .map((b) => {
        if (b.kind === 'pagebreak') return '<div class="pagebreak"></div>';
        if (b.kind === 'blank') return '<div class="ln ln-blank">&nbsp;</div>';
        const inner = b.segs
          .map((s) => {
            if (s.t === 'text') return esc(s.v);
            const cls = s.empty ? 'ph empty' : 'ph filled';
            return `<span class="${cls}" data-key="${esc(s.key)}" title="${esc(s.label)}">${esc(s.v)}</span>`;
          })
          .join('');
        return `<div class="ln ln-${b.kind}">${inner || '&nbsp;'}</div>`;
      })
      .join('');
  }

  function renderFileName(tpl, values, dateStyle) {
    let name = '';
    if (tpl.fileName) {
      name = tpl.fileName.replace(PH_RE, (all, key) => {
        const f = tpl.fieldMap.get(key.trim()) || normalizeField({ key: key.trim() });
        return formatValue(f, values[key.trim()], dateStyle);
      });
    }
    if (!name.replace(/[_\s-]/g, '')) name = `${tpl.name}_${todayIso().replace(/-/g, '')}`;
    return name.replace(/[\\/:*?"<>|]/g, '_').replace(/_+$/, '').trim();
  }

  // =====================================================================
  // 画面
  // =====================================================================
  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

  const state = {
    tpl: null,
    values: {},
    dateStyle: 'wareki',
    lastResult: null,
  };

  function toast(msg) {
    const el = $('#toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => el.classList.remove('show'), 2400);
  }

  // ---------- サイドバー ----------
  function renderTemplateList() {
    const list = $('#tpl-list');
    const builtins = builtinTemplates;
    const customs = customTemplates();
    const item = (t) =>
      `<li><button type="button" class="tpl-item${state.tpl && state.tpl.id === t.id ? ' active' : ''}" data-id="${esc(t.id)}">` +
      `<span class="tpl-name">${esc(t.name)}</span>` +
      (t.description ? `<span class="tpl-desc">${esc(t.description)}</span>` : '') +
      `</button></li>`;
    let html = '';
    if (builtins.length) html += `<li class="tpl-sep">標準ひな形</li>${builtins.map(item).join('')}`;
    html += `<li class="tpl-sep">マイひな形</li>`;
    html += customs.length
      ? customs.map(item).join('')
      : '<li class="tpl-none">まだありません。<br>「＋ 新しいひな形」から追加できます。</li>';
    list.innerHTML = html;
  }

  // ---------- テンプレート選択 ----------
  function selectTemplate(id) {
    const tpl = findTemplate(id);
    if (!tpl) return;
    state.tpl = tpl;
    state.dateStyle = store.get('dateStyle.' + tpl.id, tpl.dateStyle);
    const draft = store.get('draft.' + tpl.id, null);
    const remembered = store.get('remember', {});
    const values = {};
    for (const f of tpl.fields) {
      if (draft && f.key in draft) values[f.key] = draft[f.key];
      else if (f.remember && f.key in remembered) values[f.key] = remembered[f.key];
      else if (f.default === 'today' && f.type === 'date') values[f.key] = todayIso();
      else if (f.default !== undefined) values[f.key] = f.default;
    }
    state.values = values;
    store.set('lastTemplate', tpl.id);

    $('#empty-state').hidden = true;
    $('#form-pane').hidden = false;
    $('#actions').hidden = false;
    $('#tpl-title').textContent = tpl.name;
    $('#tpl-desc').textContent = tpl.description;
    $$('input[name="dateStyle"]').forEach((r) => (r.checked = r.value === state.dateStyle));
    renderTemplateList();
    renderForm();
    update();
    document.body.classList.remove('show-sidebar');
  }

  // ---------- フォーム ----------
  function fieldId(key) {
    return 'f-' + Array.from(key).map((c) => c.charCodeAt(0).toString(36)).join('-');
  }

  function renderForm() {
    const tpl = state.tpl;
    const form = $('#form');
    const groups = new Map();
    for (const f of tpl.fields) {
      if (!groups.has(f.group)) groups.set(f.group, []);
      groups.get(f.group).push(f);
    }
    form.innerHTML = '';
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
      (f.required || f.type === 'checkbox' ? '' : '<span class="badge">任意</span>') +
      (f.remember ? '<span class="badge badge-mem" title="入力内容を記憶して、次回から自動で入力します">記憶</span>' : '');

    if (f.type === 'checkbox') {
      wrap.innerHTML =
        `<label class="switch"><input type="checkbox" id="${id}"${v ? ' checked' : ''}>` +
        `<span class="switch-ui"></span><span class="switch-label">${esc(f.label)}</span>${badges}</label>`;
    } else {
      let control;
      const ph = esc(f.placeholder);
      const val = esc(v == null ? '' : v);
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
            `<div class="input-row"><input type="text" inputmode="numeric" id="${id}" value="${esc(formatMoney(v))}" placeholder="${ph || '例：100000'}">` +
            `<span class="suffix">円</span></div>`;
          break;
        case 'number':
          control = `<input type="text" inputmode="decimal" id="${id}" value="${val}" placeholder="${ph}">`;
          break;
        case 'select':
          control =
            `<select id="${id}"><option value="">選択してください</option>` +
            f.options
              .map((o) => `<option${o === v ? ' selected' : ''}>${esc(o)}</option>`)
              .join('') +
            `</select>`;
          break;
        default:
          control = `<input type="text" id="${id}" value="${val}" placeholder="${ph}">`;
      }
      wrap.innerHTML =
        `<label for="${id}" class="field-label">${esc(f.label)}${badges}</label>` +
        control +
        (f.hint ? `<div class="hint">${esc(f.hint)}</div>` : '');
    }
    return wrap;
  }

  function readInput(el, f) {
    if (f.type === 'checkbox') return el.checked;
    if (f.type === 'money') return digitsOnly(el.value);
    return el.value;
  }

  function onFieldInput(e) {
    const wrap = e.target.closest('.field');
    if (!wrap || !state.tpl) return;
    const f = state.tpl.fieldMap.get(wrap.dataset.key);
    if (!f) return;
    setValue(f, readInput(e.target, f));
  }

  function setValue(f, v) {
    state.values[f.key] = v;
    store.set('draft.' + state.tpl.id, state.values);
    if (f.remember) {
      const mem = store.get('remember', {});
      mem[f.key] = v;
      store.set('remember', mem);
    }
    update();
  }

  // ---------- プレビュー更新 ----------
  function update() {
    const tpl = state.tpl;
    if (!tpl) return;
    const result = evaluate(tpl, state.values, state.dateStyle);
    state.lastResult = result;
    $('#paper').innerHTML = blocksToHtml(result.blocks);

    // 進捗
    let total = 0;
    let done = 0;
    for (const f of tpl.fields) {
      const wrap = $(`.field[data-key="${CSS.escape(f.key)}"]`);
      const used = result.used.has(f.key);
      if (wrap) wrap.classList.toggle('unused', !used);
      if (!used || !f.required || f.type === 'checkbox') continue;
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

  function missingLabels() {
    if (!state.lastResult) return [];
    const labels = [];
    for (const b of state.lastResult.blocks)
      for (const s of b.segs) if (s.t === 'field' && s.empty && !labels.includes(s.label)) labels.push(s.label);
    return labels;
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
    if (tr.top < pr.top + 40 || tr.bottom > pr.bottom - 40) {
      pane.scrollTo({ top: pane.scrollTop + tr.top - pr.top - pr.height / 3, behavior: 'smooth' });
    }
  }

  function focusField(key) {
    const wrap = $(`.field[data-key="${CSS.escape(key)}"]`);
    if (!wrap) return;
    document.body.classList.remove('show-preview');
    const input = $('input, textarea, select', wrap);
    wrap.scrollIntoView({ behavior: 'smooth', block: 'center' });
    input.focus({ preventScroll: true });
    wrap.classList.add('flash');
    setTimeout(() => wrap.classList.remove('flash'), 900);
  }

  // =====================================================================
  // 出力
  // =====================================================================
  function confirmMissing() {
    const miss = missingLabels();
    if (!miss.length) return true;
    return confirm(
      `未入力の項目が ${miss.length} 件あります。\n\n・${miss.slice(0, 8).join('\n・')}${miss.length > 8 ? '\n…ほか' : ''}\n\nこのまま出力しますか？`
    );
  }

  function currentFileName() {
    return renderFileName(state.tpl, state.values, state.dateStyle);
  }

  function addHistory() {
    const list = store.get('history', []);
    const entry = {
      id: Date.now().toString(36),
      templateId: state.tpl.id,
      templateName: state.tpl.name,
      title: currentFileName(),
      at: new Date().toISOString(),
      dateStyle: state.dateStyle,
      values: { ...state.values },
    };
    // 直前と同じ内容なら上書き
    if (list[0] && list[0].templateId === entry.templateId && JSON.stringify(list[0].values) === JSON.stringify(entry.values)) list.shift();
    list.unshift(entry);
    store.set('history', list.slice(0, 100));
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

  function exportDocx() {
    if (!state.tpl || !confirmMissing()) return;
    const blob = window.DocxBuilder.build(state.lastResult.blocks);
    downloadBlob(blob, currentFileName() + '.docx');
    addHistory();
    toast('Wordファイルを保存しました');
  }

  function printDoc() {
    if (!state.tpl || !confirmMissing()) return;
    addHistory();
    const prev = document.title;
    document.title = currentFileName(); // PDF保存時のファイル名になる
    window.print();
    setTimeout(() => (document.title = prev), 500);
  }

  async function copyText() {
    if (!state.tpl || !confirmMissing()) return;
    const text = blocksToText(state.lastResult.blocks);
    try {
      await navigator.clipboard.writeText(text);
    } catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    addHistory();
    toast('本文をコピーしました');
  }

  // =====================================================================
  // 履歴
  // =====================================================================
  function openHistory() {
    const list = store.get('history', []);
    const box = $('#history-list');
    if (!list.length) {
      box.innerHTML = '<p class="muted">まだ履歴はありません。Word出力・印刷・コピーをすると自動で記録されます。</p>';
    } else {
      box.innerHTML = list
        .map((h) => {
          const d = new Date(h.at);
          const when = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
          return (
            `<div class="hist-item"><div class="hist-main"><div class="hist-title">${esc(h.title)}</div>` +
            `<div class="hist-meta">${esc(h.templateName)} ・ ${when}</div></div>` +
            `<button type="button" class="btn" data-hist-open="${h.id}">この内容で開く</button>` +
            `<button type="button" class="btn ghost" data-hist-del="${h.id}" aria-label="削除">削除</button></div>`
          );
        })
        .join('');
    }
    $('#history-dialog').showModal();
  }

  function onHistoryClick(e) {
    const openId = e.target.dataset.histOpen;
    const delId = e.target.dataset.histDel;
    const list = store.get('history', []);
    if (openId) {
      const h = list.find((x) => x.id === openId);
      if (!h) return;
      if (!findTemplate(h.templateId)) return toast('このひな形は削除されています');
      store.set('draft.' + h.templateId, h.values);
      store.set('dateStyle.' + h.templateId, h.dateStyle);
      $('#history-dialog').close();
      selectTemplate(h.templateId);
      toast('履歴の内容を読み込みました');
    } else if (delId) {
      store.set('history', list.filter((x) => x.id !== delId));
      openHistory();
    }
  }

  // =====================================================================
  // ひな形エディタ
  // =====================================================================
  const editor = { id: null, builtinSource: false, fields: new Map(), timer: null };

  function openEditor(tpl) {
    editor.id = tpl && !tpl.builtin ? tpl.id : null;
    editor.builtinSource = !!(tpl && tpl.builtin);
    editor.fields = new Map((tpl ? tpl.fields : []).map((f) => [f.key, { ...f }]));
    $('#te-heading').textContent = !tpl ? '新しいひな形' : tpl.builtin ? '標準ひな形をコピーして編集' : 'ひな形の編集';
    $('#te-note').hidden = !editor.builtinSource;
    $('#te-name').value = tpl ? (tpl.builtin ? tpl.name + '（コピー）' : tpl.name) : '';
    $('#te-desc').value = tpl ? tpl.description : '';
    $('#te-filename').value = tpl ? tpl.fileName : '';
    $('#te-body').value = tpl
      ? tpl.body
      : '# 〇〇契約書\n\n{{甲_名称}}（以下「甲」という。）と{{乙_名称}}（以下「乙」という。）は、次のとおり契約を締結する。\n\n## 第{{条}}条（目的）\nここに本文を書きます。\n\n{{契約日:date}}\n\n> 甲　{{甲_名称}}\n> 乙　{{乙_名称}}\n';
    $('#te-delete').hidden = !editor.id;
    $('#te-export').hidden = !tpl;
    renderEditorFields();
    $('#tpl-editor').showModal();
    $('#te-name').focus();
  }

  function renderEditorFields() {
    const body = $('#te-body').value;
    const detected = detectFields(body);
    const box = $('#te-fields');
    if (!detected.length) {
      box.innerHTML = '<p class="muted">本文に <code>{{項目名}}</code> を書くと、ここに入力項目が表示されます。</p>';
      return;
    }
    const groups = new Set([DEFAULT_GROUP]);
    editor.fields.forEach((f) => groups.add(f.group));
    $('#te-groups').innerHTML = Array.from(groups).map((g) => `<option value="${esc(g)}">`).join('');

    box.innerHTML = detected
      .map((d) => {
        let f = editor.fields.get(d.key);
        if (!f) {
          f = normalizeField({ key: d.key }, d.type);
          editor.fields.set(d.key, f);
        } else if (d.type !== 'text' && f.type === 'text') {
          f.type = d.type;
        }
        const k = esc(d.key);
        return (
          `<div class="te-field" data-key="${k}">` +
          `<div class="te-field-head"><code>{{${k}}}</code></div>` +
          `<div class="te-grid">` +
          `<label>表示名<input data-prop="label" value="${esc(f.label)}"></label>` +
          `<label>種類<select data-prop="type">${Object.entries(TYPE_LABELS)
            .map(([v, l]) => `<option value="${v}"${v === f.type ? ' selected' : ''}>${l}</option>`)
            .join('')}</select></label>` +
          `<label>グループ<input data-prop="group" list="te-groups" value="${esc(f.group)}"></label>` +
          `<label class="te-options"${f.type === 'select' ? '' : ' hidden'}>選択肢（カンマ区切り）<input data-prop="options" value="${esc(f.options.join(','))}"></label>` +
          `<label class="te-wide">補足説明<input data-prop="hint" value="${esc(f.hint)}" placeholder="入力欄の下に表示されます"></label>` +
          `</div><div class="te-checks">` +
          `<label><input type="checkbox" data-prop="required"${f.required ? ' checked' : ''}> 必須</label>` +
          `<label><input type="checkbox" data-prop="remember"${f.remember ? ' checked' : ''}> 入力を記憶（自社情報など）</label>` +
          (f.type === 'date'
            ? `<label><input type="checkbox" data-prop="today"${f.default === 'today' ? ' checked' : ''}> 初期値を今日にする</label>`
            : '') +
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
    const f = editor.fields.get(card.dataset.key);
    if (!f) return;
    if (prop === 'required' || prop === 'remember') f[prop] = el.checked;
    else if (prop === 'today') f.default = el.checked ? 'today' : undefined;
    else if (prop === 'options') f.options = el.value.split(/[,、]/).map((s) => s.trim()).filter(Boolean);
    else f[prop] = el.value;
    if (prop === 'type') {
      if (f.type === 'checkbox') f.required = false;
      renderEditorFields();
    }
  }

  function editorTemplate() {
    const body = $('#te-body').value;
    const keys = new Set(detectFields(body).map((d) => d.key));
    const fields = Array.from(editor.fields.values()).filter((f) => keys.has(f.key));
    return normalizeTemplate(
      {
        id: editor.id || 'custom-' + Date.now().toString(36),
        name: $('#te-name').value.trim(),
        description: $('#te-desc').value.trim(),
        fileName: $('#te-filename').value.trim(),
        body,
        fields,
      },
      false
    );
  }

  function saveEditor() {
    if (!$('#te-name').value.trim()) {
      $('#te-name').focus();
      return toast('ひな形の名前を入力してください');
    }
    const t = editorTemplate();
    saveCustomTemplate(t);
    $('#tpl-editor').close();
    renderTemplateList();
    selectTemplate(t.id);
    toast('ひな形を保存しました');
  }

  function makeFieldFromSelection() {
    const ta = $('#te-body');
    const { selectionStart: s, selectionEnd: e, value } = ta;
    const selected = value.slice(s, e);
    const name = prompt(
      selected ? `「${selected.slice(0, 30)}」を入力項目に置き換えます。\n項目名を入力してください（例：乙_名称）` : '挿入する入力項目の名前を入力してください（例：乙_名称）'
    );
    if (!name || !name.trim()) return;
    const key = name.trim().replace(/[{}#\/:]/g, '');
    const token = `{{${key}}}`;
    ta.setRangeText(token, s, e, 'end');
    // 同じ文言が他にもあれば一括置換するか確認
    if (selected && selected.length >= 2 && ta.value.includes(selected)) {
      const count = ta.value.split(selected).length - 1;
      if (confirm(`本文の他の場所にも「${selected.slice(0, 30)}」が ${count} 箇所あります。すべて {{${key}}} に置き換えますか？`)) {
        ta.value = ta.value.split(selected).join(token);
      }
    }
    ta.focus();
    renderEditorFields();
  }

  function insertAtCursor(text) {
    const ta = $('#te-body');
    ta.setRangeText(text, ta.selectionStart, ta.selectionEnd, 'end');
    ta.focus();
    renderEditorFields();
  }

  // ---------- 読み込み / 書き出し ----------
  function exportTemplates(templates, name) {
    const data = {
      format: 'contract-maker-templates',
      version: 1,
      templates: templates.map(serializeTemplate),
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    downloadBlob(blob, name + '.json');
  }

  async function importTemplates(file) {
    try {
      const data = JSON.parse(await file.text());
      const list = Array.isArray(data) ? data : data.templates || (data.template ? [data.template] : [data]);
      let n = 0;
      const existing = new Set(allTemplates().map((t) => t.id));
      for (const raw of list) {
        if (!raw || typeof raw.body !== 'string') continue;
        const t = normalizeTemplate(raw, false);
        if (existing.has(t.id) && builtinTemplates.some((b) => b.id === t.id)) t.id = 'custom-' + Date.now().toString(36) + n;
        saveCustomTemplate(t);
        n++;
      }
      renderTemplateList();
      toast(n ? `${n} 件のひな形を読み込みました` : '読み込めるひな形がありませんでした');
    } catch (e) {
      toast('ファイルを読み込めませんでした');
    }
  }

  // =====================================================================
  // 初期化
  // =====================================================================
  function init() {
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
      const wrap = e.target.closest('.field-money');
      if (wrap) e.target.value = formatMoney(e.target.value);
    });
    form.addEventListener('submit', (e) => e.preventDefault());
    // Enter で次の項目へ
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
      const f = state.tpl.fieldMap.get(key);
      const input = $('input', e.target.closest('.field'));
      input.value = todayIso();
      setValue(f, input.value);
    });

    $('#paper').addEventListener('click', (e) => {
      const ph = e.target.closest('.ph');
      if (ph) focusField(ph.dataset.key);
    });

    $$('input[name="dateStyle"]').forEach((r) =>
      r.addEventListener('change', () => {
        state.dateStyle = r.value;
        store.set('dateStyle.' + state.tpl.id, r.value);
        update();
      })
    );

    $('#btn-clear').addEventListener('click', () => {
      if (!confirm('入力内容をクリアしますか？（「記憶」の項目は残ります）')) return;
      store.del('draft.' + state.tpl.id);
      selectTemplate(state.tpl.id);
    });
    $('#btn-docx').addEventListener('click', exportDocx);
    $('#btn-print').addEventListener('click', printDoc);
    $('#btn-copy').addEventListener('click', copyText);
    $('#btn-history').addEventListener('click', openHistory);
    $('#history-list').addEventListener('click', onHistoryClick);
    $('#btn-toggle-preview').addEventListener('click', () => document.body.classList.toggle('show-preview'));
    $('#btn-menu').addEventListener('click', () => document.body.classList.toggle('show-sidebar'));

    $('#btn-new-tpl').addEventListener('click', () => openEditor(null));
    $('#btn-edit-tpl').addEventListener('click', () => openEditor(state.tpl));
    $('#btn-import').addEventListener('click', () => $('#import-file').click());
    $('#import-file').addEventListener('change', (e) => {
      if (e.target.files[0]) importTemplates(e.target.files[0]);
      e.target.value = '';
    });
    $('#btn-export-all').addEventListener('click', () => {
      const customs = customTemplates();
      if (!customs.length) return toast('書き出すマイひな形がありません');
      exportTemplates(customs, 'マイひな形_' + todayIso().replace(/-/g, ''));
    });

    // エディタ
    $('#te-body').addEventListener('input', () => {
      clearTimeout(editor.timer);
      editor.timer = setTimeout(renderEditorFields, 300);
    });
    $('#te-fields').addEventListener('input', onEditorFieldChange);
    $('#te-fields').addEventListener('change', onEditorFieldChange);
    $('#te-save').addEventListener('click', saveEditor);
    $('#te-make-field').addEventListener('click', makeFieldFromSelection);
    $('#te-insert-article').addEventListener('click', () => insertAtCursor('## 第{{条}}条（見出し）\n'));
    $('#te-insert-if').addEventListener('click', () => {
      const ta = $('#te-body');
      const sel = ta.value.slice(ta.selectionStart, ta.selectionEnd) || 'ONのときだけ表示する文章';
      const name = prompt('ON/OFF 項目の名前を入力してください（例：自動更新あり）');
      if (!name || !name.trim()) return;
      insertAtCursor(`{{#if ${name.trim()}}}${sel}{{/if}}`);
    });
    $('#te-delete').addEventListener('click', () => {
      if (!editor.id || !confirm('このひな形を削除しますか？（元に戻せません）')) return;
      deleteCustomTemplate(editor.id);
      store.del('draft.' + editor.id);
      $('#tpl-editor').close();
      if (state.tpl && state.tpl.id === editor.id) {
        state.tpl = null;
        $('#form-pane').hidden = true;
        $('#actions').hidden = true;
        $('#empty-state').hidden = false;
        $('#paper').innerHTML = '';
      }
      renderTemplateList();
      toast('ひな形を削除しました');
    });
    $('#te-export').addEventListener('click', () => {
      const t = editorTemplate();
      exportTemplates([t], t.name || 'ひな形');
    });
    $$('[data-close]').forEach((b) => b.addEventListener('click', () => b.closest('dialog').close()));

    const last = store.get('lastTemplate', null);
    if (last && findTemplate(last)) selectTemplate(last);
    else if (builtinTemplates.length) selectTemplate(builtinTemplates[0].id);
  }

  window.ContractApp = {
    register,
    // テスト・拡張用
    _internal: { normalizeTemplate, evaluate, blocksToText, formatDate, formatMoney, detectFields },
  };

  document.addEventListener('DOMContentLoaded', init);
})();
