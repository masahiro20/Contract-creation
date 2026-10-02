/*
 * 入力項目の辞書（株式会社ホームランディックの契約書類に合わせたもの）
 *
 * 項目名（key）を書類間で揃えておくと、1件の案件で一度入力した内容が
 * 請負契約書・重要事項説明書・合意書・変更契約書などすべてに差し込まれる。
 *
 * - remember: true … 自社情報など。次回以降も自動で入る
 * - compute       … ほかの項目から自動計算する（入力欄は表示のみ）
 * - dropEmptyPara … 空欄のとき、その行（段落）ごと削除する
 *
 * 氏名は「姓」「名」を別々に差し込む（ひな形の「様」「, 」やスペースは元の文字のまま残す）。
 * 書類ごとの書式（全角数字か半角か等）は、各ひな形の項目設定「数字を全角にする」で決める。
 */
(function () {
  'use strict';

  // ---------- 計算用の小さな関数 ----------
  const num = (v) => {
    const s = String(v == null ? '' : v)
      .replace(/[０-９．]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
      .replace(/[^\d.-]/g, '');
    return s === '' ? null : Number(s);
  };
  const yen = (n) => (n == null || !isFinite(n) ? '' : String(Math.round(n)));
  const area = (n) => (n == null || !isFinite(n) ? '' : n.toFixed(2));
  const sum = (...vals) => {
    const ns = vals.map(num).filter((n) => n != null);
    return ns.length ? ns.reduce((a, b) => a + b, 0) : null;
  };
  // 令和の年だけを返す（例：2026-09-26 → "8"）
  // 配筋工事完了時金・上棟時金の基本ルール：請負額 × 30% を10万円未満切り捨て
  const thirtyPct = (total) => (total == null ? null : Math.floor((total * 3) / 1000000) * 100000);
  const reiwaYear = (iso) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
    if (!m) return '';
    const y = +m[1] - 2018;
    return y >= 1 ? String(y) : '';
  };
  const datePart = (iso, i) => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
    return m ? String(+m[i]) : '';
  };

  window.FIELD_PRESETS = [
    // ---------------- 発注者（お客様） ----------------
    { key: '発注者_姓', label: '発注者 姓', group: '発注者（お客様）', placeholder: '例：山田' },
    { key: '発注者_名', label: '発注者 名', group: '発注者（お客様）', placeholder: '例：太郎' },
    { key: '発注者2_名', label: '連名者 名', group: '発注者（お客様）', required: false, placeholder: '例：花子', hint: '連名（ご夫婦など）の場合のみ' },
    {
      key: '発注者_氏名詰め',
      label: '発注者 氏名（詰め表記）',
      group: '発注者（お客様）',
      hint: '合意書・変更契約書用。例：山田太郎様,花子（後ろの「様」はひな形側）',
      compute: (v) => {
        if (!v['発注者_姓'] && !v['発注者_名']) return '';
        const main = `${v['発注者_姓'] || ''}${v['発注者_名'] || ''}`;
        return v['発注者2_名'] ? `${main}様,${v['発注者2_名']}` : main;
      },
    },
    { key: '発注者_住所', label: '発注者 現住所', group: '発注者（お客様）', required: false },

    // ---------------- 工事の概要 ----------------
    { key: '建築地', label: '工事現場住所（建築地）', group: '工事の概要', placeholder: '例：愛知県○○市○○町一丁目12番,13番', hint: '地名地番。書類によって「,」「、」の表記が異なる場合は注意' },
    { key: '階数', label: '階数', type: 'select', options: ['２', '３', '平屋'], group: '工事の概要', hint: '「木造ガルバリウム鋼板葺２階建て」の数字部分' },
    { key: '面積_1階', label: '1階 床面積（㎡）', type: 'number', group: '工事の概要' },
    { key: '面積_2階', label: '2階 床面積（㎡）', type: 'number', group: '工事の概要', required: false },
    { key: '面積_3階', label: '3階 床面積（㎡）', type: 'number', group: '工事の概要', required: false },
    {
      key: '延床面積',
      label: '延べ床面積（㎡）',
      type: 'number',
      group: '工事の概要',
      hint: '各階の合計を自動計算',
      compute: (v) => area(sum(v['面積_1階'], v['面積_2階'], v['面積_3階'])),
    },
    { key: '面積_ポーチバルコニー', label: 'ポーチ・バルコニー面積（㎡）', type: 'number', group: '工事の概要', required: false, dropEmptyPara: true, hint: '無い場合は空欄（行ごと削除されます）' },

    // ---------------- 日付 ----------------
    { key: '契約日', label: '請負契約日', type: 'date', group: '日付', default: 'today' },
    { key: '契約日_年', label: '請負契約日（令和の年）', group: '日付', compute: (v) => reiwaYear(v['契約日']), hint: '「令和 ● 年」の●部分' },
    { key: '契約日_月', label: '請負契約日（月）', group: '日付', compute: (v) => datePart(v['契約日'], 2) },
    { key: '契約日_日', label: '請負契約日（日）', group: '日付', compute: (v) => datePart(v['契約日'], 3) },

    // ---------------- 請負代金 ----------------
    { key: '請負代金', label: '請負代金（税込）', type: 'money', group: '請負代金', hint: 'お見積書の「金額(税込)」' },
    { key: '工事価格', label: '工事価格（税抜）', type: 'money', group: '請負代金', compute: (v) => (num(v['請負代金']) == null ? '' : yen(num(v['請負代金']) / 1.1)), hint: '請負代金 ÷ 1.1 で自動計算' },
    {
      key: '消費税',
      label: '消費税（10%）',
      type: 'money',
      group: '請負代金',
      compute: (v) => {
        const t = num(v['請負代金']);
        return t == null ? '' : yen(t - Math.round(t / 1.1));
      },
    },

    // ---------------- 支払い ----------------
    { key: '支払_契約時', label: '契約時 金額', type: 'money', group: '請負工事代金の支払い', default: '1000000' },
    { key: '支払_配筋', label: '配筋工事完了時 金額', type: 'money', group: '請負工事代金の支払い', hint: '基本は請負代金×30%（10万円未満切り捨て）' },
    { key: '支払_上棟', label: '上棟時 金額', type: 'money', group: '請負工事代金の支払い', hint: '基本は請負代金×30%（10万円未満切り捨て）' },
    {
      key: '支払_完成',
      label: '完成引き渡し時 金額',
      type: 'money',
      group: '請負工事代金の支払い',
      hint: '請負代金 − 契約時 − 配筋 − 上棟 で自動計算',
      compute: (v) => {
        const t = num(v['請負代金']);
        if (t == null) return '';
        return yen(t - (num(v['支払_契約時']) || 0) - (num(v['支払_配筋']) || 0) - (num(v['支払_上棟']) || 0));
      },
    },

    // ---------------- 重要事項説明 ----------------
    { key: '説明建築士_氏名', label: '説明をする建築士', type: 'select', options: ['伊 岐 見 恭 子', '中 谷 真 弘'], group: '重要事項説明', remember: true },
    { key: '説明建築士_資格', label: '建築士の資格', type: 'select', options: ['一級', '二級'], group: '重要事項説明', remember: true },

    // ---------------- 設計契約 ----------------
    { key: '設計契約金', label: '設計契約金（税込）', type: 'money', group: '設計契約' },
    { key: '設計契約金_期日', label: '設計契約金の支払期日', type: 'date', group: '設計契約' },
    { key: '構造_設計', label: '構造（設計契約書）', group: '設計契約', placeholder: '例：２階建て 木造軸組在来工法 / 平屋建て SE構法' },

    // ---------------- 変更契約・一部変更合意 ----------------
    { key: '原契約日', label: '原契約（請負契約）の締結日', type: 'date', group: '変更契約' },
    { key: '原請負代金', label: '原契約の請負代金（税込）', type: 'money', group: '変更契約' },
    { key: '変更内容', label: '追加変更工事の内容', type: 'textarea', group: '変更契約' },
    { key: '変更_差額', label: '追加変更工事 差額（税込）', type: 'money', group: '変更契約' },
    {
      key: '変更後_請負代金',
      label: '変更後の請負代金（税込）',
      type: 'money',
      group: '変更契約',
      hint: '原契約の請負代金 ＋ 差額',
      compute: (v) => (num(v['原請負代金']) == null ? '' : yen(num(v['原請負代金']) + (num(v['変更_差額']) || 0))),
    },
    {
      key: '変更後_配筋',
      label: '変更後 配筋工事完了時金',
      type: 'money',
      group: '変更契約',
      hint: '変更後の請負代金×30%（10万円未満切り捨て）',
      compute: (v) => (num(v['原請負代金']) == null ? '' : yen(thirtyPct(num(v['原請負代金']) + (num(v['変更_差額']) || 0)))),
    },
    {
      key: '変更後_上棟',
      label: '変更後 上棟時金',
      type: 'money',
      group: '変更契約',
      hint: '変更後の請負代金×30%（10万円未満切り捨て）',
      compute: (v) => (num(v['原請負代金']) == null ? '' : yen(thirtyPct(num(v['原請負代金']) + (num(v['変更_差額']) || 0)))),
    },
    {
      key: '変更後_完成',
      label: '変更後 完成引き渡し時金',
      type: 'money',
      group: '変更契約',
      hint: '変更後の請負代金 − 契約時 − 配筋 − 上棟（残額を1円単位で）',
      compute: (v) => {
        if (num(v['原請負代金']) == null) return '';
        const total = num(v['原請負代金']) + (num(v['変更_差額']) || 0);
        return yen(total - (num(v['支払_契約時']) || 0) - 2 * thirtyPct(total));
      },
    },
  ];
})();
