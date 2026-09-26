/*
 * よく使う入力項目（住宅会社・施工会社 × お客様の建築工事請負向け）
 *
 * 項目名（key）を書類間で揃えておくと、1件の案件で一度入力した内容が
 * 請負契約書・約款・合意書など、すべての書類に共通で差し込まれる。
 * remember: true の項目は自社情報として記憶され、次回以降も自動で入る。
 */
window.FIELD_PRESETS = [
  // 注文者（お客様）
  { key: '注文者_氏名', label: '注文者（施主）氏名', group: '注文者（お客様）' },
  { key: '注文者_フリガナ', label: 'フリガナ', group: '注文者（お客様）', required: false },
  { key: '注文者_住所', label: '注文者 住所', group: '注文者（お客様）' },
  { key: '注文者_電話', label: '注文者 電話番号', group: '注文者（お客様）', required: false },
  { key: '注文者2_氏名', label: '連名者 氏名', group: '注文者（お客様）', required: false, hint: '共有名義の場合' },
  { key: '注文者2_住所', label: '連名者 住所', group: '注文者（お客様）', required: false },

  // 請負者（自社）
  { key: '請負者_会社名', label: '請負者 会社名', group: '請負者（自社）', remember: true },
  { key: '請負者_住所', label: '請負者 所在地', group: '請負者（自社）', remember: true },
  { key: '請負者_代表者', label: '代表者 役職・氏名', group: '請負者（自社）', remember: true },
  { key: '請負者_電話', label: '請負者 電話番号', group: '請負者（自社）', remember: true, required: false },
  { key: '建設業許可番号', label: '建設業許可番号', group: '請負者（自社）', remember: true },
  { key: '建築士事務所登録', label: '建築士事務所登録番号', group: '請負者（自社）', remember: true, required: false },
  { key: '担当者', label: '担当者（営業・現場）', group: '請負者（自社）', required: false },

  // 工事の概要
  { key: '工事名称', label: '工事名称', group: '工事の概要', placeholder: '例：〇〇様邸新築工事' },
  { key: '工事場所', label: '工事場所（地名地番）', group: '工事の概要' },
  { key: '建物_構造', label: '構造・階数', group: '工事の概要', placeholder: '例：木造2階建て' },
  { key: '建物_用途', label: '用途', group: '工事の概要', required: false, placeholder: '例：専用住宅' },
  { key: '延床面積', label: '延床面積（㎡）', type: 'number', group: '工事の概要' },
  { key: '敷地面積', label: '敷地面積（㎡）', type: 'number', group: '工事の概要', required: false },

  // 工期・日付
  { key: '契約日', label: '契約締結日', type: 'date', group: '工期・日付', default: 'today' },
  { key: '着工日', label: '着工日', type: 'date', group: '工期・日付' },
  { key: '上棟日', label: '上棟予定日', type: 'date', group: '工期・日付', required: false },
  { key: '完成日', label: '完成日（竣工）', type: 'date', group: '工期・日付' },
  { key: '引渡日', label: '引渡日', type: 'date', group: '工期・日付' },

  // 請負代金
  { key: '請負代金', label: '請負代金額（税込）', type: 'money', group: '請負代金' },
  { key: '工事価格', label: '工事価格（税抜）', type: 'money', group: '請負代金' },
  { key: '消費税額', label: '取引に係る消費税額', type: 'money', group: '請負代金' },

  // 支払方法
  { key: '支払1_金額', label: '契約時 金額', type: 'money', group: '支払方法' },
  { key: '支払1_期日', label: '契約時 支払期日', type: 'date', group: '支払方法' },
  { key: '支払2_金額', label: '着工時 金額', type: 'money', group: '支払方法', required: false },
  { key: '支払2_期日', label: '着工時 支払期日', type: 'date', group: '支払方法', required: false },
  { key: '支払3_金額', label: '上棟時 金額', type: 'money', group: '支払方法', required: false },
  { key: '支払3_期日', label: '上棟時 支払期日', type: 'date', group: '支払方法', required: false },
  { key: '支払4_金額', label: '完成・引渡時 金額', type: 'money', group: '支払方法', required: false },
  { key: '支払4_期日', label: '完成・引渡時 支払期日', type: 'date', group: '支払方法', required: false },
  { key: '振込先', label: '振込先口座', type: 'textarea', group: '支払方法', remember: true, required: false },

  // 変更・追加工事の合意書
  { key: '原契約日', label: '原契約の締結日', type: 'date', group: '変更・追加工事' },
  { key: '変更内容', label: '変更・追加の内容', type: 'textarea', group: '変更・追加工事' },
  { key: '増減額', label: '増減額（税込）', type: 'money', group: '変更・追加工事' },
  { key: '変更後_請負代金', label: '変更後の請負代金額（税込）', type: 'money', group: '変更・追加工事' },
  { key: '変更後_完成日', label: '変更後の完成日', type: 'date', group: '変更・追加工事', required: false },
  { key: '合意日', label: '合意日', type: 'date', group: '変更・追加工事', default: 'today' },
];
