/*
 * 依存ライブラリなしで .docx (Word) ファイルを生成する小さなビルダー。
 * app.js が作る「ブロック」配列を受け取り、Blob を返す。
 */
(function () {
  'use strict';

  const enc = new TextEncoder();

  // ---------- ZIP (無圧縮) ----------
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  function zip(files, mime) {
    const parts = [];
    const central = [];
    let offset = 0;
    const DOS_DATE = 0x21; // 1980-01-01

    for (const f of files) {
      const name = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const crc = crc32(data);

      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, 0, true);
      lh.setUint16(10, 0, true);
      lh.setUint16(12, DOS_DATE, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true);
      lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true);
      lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, data);

      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, 0, true);
      ch.setUint16(12, 0, true);
      ch.setUint16(14, DOS_DATE, true);
      ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true);
      ch.setUint32(24, data.length, true);
      ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);

      offset += 30 + name.length + data.length;
    }

    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);

    return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: mime });
  }

  // ---------- WordprocessingML ----------
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const MINCHO = 'ＭＳ 明朝';

  function xmlEscape(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function run(text, props) {
    let rPr = '';
    if (props.bold) rPr += '<w:b/><w:bCs/>';
    if (props.size) rPr += `<w:sz w:val="${props.size}"/><w:szCs w:val="${props.size}"/>`;
    if (props.highlight) rPr += '<w:highlight w:val="yellow"/>';
    if (props.underline) rPr += '<w:u w:val="single"/>';
    rPr = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';

    // 値の中の改行は Word の改行 (w:br) にする
    const body = String(text)
      .split('\n')
      .map((line) => `<w:t xml:space="preserve">${xmlEscape(line)}</w:t>`)
      .join('<w:br/>');
    return `<w:r>${rPr}${body}</w:r>`;
  }

  const KIND_STYLE = {
    title: { jc: 'center', bold: true, size: 32, after: 360 },
    heading: { bold: true },
    indent: { ind: 4252 },
    right: { jc: 'right' },
    center: { jc: 'center' },
    para: {},
    blank: {},
  };

  function paragraph(block) {
    if (block.kind === 'pagebreak') return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const st = KIND_STYLE[block.kind] || {};
    let pPr = '';
    if (st.before || st.after) pPr += `<w:spacing w:before="${st.before || 0}" w:after="${st.after || 0}"/>`;
    if (st.ind) pPr += `<w:ind w:left="${st.ind}"/>`;
    if (st.jc) pPr += `<w:jc w:val="${st.jc}"/>`;
    pPr = pPr ? `<w:pPr>${pPr}</w:pPr>` : '';

    const runs = (block.segs || [])
      .map((seg) => {
        const props = { bold: st.bold, size: st.size };
        if (seg.t === 'field' && seg.empty) props.highlight = true;
        return seg.v === '' ? '' : run(seg.v, props);
      })
      .join('');
    return `<w:p>${pPr}${runs}</w:p>`;
  }

  function documentXml(blocks) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:document xmlns:w="${W_NS}"><w:body>` +
      blocks.map(paragraph).join('') +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1701" w:right="1418" w:bottom="1418" w:left="1418" w:header="851" w:footer="992" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>'
    );
  }

  const STYLES_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    `<w:rFonts w:ascii="${MINCHO}" w:eastAsia="${MINCHO}" w:hAnsi="${MINCHO}" w:cs="Times New Roman"/>` +
    '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="21"/>' +
    '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>' +
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr></w:style></w:styles>';

  const CONTENT_TYPES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>';

  const ROOT_RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  const DOC_RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>';

  function build(blocks) {
    return zip(
      [
        { name: '[Content_Types].xml', data: CONTENT_TYPES },
        { name: '_rels/.rels', data: ROOT_RELS },
        { name: 'word/document.xml', data: documentXml(blocks) },
        { name: 'word/_rels/document.xml.rels', data: DOC_RELS },
        { name: 'word/styles.xml', data: STYLES_XML },
      ],
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
  }

  window.DocxBuilder = { build };
})();
