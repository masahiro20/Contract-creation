/*
 * ZIP の読み書き（.docx は ZIP 形式）
 *
 * 変更しなかったファイルは、圧縮データをそのままコピーして書き戻す。
 * → 画像・フォント・書式設定など、差し込み箇所以外は1バイトも変わらない。
 */
(function () {
  'use strict';

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

  async function pipe(bytes, stream) {
    const buf = await new Response(new Blob([bytes]).stream().pipeThrough(stream)).arrayBuffer();
    return new Uint8Array(buf);
  }

  /** ZIP を読み、エントリ一覧を返す（中身はまだ展開しない） */
  function read(buffer) {
    const u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 22 - 65535); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error('ZIP形式ではありません');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const decoder = new TextDecoder();
    const entries = [];
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIPの構造が壊れています');
      const flags = dv.getUint16(p + 8, true);
      const method = dv.getUint16(p + 10, true);
      const nlen = dv.getUint16(p + 28, true);
      const elen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const nameBytes = u8.slice(p + 46, p + 46 + nlen);
      const csize = dv.getUint32(p + 20, true);
      const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      if (flags & 0x1) throw new Error('パスワード付きのファイルは読み込めません');
      entries.push({
        name: decoder.decode(nameBytes),
        nameBytes,
        flags,
        method,
        time: dv.getUint16(p + 12, true),
        date: dv.getUint16(p + 14, true),
        crc: dv.getUint32(p + 16, true),
        csize,
        usize: dv.getUint32(p + 24, true),
        extAttr: dv.getUint32(p + 38, true),
        raw: u8.subarray(start, start + csize),
      });
      p += 46 + nlen + elen + clen;
    }
    return entries;
  }

  async function inflate(entry) {
    if (entry.method === 0) return entry.raw;
    if (entry.method === 8) return pipe(entry.raw, new DecompressionStream('deflate-raw'));
    throw new Error('未対応の圧縮形式です');
  }

  /** 中身を差し替えたエントリを作る（元エントリの日時などは引き継ぐ） */
  async function replaceData(entry, data) {
    const raw = await pipe(data, new CompressionStream('deflate-raw'));
    return { ...entry, method: 8, crc: crc32(data), csize: raw.length, usize: data.length, raw };
  }

  /** エントリ一覧から ZIP を組み立てる */
  function write(entries, mime) {
    const parts = [];
    const central = [];
    let offset = 0;
    for (const e of entries) {
      const flags = e.flags & ~0x0008; // サイズはヘッダーに書くのでデータディスクリプタは使わない
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, flags, true);
      lh.setUint16(8, e.method, true);
      lh.setUint16(10, e.time, true);
      lh.setUint16(12, e.date, true);
      lh.setUint32(14, e.crc, true);
      lh.setUint32(18, e.csize, true);
      lh.setUint32(22, e.usize, true);
      lh.setUint16(26, e.nameBytes.length, true);
      lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), e.nameBytes, e.raw);

      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, flags, true);
      ch.setUint16(10, e.method, true);
      ch.setUint16(12, e.time, true);
      ch.setUint16(14, e.date, true);
      ch.setUint32(16, e.crc, true);
      ch.setUint32(20, e.csize, true);
      ch.setUint32(24, e.usize, true);
      ch.setUint16(28, e.nameBytes.length, true);
      ch.setUint32(38, e.extAttr || 0, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), e.nameBytes);
      offset += 30 + e.nameBytes.length + e.raw.length;
    }
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, entries.length, true);
    end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: mime || 'application/zip' });
  }

  window.Zip = { read, inflate, replaceData, write, crc32 };
})();
