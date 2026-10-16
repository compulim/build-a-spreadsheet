/**
 * Minimal, dependency-free ZIP reader/writer.
 *
 * Uses the native `CompressionStream`/`DecompressionStream` (`deflate-raw`) when available.
 * Supports "stored" (0) and "deflate" (8) methods. ZIP64 is not supported.
 *
 * @typedef {{ name: string; data: Uint8Array }} ZipEntry
 */

const LOCAL_FILE_HEADER_SIGNATURE = 0x04034b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY_SIGNATURE = 0x06054b50;

const METHOD_STORE = 0;
const METHOD_DEFLATE = 8;

const FLAG_UTF8 = 0x0800;

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);

  for (let index = 0; index < 256; index++) {
    let crc = index;

    for (let bit = 0; bit < 8; bit++) {
      crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
    }

    table[index] = crc >>> 0;
  }

  return table;
})();

/**
 * @param {Uint8Array} data
 * @returns {number}
 */
export function crc32(data) {
  let crc = 0xffffffff;

  for (let index = 0; index < data.length; index++) {
    crc = CRC_TABLE[(crc ^ data[index]) & 0xff] ^ (crc >>> 8);
  }

  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * @param {Uint8Array} data
 * @param {TransformStream<Uint8Array, Uint8Array>} transformStream
 * @returns {Promise<Uint8Array>}
 */
async function transform(data, transformStream) {
  const stream = new Blob([data]).stream().pipeThrough(transformStream);

  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/**
 * @param {Date} date
 * @returns {{ date: number; time: number }}
 */
function toDOSDateTime(date) {
  // DOS date/time can only represent 1980-01-01 to 2107-12-31.
  if (date.getFullYear() < 1980) {
    date = new Date(1980, 0, 1);
  } else if (date.getFullYear() > 2107) {
    date = new Date(2107, 11, 31, 23, 59, 58);
  }

  return {
    date: ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1)
  };
}

/**
 * Creates a ZIP archive.
 *
 * @param {ZipEntry[]} entries
 * @param {{ compress?: boolean; date?: Date }} [options]
 * @returns {Promise<Uint8Array>}
 */
export async function createZip(entries, { compress = true, date = new Date() } = {}) {
  const encoder = new TextEncoder();
  const canCompress = compress && typeof CompressionStream !== 'undefined';
  const { date: dosDate, time: dosTime } = toDOSDateTime(date);

  /** @type {Uint8Array[]} */
  const localParts = [];
  /** @type {Uint8Array[]} */
  const centralParts = [];
  let offset = 0;

  for (const { name, data } of entries) {
    const nameBytes = encoder.encode(name);
    const crc = crc32(data);
    let method = METHOD_STORE;
    let content = data;

    if (canCompress && data.length) {
      const compressed = await transform(data, new CompressionStream('deflate-raw'));

      if (compressed.length < data.length) {
        method = METHOD_DEFLATE;
        content = compressed;
      }
    }

    const localHeader = new Uint8Array(30 + nameBytes.length);
    const local = new DataView(localHeader.buffer);

    local.setUint32(0, LOCAL_FILE_HEADER_SIGNATURE, true);
    local.setUint16(4, 20, true); // Version needed to extract.
    local.setUint16(6, FLAG_UTF8, true);
    local.setUint16(8, method, true);
    local.setUint16(10, dosTime, true);
    local.setUint16(12, dosDate, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, content.length, true);
    local.setUint32(22, data.length, true);
    local.setUint16(26, nameBytes.length, true);
    local.setUint16(28, 0, true); // Extra field length.
    localHeader.set(nameBytes, 30);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const central = new DataView(centralHeader.buffer);

    central.setUint32(0, CENTRAL_DIRECTORY_SIGNATURE, true);
    central.setUint16(4, 20, true); // Version made by.
    central.setUint16(6, 20, true); // Version needed to extract.
    central.setUint16(8, FLAG_UTF8, true);
    central.setUint16(10, method, true);
    central.setUint16(12, dosTime, true);
    central.setUint16(14, dosDate, true);
    central.setUint32(16, crc, true);
    central.setUint32(20, content.length, true);
    central.setUint32(24, data.length, true);
    central.setUint16(28, nameBytes.length, true);
    // Extra field length, comment length, disk number, internal and external attributes are all zero.
    central.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);

    localParts.push(localHeader, content);
    centralParts.push(centralHeader);

    offset += localHeader.length + content.length;
  }

  const centralSize = centralParts.reduce((size, part) => size + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);

  endView.setUint32(0, END_OF_CENTRAL_DIRECTORY_SIGNATURE, true);
  endView.setUint16(8, entries.length, true);
  endView.setUint16(10, entries.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);

  const result = new Uint8Array(offset + centralSize + end.length);
  let position = 0;

  for (const part of [...localParts, ...centralParts, end]) {
    result.set(part, position);
    position += part.length;
  }

  return result;
}

/**
 * Reads all file entries (directories are skipped) from a ZIP archive.
 *
 * @param {ArrayBuffer | Uint8Array} input
 * @returns {Promise<ZipEntry[]>}
 */
export async function readZip(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const decoder = new TextDecoder();

  let endOffset = -1;

  // End of central directory record is at least 22 bytes, followed by an optional comment of up to 65535 bytes.
  for (let index = bytes.length - 22; index >= Math.max(0, bytes.length - 22 - 0xffff); index--) {
    if (view.getUint32(index, true) === END_OF_CENTRAL_DIRECTORY_SIGNATURE) {
      endOffset = index;
      break;
    }
  }

  if (endOffset === -1) {
    throw new Error('Invalid ZIP file: cannot find end of central directory.');
  }

  const entryCount = view.getUint16(endOffset + 10, true);
  let offset = view.getUint32(endOffset + 16, true);

  /** @type {ZipEntry[]} */
  const entries = [];

  for (let index = 0; index < entryCount; index++) {
    if (offset + 46 > bytes.length || view.getUint32(offset, true) !== CENTRAL_DIRECTORY_SIGNATURE) {
      throw new Error('Invalid ZIP file: corrupted central directory.');
    }

    const method = view.getUint16(offset + 10, true);
    const crc = view.getUint32(offset + 16, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));

    offset += 46 + nameLength + extraLength + commentLength;

    if (name.endsWith('/')) {
      continue;
    }

    if (localOffset + 30 > bytes.length || view.getUint32(localOffset, true) !== LOCAL_FILE_HEADER_SIGNATURE) {
      throw new Error(`Invalid ZIP file: corrupted local header for "${name}".`);
    }

    const dataOffset =
      localOffset + 30 + view.getUint16(localOffset + 26, true) + view.getUint16(localOffset + 28, true);

    if (dataOffset + compressedSize > bytes.length) {
      throw new Error(`Invalid ZIP file: truncated data for "${name}".`);
    }

    const content = bytes.subarray(dataOffset, dataOffset + compressedSize);
    let data;

    if (method === METHOD_STORE) {
      data = content.slice();
    } else if (method === METHOD_DEFLATE) {
      data = await transform(content, new DecompressionStream('deflate-raw'));
    } else {
      throw new Error(`Unsupported ZIP compression method ${method} for "${name}".`);
    }

    if (data.length !== uncompressedSize || crc32(data) !== crc) {
      throw new Error(`Invalid ZIP file: checksum mismatch for "${name}".`);
    }

    entries.push({ name, data });
  }

  return entries;
}
