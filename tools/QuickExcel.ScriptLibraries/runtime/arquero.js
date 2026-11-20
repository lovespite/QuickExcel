import { fromCSV, fromJSON, fromFixed, fromArrow } from 'arquero';
export * from 'arquero';

// The host owns network consent and bounded response retention. Avoid browser
// TransformStream/TextDecoderStream/DecompressionStream dependencies here.
async function collect(stream, text = false) {
  const reader = stream.getReader(), chunks = [];
  let size = 0, kind;
  try {
    for (;;) {
      const part = await reader.read();
      if (part.done) break;
      const nextKind = typeof part.value === 'string' && text ? 'text' : part.value instanceof Uint8Array ? 'bytes' : null;
      if (!nextKind || kind && kind !== nextKind) throw new TypeError('Expected consistent string or Uint8Array chunks.');
      kind = nextKind;
      size += kind === 'text' ? new TextEncoder().encode(part.value).byteLength : part.value.byteLength;
      if (size > 10 * 1024 * 1024) throw new RangeError('Arquero stream exceeds the 10 MiB input limit.');
      chunks.push(part.value);
    }
  } catch (error) { try { await reader.cancel(); } catch {} throw error; }
  finally { reader.releaseLock(); }
  if (kind === 'text') return chunks.join('');
  const result = new Uint8Array(size); let offset = 0;
  for (const chunk of chunks) { result.set(chunk,offset); offset += chunk.length; }
  return text ? new TextDecoder().decode(result) : result;
}
export async function fromCSVStream(stream, options) { return fromCSV(await collect(stream,true), options); }
export async function fromJSONStream(stream, options) { return fromJSON(await collect(stream,true), options); }
export async function fromFixedStream(stream, options) { return fromFixed(await collect(stream,true), options); }
export async function fromArrowStream(stream, options) { return fromArrow(await collect(stream), options); }
async function response(url, options = {}) {
  if (options.decompress || options.decompress !== null && /\.(?:gz|zz)(?:[?#]|$)/i.test(String(url)))
    throw new TypeError('Compressed Arquero files require an unavailable decompression codec. Supply uncompressed data.');
  const result = await fetch(url, options.fetch);
  if (!result.ok) throw new Error('Arquero HTTP request failed: ' + result.status);
  return result;
}
export async function loadCSV(url, options) { return fromCSV(await (await response(url,options)).text(),options); }
export async function loadJSON(url, options) { return fromJSON(await (await response(url,options)).text(),options); }
export async function loadFixed(url, options) { return fromFixed(await (await response(url,options)).text(),options); }
export async function loadArrow(url, options) { return fromArrow(await (await response(url,options)).arrayBuffer(),options); }
