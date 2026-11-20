// Pure data helpers. These objects do not grant filesystem or network access.
import URL from 'core-js-pure/actual/url/index.js';
Object.defineProperty(globalThis, 'URL', { value: URL });
class File extends Blob {
  constructor(parts, name, options = {}) {
    super(parts, options);
    const timestamp = options.lastModified === undefined ? Date.now() : Number(options.lastModified);
    Object.defineProperties(this, {
      name: { value: String(name).toWellFormed(), enumerable: true },
      lastModified: { value: Number.isFinite(timestamp) ? Math.trunc(timestamp) : 0, enumerable: true },
      webkitRelativePath: { value: '', enumerable: true }
    });
  }
  get [Symbol.toStringTag]() { return 'File'; }
}
Object.defineProperty(globalThis, 'File', { value: File });
