namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    // QuickExcel-owned, bounded Fetch subset. HTTP is brokered by the host;
    // these objects expose no browser, operating-system, or CLR handles.
    private const string NetworkBootstrap = """
        const NativePromise = Promise;
        const promiseResolve = NativePromise.resolve.bind(NativePromise);
        const promiseReject = NativePromise.reject.bind(NativePromise);
        const networkHosts = new Set();
        const permissions = freeze({ network: freeze({ request(hostname) {
            // A shared dependency may also be loaded in the isolated UDF VM.
            if (!actionMode) return;
            if (loaded || registrationRole(true) !== 4) throw new TypeError('Network permissions must be declared at the top level of the initialization root module.');
            if (typeof hostname !== 'string' || !hostname || hostname.length > 253 || /[\s\/?#@\\]/.test(hostname))
                throw new TypeError('Expected an exact domain name, IPv4 address, or IPv6 address.');
            if (networkHosts.size >= 64 && !networkHosts.has(hostname)) throw new TypeError('At most 64 network hosts may be declared.');
            networkHosts.add(hostname);
        } }) });
        let maxFetchBody = 8 * 1024 * 1024;
        let maxResponseBodies = 16 * 1024 * 1024;
        const bodySlots = [], blobBytes = new WeakMap(), headerValues = new WeakMap(), responseValues = new WeakMap();
        let retainedBodyBytes = 0, fetchSequence = 0, actionProgress = null, fetchAllowed = false;
        const fetchWaiters = new Map();
        function configureNetwork(request) {
            const limit = request.maxNetworkBodyBytes ?? 8 * 1024 * 1024;
            if (!Number.isInteger(limit) || limit < 1024 * 1024 || limit > 10 * 1024 * 1024)
                throw new TypeError('Invalid host network body limit.');
            maxFetchBody = limit;
            maxResponseBodies = limit > 4 * 1024 * 1024 ? 2 * limit : 8 * 1024 * 1024;
            evictResponseBodies();
        }
        // Private fetch wire bodies use base64url so JSON's HTML-safe encoder
        // cannot expand '+' into six bytes and defeat the 16 MiB frame budget.
        const alphabet64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
        const alphabet64Codes = new Uint8Array(64), alphabet64Values = new Uint8Array(128).fill(255);
        for (let i = 0; i < 64; i++) { alphabet64Codes[i] = alphabet64.charCodeAt(i); alphabet64Values[alphabet64Codes[i]] = i; }
        alphabet64Values[43] = 62; alphabet64Values[47] = 63; // Also accept ordinary base64 replies.
        function bodyLimitError(kind = 'Body') { return new TypeError(kind + ' exceeds the configured ' + (maxFetchBody / 1048576) + ' MiB body limit.'); }
        function characters(units) { return String.fromCharCode.apply(null, units); }
        function encode64(bytes) {
            if (bytes.length > maxFetchBody) throw bodyLimitError();
            const chunks = [];
            for (let start = 0; start < bytes.length; start += 3072) {
                const end = Math.min(start + 3072, bytes.length), codes = new Uint8Array(Math.ceil((end - start) / 3) * 4);
                let at = 0;
                for (let i = start; i < end; i += 3) {
                    const n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
                    codes[at++] = alphabet64Codes[n >>> 18]; codes[at++] = alphabet64Codes[(n >>> 12) & 63];
                    codes[at++] = i + 1 < end ? alphabet64Codes[(n >>> 6) & 63] : 61;
                    codes[at++] = i + 2 < end ? alphabet64Codes[n & 63] : 61;
                }
                chunks.push(characters(codes));
            }
            return chunks.join('');
        }
        function decode64(text) {
            const length = text.length / 4 * 3 - (text.endsWith('==') ? 2 : text.endsWith('=') ? 1 : 0);
            if (text.length % 4 || length < 0) throw new TypeError('Invalid response body encoding.');
            if (length > maxFetchBody) throw bodyLimitError('Response');
            const bytes = new Uint8Array(length); let at = 0;
            for (let i = 0; i < text.length; i += 4) {
                const a = alphabet64Values[text.charCodeAt(i)], b = alphabet64Values[text.charCodeAt(i + 1)];
                const c = text.charCodeAt(i + 2) === 61 ? 0 : alphabet64Values[text.charCodeAt(i + 2)];
                const d = text.charCodeAt(i + 3) === 61 ? 0 : alphabet64Values[text.charCodeAt(i + 3)];
                if (!(a < 64 && b < 64 && c < 64 && d < 64)) throw new TypeError('Invalid response body encoding.');
                const n = (a << 18) | (b << 12) | (c << 6) | d;
                if (at < length) bytes[at++] = n >>> 16;
                if (at < length) bytes[at++] = n >>> 8;
                if (at < length) bytes[at++] = n;
            }
            return bytes;
        }
        class TextEncoder {
            get encoding() { return 'utf-8'; }
            encode(input = '') {
                const text = String(input);
                if (text.length > maxFetchBody) throw bodyLimitError('Text');
                let length = 0;
                for (let i = 0; i < text.length; i++) {
                    const cp = text.charCodeAt(i);
                    if (cp < 128) length++;
                    else if (cp < 2048) length += 2;
                    else if (cp >= 0xd800 && cp <= 0xdbff && text.charCodeAt(i + 1) >= 0xdc00 && text.charCodeAt(i + 1) <= 0xdfff) { length += 4; i++; }
                    else length += 3;
                    if (length > maxFetchBody) throw bodyLimitError('Text');
                }
                const bytes = new Uint8Array(length); let at = 0;
                for (let i = 0; i < text.length; i++) {
                    let cp = text.codePointAt(i);
                    if (cp > 65535) i++;
                    else if (cp >= 0xd800 && cp <= 0xdfff) cp = 0xfffd;
                    if (cp < 128) bytes[at++] = cp;
                    else if (cp < 2048) { bytes[at++] = 192 | cp >> 6; bytes[at++] = 128 | cp & 63; }
                    else if (cp < 65536) { bytes[at++] = 224 | cp >> 12; bytes[at++] = 128 | cp >> 6 & 63; bytes[at++] = 128 | cp & 63; }
                    else { bytes[at++] = 240 | cp >> 18; bytes[at++] = 128 | cp >> 12 & 63; bytes[at++] = 128 | cp >> 6 & 63; bytes[at++] = 128 | cp & 63; }
                }
                return bytes;
            }
        }
        class TextDecoder {
            constructor(label = 'utf-8', options = {}) {
                if (!['utf-8', 'utf8', 'unicode-1-1-utf-8'].includes(String(label).toLowerCase())) throw new RangeError('Only UTF-8 decoding is supported.');
                this.fatal = !!options.fatal; this.ignoreBOM = !!options.ignoreBOM; this.encoding = 'utf-8';
            }
            decode(input = new Uint8Array(), options = {}) {
                if (options.stream) throw new TypeError('Streaming TextDecoder is unsupported.');
                const bytes = rawBytes(input), chunks = [], units = new Uint16Array(4096); let at = 0;
                const put = unit => { units[at++] = unit; if (at === units.length) { chunks.push(characters(units)); at = 0; } };
                const invalid = () => { if (this.fatal) throw new TypeError('Invalid UTF-8.'); put(0xfffd); };
                for (let i = 0; i < bytes.length;) {
                    const first = bytes[i++];
                    if (first < 128) { put(first); continue; }
                    const size = first >= 194 && first <= 223 ? 2 : first >= 224 && first <= 239 ? 3 : first >= 240 && first <= 244 ? 4 : 0;
                    if (!size) { invalid(); continue; }
                    let cp = first & (127 >> size), read = 1;
                    while (read < size && i < bytes.length && (bytes[i] & 192) === 128 &&
                        !(read === 1 && (first === 224 && bytes[i] < 160 || first === 237 && bytes[i] >= 160 || first === 240 && bytes[i] < 144 || first === 244 && bytes[i] >= 144))) {
                        cp = cp << 6 | bytes[i++] & 63; read++;
                    }
                    if (read !== size) invalid();
                    else if (cp <= 0xffff) put(cp);
                    else { cp -= 0x10000; put(0xd800 | cp >> 10); put(0xdc00 | cp & 1023); }
                }
                if (at) chunks.push(characters(units.subarray(0, at)));
                const out = chunks.join('');
                return !this.ignoreBOM && out[0] === '\ufeff' ? out.slice(1) : out;
            }
        }
        function rawBytes(value) {
            let bytes;
            if (value instanceof ArrayBuffer) bytes = new Uint8Array(value);
            else if (ArrayBuffer.isView(value)) bytes = new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
            else throw new TypeError('Expected an ArrayBuffer or typed array.');
            if (bytes.length > maxFetchBody) throw bodyLimitError();
            return bytes;
        }
        class Blob {
            constructor(parts = [], options = {}) {
                const chunks = []; let size = 0;
                for (const part of parts) {
                    const chunk = blobBytes.has(part) ? blobBytes.get(part) : typeof part === 'string' ? new TextEncoder().encode(part) : rawBytes(part);
                    size += chunk.length;
                    if (size > maxFetchBody) throw bodyLimitError('Blob');
                    chunks.push(chunk);
                }
                const data = new Uint8Array(size); let at = 0;
                for (const chunk of chunks) { data.set(chunk, at); at += chunk.length; }
                blobBytes.set(this, data); Object.defineProperty(this, 'type', { value: String(options.type || '').toLowerCase() });
            }
            get size() { return blobBytes.get(this).length; }
            text() { return promiseResolve(new TextDecoder().decode(blobBytes.get(this))); }
            arrayBuffer() { return promiseResolve(blobBytes.get(this).slice().buffer); }
            slice(start = 0, end = this.size, type = '') { return new Blob([blobBytes.get(this).slice(start, end)], { type }); }
        }
        const encodeForm = value => encodeURIComponent(String(value).toWellFormed()).replace(/%20/g, '+').replace(/[!'()~]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase());
        const decodeForm = value => { try { return decodeURIComponent(value.replace(/\+/g, ' ')); } catch { throw new TypeError('Invalid URL-encoded form.'); } };
        class URLSearchParams {
            constructor(init = '') {
                this._pairs = [];
                if (typeof init === 'string') {
                    for (const part of init.replace(/^\?/, '').split('&')) if (part) {
                        const at = part.indexOf('='); this.append(decodeForm(at < 0 ? part : part.slice(0, at)), decodeForm(at < 0 ? '' : part.slice(at + 1)));
                    }
                } else if (init && typeof init[Symbol.iterator] === 'function') {
                    for (const pair of init) { if (pair.length !== 2) throw new TypeError('Expected name/value pairs.'); this.append(pair[0], pair[1]); }
                } else for (const key of ownKeys(init)) this.append(key, init[key]);
            }
            append(name, value) { this._pairs.push([String(name), String(value)]); }
            set(name, value) { this.delete(name); this.append(name, value); }
            get(name) { return this._pairs.find(pair => pair[0] === String(name))?.[1] ?? null; }
            getAll(name) { return this._pairs.filter(pair => pair[0] === String(name)).map(pair => pair[1]); }
            has(name) { return this.get(name) !== null; }
            delete(name) { this._pairs = this._pairs.filter(pair => pair[0] !== String(name)); }
            sort() { this._pairs.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0); }
            get size() { return this._pairs.length; }
            *entries() { for (const pair of this._pairs) yield pair.slice(); }
            *keys() { for (const pair of this._pairs) yield pair[0]; }
            *values() { for (const pair of this._pairs) yield pair[1]; }
            [Symbol.iterator]() { return this.entries(); }
            forEach(callback, thisArg) { for (const [key, value] of this) callback.call(thisArg, value, key, this); }
            toString() { return this._pairs.map(pair => encodeForm(pair[0]) + '=' + encodeForm(pair[1])).join('&'); }
        }
        class FormData {
            constructor() { this._pairs = []; }
            append(name, value, filename) { this._pairs.push([String(name), value instanceof Blob ? value : String(value), value instanceof Blob ? String(filename ?? 'blob') : null]); }
            set(name, value, filename) { this.delete(name); this.append(name, value, filename); }
            get(name) { return this._pairs.find(pair => pair[0] === String(name))?.[1] ?? null; }
            getAll(name) { return this._pairs.filter(pair => pair[0] === String(name)).map(pair => pair[1]); }
            has(name) { return this.get(name) !== null; }
            delete(name) { this._pairs = this._pairs.filter(pair => pair[0] !== String(name)); }
            *entries() { for (const pair of this._pairs) yield pair.slice(0, 2); }
            *keys() { for (const pair of this._pairs) yield pair[0]; }
            *values() { for (const pair of this._pairs) yield pair[1]; }
            [Symbol.iterator]() { return this.entries(); }
            forEach(callback, thisArg) { for (const [key, value] of this) callback.call(thisArg, value, key, this); }
        }
        class Headers {
            constructor(init = undefined) {
                headerValues.set(this, { map: new Map(), immutable: false });
                if (init === undefined) return;
                if (init && typeof init[Symbol.iterator] === 'function') {
                    for (const pair of init) { if (pair.length !== 2) throw new TypeError('Expected header pairs.'); this.append(pair[0], pair[1]); }
                } else for (const name of ownKeys(init)) this.append(name, init[name]);
            }
            _name(name) { name = String(name).toLowerCase(); if (!/^[!#$%&'*+.^_`|~0-9a-z-]+$/.test(name)) throw new TypeError('Invalid header name.'); return name; }
            _write() { const state = headerValues.get(this); if (state.immutable) throw new TypeError('Response headers are immutable.'); return state.map; }
            _value(value) { value = String(value).trim(); if (/[\r\n\0]/.test(value)) throw new TypeError('Invalid header value.'); return value; }
            append(name, value) { name = this._name(name); value = this._value(value); const map = this._write(); map.set(name, map.has(name) ? map.get(name) + ', ' + value : value); }
            set(name, value) { this._write().set(this._name(name), this._value(value)); }
            delete(name) { this._write().delete(this._name(name)); }
            get(name) { return headerValues.get(this).map.get(this._name(name)) ?? null; }
            has(name) { return headerValues.get(this).map.has(this._name(name)); }
            *entries() { yield* [...headerValues.get(this).map.entries()].sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0); }
            *keys() { for (const pair of this) yield pair[0]; }
            *values() { for (const pair of this) yield pair[1]; }
            [Symbol.iterator]() { return this.entries(); }
            forEach(callback, thisArg) { for (const [key, value] of this) callback.call(thisArg, value, key, this); }
        }
        function bodyBytes(body, headers) {
            if (body === null || body === undefined) return null;
            if (body instanceof FormData) {
                const boundary = '----QuickExcel' + Math.random().toString(36).slice(2), parts = [];
                const quote = text => String(text).replace(/\r/g, '%0D').replace(/\n/g, '%0A').replace(/"/g, '%22');
                for (const [name, value, filename] of body._pairs) {
                    let head = '--' + boundary + '\r\nContent-Disposition: form-data; name="' + quote(name) + '"';
                    if (value instanceof Blob) head += '; filename="' + quote(filename) + '"\r\nContent-Type: ' + (value.type || 'application/octet-stream');
                    parts.push(head + '\r\n\r\n', value, '\r\n');
                }
                parts.push('--' + boundary + '--\r\n');
                if (!headers.has('content-type')) headers.set('content-type', 'multipart/form-data; boundary=' + boundary);
                return blobBytes.get(new Blob(parts));
            }
            if (body instanceof Blob) { if (body.type && !headers.has('content-type')) headers.set('content-type', body.type); return rawBytes(blobBytes.get(body)).slice(); }
            if (body instanceof URLSearchParams) {
                if (!headers.has('content-type')) headers.set('content-type', 'application/x-www-form-urlencoded;charset=UTF-8');
                return new TextEncoder().encode(body.toString());
            }
            if (body instanceof ArrayBuffer || ArrayBuffer.isView(body)) return rawBytes(body).slice();
            if (!headers.has('content-type')) headers.set('content-type', 'text/plain;charset=UTF-8');
            return new TextEncoder().encode(String(body));
        }
        function retainBody(bytes) {
            const slot = { bytes: bytes ?? new Uint8Array(), used: false, evicted: false, locked: false, nullBody: bytes === null, offset: 0, streamEnded: false };
            retainedBodyBytes += slot.bytes.length; bodySlots.push(slot);
            evictResponseBodies();
            return slot;
        }
        function evictResponseBodies() {
            while (retainedBodyBytes > maxResponseBodies || bodySlots.length > 64) {
                const old = bodySlots.shift(); retainedBodyBytes -= old.bytes?.length ?? 0; old.bytes = null; old.evicted = true;
            }
        }
        function takeBody(slot) {
            if (slot.evicted) throw new TypeError('Response body was evicted by the memory limit.');
            if (slot.used) throw new TypeError('Body has already been consumed.');
            slot.used = true; slot.streamEnded = true;
            const bytes = slot.bytes; retainedBodyBytes -= bytes.length; slot.bytes = null; return bytes;
        }
        class Response {
            constructor(body = null, init = {}) {
                const status = init.status ?? 200;
                if (!Number.isInteger(status) || status < 200 || status > 599 || [204, 205, 304].includes(status) && body !== null) throw new RangeError('Invalid response status/body.');
                const headers = new Headers(init.headers);
                responseValues.set(this, { status, statusText: String(init.statusText || ''), headers, url: '', redirected: false, type: 'default', slot: retainBody(bodyBytes(body, headers)) });
            }
            get status() { return responseValues.get(this).status; }
            get statusText() { return responseValues.get(this).statusText; }
            get ok() { return this.status >= 200 && this.status <= 299; }
            get headers() { return responseValues.get(this).headers; }
            get url() { return responseValues.get(this).url; }
            get redirected() { return responseValues.get(this).redirected; }
            get type() { return responseValues.get(this).type; }
            get bodyUsed() { return responseValues.get(this).slot.used; }
            _consume(convert) { try {
                const slot = responseValues.get(this).slot;
                if (slot.locked) throw new TypeError('Body is locked by a reader.');
                return promiseResolve(convert(slot.nullBody ? new Uint8Array() : takeBody(slot)));
            } catch (error) { return promiseReject(error); } }
            text() { return this._consume(bytes => new TextDecoder().decode(bytes)); }
            json() { return this._consume(bytes => JSON.parse(new TextDecoder().decode(bytes))); }
            arrayBuffer() { return this._consume(bytes => bytes.buffer); }
            bytes() { return this._consume(bytes => bytes); }
            blob() { return this._consume(bytes => new Blob([bytes], { type: this.headers.get('content-type') || '' })); }
            formData() { return this._consume(bytes => {
                if (!(this.headers.get('content-type') || '').toLowerCase().startsWith('application/x-www-form-urlencoded')) throw new TypeError('Only URL-encoded response formData is supported.');
                const form = new FormData(); for (const [key, value] of new URLSearchParams(new TextDecoder().decode(bytes))) form.append(key, value); return form;
            }); }
            clone() {
                const state = responseValues.get(this);
                if (state.slot.used || state.slot.evicted || state.slot.locked) throw new TypeError('Body is unavailable.');
                const result = Object.create(Response.prototype), headers = new Headers(state.headers);
                headerValues.get(headers).immutable = headerValues.get(state.headers).immutable;
                responseValues.set(result, { ...state, headers, body: undefined, slot: retainBody(state.slot.nullBody ? null : state.slot.bytes.slice()) }); return result;
            }
            get body() {
                const state = responseValues.get(this), slot = state.slot;
                if (slot.nullBody) return null;
                if (state.body) return state.body;
                return state.body = { get locked() { return slot.locked; }, getReader() {
                    if (slot.locked) throw new TypeError('Body is already locked by a reader.');
                    slot.locked = true;
                    let released = false;
                    return { read() { try {
                        if (released) throw new TypeError('Reader lock has been released.');
                        if (slot.evicted) throw new TypeError('Response body was evicted by the memory limit.');
                        if (slot.streamEnded) return promiseResolve({ done: true, value: undefined });
                        slot.used = true;
                        if (slot.offset === slot.bytes.length) {
                            retainedBodyBytes -= slot.bytes.length; slot.bytes = null; slot.streamEnded = true;
                            return promiseResolve({ done: true, value: undefined });
                        }
                        const value = slot.bytes.slice(slot.offset, slot.offset + 65536); slot.offset += value.length;
                        if (slot.offset === slot.bytes.length) {
                            retainedBodyBytes -= slot.bytes.length; slot.bytes = null; slot.streamEnded = true;
                        }
                        return promiseResolve({ done: false, value });
                    } catch (error) { return promiseReject(error); } }, cancel() {
                        if (released) return promiseReject(new TypeError('Reader lock has been released.'));
                        retainedBodyBytes -= slot.bytes?.length ?? 0; slot.bytes = null; slot.used = true; slot.streamEnded = true;
                        return promiseResolve();
                    }, releaseLock() { released = true; slot.locked = false; } };
                } };
            }
            static json(value, init = {}) { const headers = new Headers(init.headers); if (!headers.has('content-type')) headers.set('content-type', 'application/json'); return new Response(JSON.stringify(value), { ...init, headers }); }
            static error() { return networkResponse({ status: 0, statusText: '', headers: [], url: '', type: 'error', redirected: false, bodyBase64: '' }); }
            static redirect(url, status = 302) { if (![301,302,303,307,308].includes(status)) throw new RangeError('Invalid redirect status.'); return new Response(null, { status, headers: { location: String(url) } }); }
        }
        function networkResponse(reply) {
            const result = Object.create(Response.prototype), headers = new Headers(reply.headers);
            headerValues.get(headers).immutable = true;
            responseValues.set(result, { status: reply.status, statusText: reply.statusText, headers, url: reply.url, redirected: reply.redirected, type: reply.type || 'basic', slot: retainBody(reply.nullBody || [0,204,205,304].includes(reply.status) ? null : decode64(reply.bodyBase64 || '')) });
            return result;
        }
        function abortError() { const error = new Error('The operation was aborted.'); error.name = 'AbortError'; return error; }
        class AbortSignal {
            constructor() { this.aborted = false; this.reason = undefined; this._listeners = new Set(); this.onabort = null; }
            addEventListener(name, listener) { if (name === 'abort') this._listeners.add(listener); }
            removeEventListener(name, listener) { if (name === 'abort') this._listeners.delete(listener); }
            throwIfAborted() { if (this.aborted) throw this.reason; }
            static abort(reason = abortError()) { const signal = new AbortSignal(); signal.aborted = true; signal.reason = reason; return signal; }
        }
        class AbortController {
            constructor() { Object.defineProperty(this, 'signal', { value: new AbortSignal() }); }
            abort(reason = abortError()) {
                const signal = this.signal; if (signal.aborted) return; signal.aborted = true; signal.reason = reason;
                const event = { type: 'abort', target: signal };
                for (const listener of signal._listeners) listener.call(signal, event);
                if (typeof signal.onabort === 'function') signal.onabort(event);
            }
        }
        class Request {
            constructor(input, init = {}) {
                const source = input instanceof Request ? input : null;
                if (source?.bodyUsed) throw new TypeError('Request body has already been consumed.');
                this.url = source ? source.url : String(input);
                if (!/^https?:\/\//i.test(this.url)) throw new TypeError('fetch requires an absolute HTTP or HTTPS URL.');
                this.method = String(init.method ?? source?.method ?? 'GET').toUpperCase();
                if (!/^[!#$%&'*+.^_`|~0-9A-Z-]+$/.test(this.method) || ['CONNECT','TRACE','TRACK'].includes(this.method)) throw new TypeError('Unsupported HTTP method.');
                this.headers = new Headers(init.headers ?? source?.headers);
                this.redirect = init.redirect ?? source?.redirect ?? 'follow';
                if (!['follow','error','manual'].includes(this.redirect)) throw new TypeError('Invalid redirect mode.');
                this.credentials = init.credentials ?? source?.credentials ?? 'same-origin';
                if (!['omit','same-origin'].includes(this.credentials)) throw new TypeError('Ambient credentials and cookies are unavailable.');
                this.mode = init.mode ?? source?.mode ?? 'cors';
                if (this.mode !== 'cors') throw new TypeError('Browser origin modes are unsupported.');
                if (init.integrity || init.keepalive || init.cache && !['default','no-store','no-cache','reload'].includes(init.cache)) throw new TypeError('This fetch option is unsupported.');
                this.signal = init.signal ?? source?.signal ?? new AbortSignal();
                if (!(this.signal instanceof AbortSignal)) throw new TypeError('Expected an AbortSignal.');
                this._bytes = Object.prototype.hasOwnProperty.call(init, 'body') ? bodyBytes(init.body, this.headers) : source?._bytes?.slice() ?? null;
                if (['GET','HEAD'].includes(this.method) && this._bytes !== null) throw new TypeError('GET and HEAD cannot have a body.');
                this.bodyUsed = false;
            }
            clone() { return new Request(this); }
            _consume(convert) { try { if (this.bodyUsed) throw new TypeError('Body has already been consumed.'); this.bodyUsed = this._bytes !== null; const bytes = this._bytes || new Uint8Array(); this._bytes = null; return promiseResolve(convert(bytes)); } catch (error) { return promiseReject(error); } }
            text() { return this._consume(bytes => new TextDecoder().decode(bytes)); }
            json() { return this._consume(bytes => JSON.parse(new TextDecoder().decode(bytes))); }
            arrayBuffer() { return this._consume(bytes => bytes.buffer); }
            blob() { return this._consume(bytes => new Blob([bytes], { type: this.headers.get('content-type') || '' })); }
            formData() { return this._consume(bytes => {
                if (!(this.headers.get('content-type') || '').toLowerCase().startsWith('application/x-www-form-urlencoded')) throw new TypeError('Only URL-encoded request formData is supported.');
                const form = new FormData(); for (const [key, value] of new URLSearchParams(new TextDecoder().decode(bytes))) form.append(key, value); return form;
            }); }
        }
        function fetch(input, init) {
            if (!fetchAllowed || !activeModel || !actionProgress) throw new TypeError('fetch is only available in workbook Actions.');
            try {
                const request = new Request(input, init);
                request.signal.throwIfAborted();
                if (++fetchSequence > 16) throw new TypeError('At most 16 fetch requests are allowed per action.');
                const id = fetchSequence;
                const wire = { id, url: request.url, method: request.method, headers: [...request.headers], bodyBase64: request._bytes === null ? null : encode64(request._bytes), redirect: request.redirect, credentials: request.credentials };
                request._bytes = null;
                if (input instanceof Request && init?.body === undefined && input._bytes !== null) { input.bodyUsed = true; input._bytes = null; }
                const signal = request.signal, method = request.method;
                const result = new NativePromise((resolve, reject) => {
                    const abort = () => {
                        signal.removeEventListener('abort', abort); fetchWaiters.delete(id);
                        actionProgress.requests = actionProgress.requests.filter(item => item.id !== id); reject(signal.reason);
                    };
                    fetchWaiters.set(id, { resolve, reject, signal, abort, method }); signal.addEventListener('abort', abort);
                });
                actionProgress.requests.push(wire); return result;
            } catch (error) { return promiseReject(error); }
        }
        function settleFetch(id, reply) {
            if (id < 0) return settleFileRequest(id, reply);
            const waiter = fetchWaiters.get(id);
            if (!waiter || !actionProgress) throw new TypeError('Unknown fetch response.');
            fetchWaiters.delete(id); waiter.signal.removeEventListener('abort', waiter.abort);
            actionProgress.requests = actionProgress.requests.filter(item => item.id !== id);
            if (reply.ok) { try { waiter.resolve(networkResponse({ ...reply.value, nullBody: waiter.method === 'HEAD' })); } catch (error) { waiter.reject(error); } }
            else waiter.reject(new TypeError(reply.error || 'Network request failed.'));
            return actionProgress;
        }
        function runAsyncAction(name, request) {
            if (!actionMode || !loaded || actionInvoked || request.event || name === '__console__') throw new ScriptError('InvalidReference', 'Asynchronous Actions are unavailable.');
            if (automatic && name !== '@initialize' && !initialized) throw new ScriptError('InvalidReference', '请先应用脚本，完成主入口初始化。');
            configureNetwork(request);
            actionInvoked = true; activeModel = true; hostWriteFailed = false; hostWrote = false; fileFailed = false;
            fetchAllowed = request.allowNetwork === true && name !== '@initialize'; fetchSequence = 0; fetchWaiters.clear();
            actionProgress = { pending: true, requests: [], patch: null, error: null };
            const progress = actionProgress;
            startUserFiles(automatic && request.allowUserFiles === true && name !== '@initialize');
            const cleanup = () => { activeModel = false; fetchAllowed = false; endUserFiles(); releaseFileContext(); };
            const fail = error => {
                let message = 'Action Promise rejected.';
                try {
                    if (typeof error === 'string') message = error;
                    else if (error && (typeof error === 'object' || typeof error === 'function')) {
                        // Inspect only data properties: rejection diagnostics must
                        // not execute user message/stack/toString accessors.
                        const detail = Object.getOwnPropertyDescriptor(error, 'message');
                        const stack = Object.getOwnPropertyDescriptor(error, 'stack');
                        if (detail && typeof detail.value === 'string') message = detail.value;
                        if (stack && typeof stack.value === 'string') message += '\n' + stack.value;
                    } else message = String(error);
                } catch { }
                progress.error = message.slice(0, 4000); progress.pending = false; cleanup();
            };
            try {
                let promise;
                if (name === '@initialize') {
                    if (initialized) throw new ScriptError('InvalidReference', '主入口在本次授权中已经运行。');
                    promise = promiseResolve();
                    for (const initialize of initializers) promise = promise.then(() => initialize(workbook, host));
                } else {
                    const main = mains[name];
                    if (typeof main !== 'function') throw new ScriptError('ItemNotFound', '找不到操作入口：' + name);
                    promise = promiseResolve(main(workbook, host));
                }
                promise.then(() => {
                    if (hostWriteFailed || fileFailed) throw new ScriptError('TransactionAborted', '有一次写入失败，本次变更未提交。');
                    if (fetchWaiters.size) throw new TypeError('The action must await or return every fetch request.');
                    if (fileWaiters.size) throw new TypeError('The action must await every file operation.');
                    if (name === '@initialize') initialized = true;
                    progress.patch = { operationId: request.operationId, baseRevision: request.snapshot.baseRevision, name, edits: [], fileEdits: runFiles ? runFiles.edits : [] };
                    progress.pending = false; cleanup();
                }).catch(fail);
            } catch (error) { fail(error); }
            return progress;
        }
        for (const [name, value] of Object.entries({ fetch, Headers, Request, Response, Blob, FormData, URLSearchParams, TextEncoder, TextDecoder, AbortController, AbortSignal }))
            Object.defineProperty(globalThis, name, { value, writable: false, configurable: false });
        """;
}
