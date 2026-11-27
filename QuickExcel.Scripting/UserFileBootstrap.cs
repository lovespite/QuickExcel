namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private const string UserFileBootstrap = """
        let fileRound = null, fileSequence = 0;
        const fileWaiters = new Map(), fileBlock = 65536, bufferLimit = 4 * 1048576;
        function fsError(code, message) { return new ScriptError(code, message); }
        function startUserFiles(enabled) {
            endUserFiles();
            if (enabled) fileRound = { buffers: 0, streams: new Set(), files: new WeakMap() };
            fileSequence = 0;
        }
        function endUserFiles() {
            if (fileRound) for (const clear of fileRound.streams) clear();
            fileWaiters.clear(); fileRound = null;
        }
        function requireFileRound(round) {
            if (!round || round !== fileRound || !activeModel)
                throw fsError('InvalidReference', 'File access requires the current user Action.');
        }
        function fileCall(round, fs, ...args) {
            requireFileRound(round);
            if (++fileSequence > 4096) throw fsError('BudgetExceeded', 'Too many file operations.');
            const id = -fileSequence;
            return new NativePromise((resolve, reject) => {
                fileWaiters.set(id, { resolve, reject });
                actionProgress.requests.push({ id, fs, args });
            });
        }
        function settleFileRequest(id, reply) {
            const waiter = fileWaiters.get(id);
            if (!waiter) throw new TypeError('Unknown file response.');
            fileWaiters.delete(id);
            actionProgress.requests = actionProgress.requests.filter(r => r.id !== id);
            if (reply.ok) waiter.resolve(reply.value);
            else waiter.reject(fsError(reply.code || 'FileError', reply.error || 'File operation failed.'));
            return actionProgress;
        }
        const userFileSystem = freeze({
            async getUserSelectedFiles(options) {
                const round = fileRound;
                return freeze((await fileCall(round, 'selectFiles', options)).map(e => makeEntry(round, e)));
            },
            async getUserSelectedDirectory(options = {}) {
                const multi = options.multiSelect === true, round = fileRound;
                const entries = (await fileCall(round, 'selectDirectory', options)).map(e => makeEntry(round, e));
                return multi ? freeze(entries) : entries[0] ?? null;
            }
        });
        function makeEntry(round, entry) {
            const call = (op, ...args) => fileCall(round, op, entry.id, ...args);
            const base = { type: entry.type, name: entry.name, stat: async () => freeze(await call('stat')), release: () => call('release') };
            if (entry.type === 'file') {
                const file = freeze({ ...base,
                    async openWorkbook() { return makeSelectedWorkbook(round, (await call('openWorkbook')).id); },
                    async open(options) {
                        const buffered = options === undefined;
                        if (!buffered && options?.buffering !== 'stream') throw fsError('InvalidArgument', 'Unknown buffering option.');
                        return makeFileStream(round, entry.access, await call('open', buffered), buffered);
                    }
                });
                round.files.set(file, entry.id);
                return file;
            }
            return freeze({ ...base,
                getFile: async (name, opts = {}) => makeEntry(round, await call('getFile', name, opts)),
                getDirectory: async (name, opts = {}) => makeEntry(round, await call('getDirectory', name, opts)),
                removeEntry: (name, opts = {}) => call('remove', name, opts.recursive ?? false),
                async *entries() {
                    const cursor = await call('entries');
                    try {
                        for (;;) {
                            const page = await fileCall(round, 'next', cursor);
                            for (const item of page.items) yield freeze(item);
                            if (page.done) return;
                        }
                    } finally { if (round === fileRound) await fileCall(round, 'endEntries', cursor); }
                }
            });
        }

        function saveWorkbookCopy(round, file, type, source, edits) {
            requireFileRound(round);
            const target = round.files.get(file);
            if (!target) throw fsError('InvalidReference', 'Expected a file handle from the current Action.');
            if (type !== 'xlsx' && type !== 'xlsq') throw fsError('InvalidArgument', 'Expected xlsx or xlsq.');
            // Pin host model inputs at invocation, before later synchronous JS edits can run.
            if (source === null) source = model('fs.call', { fs: 'beginSaveAs', args: [] });
            return fileCall(round, 'saveAs', target, type, source, edits);
        }
        async function saveActiveWorkbookAs(file, type) {
            if (hostWriteFailed || fileFailed) throw fsError('TransactionAborted', 'The workbook transaction has a failed write.');
            const edits = runFiles ? runFiles.edits.filter(edit => edit.path.startsWith('//xlsq/')) : [];
            await saveWorkbookCopy(fileRound, file, type, null, edits);
        }

        function makeSelectedWorkbook(round, id) {
            let closed = false, busy = false, failed = false;
            const clear = () => { closed = true; round.streams.delete(clear); };
            round.streams.add(clear);
            async function operation(method, args, write = false) {
                requireFileRound(round);
                if (closed) throw fsError('InvalidReference', 'Workbook closed.');
                if (busy) throw fsError('Locked', 'Another workbook operation is pending.');
                if (failed) throw fsError('TransactionAborted', 'Workbook draft was discarded.');
                busy = true;
                try {
                    if (method === 'saveAs') return await saveWorkbookCopy(round, args[0], args[1], id, []);
                    if (method === 'setValues' || method === 'setFormulas') {
                        // Snapshot the matrix before yielding. JSON would otherwise turn NaN or undefined into blanks.
                        const matrix = args[2];
                        if (!isArray(matrix)) throw fsError('InvalidArgument', 'Expected a cell matrix.');
                        let cells = 0;
                        args[2] = Array.from(matrix, row => {
                            if (!isArray(row)) throw fsError('InvalidArgument', 'Expected a cell row.');
                            if ((cells += row.length) > 20000) throw fsError('BudgetExceeded', 'Cell matrix is too large.');
                            return Array.from(row, value => {
                                if (method === 'setFormulas' ? typeof value !== 'string' :
                                    value !== null && typeof value !== 'boolean' && typeof value !== 'string' && !isFiniteNumber(value))
                                    throw fsError('InvalidArgument', 'Invalid cell value.');
                                return value;
                            });
                        });
                    }
                    return await fileCall(round, 'workbook', id, method, ...args);
                } catch (error) { if (write) failed = true; throw error; }
                finally { busy = false; }
            }
            async function finish(save) {
                if (closed) return;
                requireFileRound(round);
                if (busy) throw fsError('Locked', 'Another workbook operation is pending.');
                busy = true;
                try {
                    await fileCall(round, save && !failed ? 'close' : 'abort', id);
                    if (save && failed) throw fsError('TransactionAborted', 'Workbook draft was discarded.');
                } finally { clear(); }
            }
            return freeze({
                getWorksheetNames: () => operation('getWorksheetNames', []),
                addWorksheet: name => operation('addWorksheet', [name], true),
                getUsedRange: sheet => operation('getUsedRange', [sheet]),
                getValues: (sheet, address) => operation('getValues', [sheet, address]),
                getFormulas: (sheet, address) => operation('getFormulas', [sheet, address]),
                setValues: (sheet, address, values) => operation('setValues', [sheet, address, values], true),
                setFormulas: (sheet, address, formulas) => operation('setFormulas', [sheet, address, formulas], true),
                saveAs: (file, type) => operation('saveAs', [file, type]),
                close: () => finish(true), abort: () => finish(false)
            });
        }

        function fileEncoding(name) {
            if (name !== 'utf-8' && name !== 'utf-16le' && name !== 'utf-16be')
                throw fsError('InvalidArgument', 'Unsupported file encoding.');
            return name;
        }
        function fileText(bytes, encoding, pos) {
            const malformed = () => fsError('InvalidEncoding', 'Invalid file text encoding.');
            if (pos === 0) {
                const bom = bytes[0] === 255 && bytes[1] === 254 ? 'utf-16le' :
                    bytes[0] === 254 && bytes[1] === 255 ? 'utf-16be' :
                    bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191 ? 'utf-8' : null;
                if (bom && bom !== encoding) throw malformed();
                if (bom) bytes = bytes.subarray(encoding === 'utf-8' ? 3 : 2);
            }
            const chunks = [];
            if (encoding === 'utf-8') {
                const decoder = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true });
                for (let at = 0; at < bytes.length;) {
                    let end = Math.min(at + fileBlock, bytes.length);
                    while (end < bytes.length && end > at && (bytes[end] & 192) === 128) end--;
                    if (end === at) throw malformed();
                    try { chunks.push(decoder.decode(bytes.subarray(at, end))); } catch { throw malformed(); }
                    at = end;
                }
            } else {
                if (bytes.length % 2) throw malformed();
                const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
                for (let at = 0; at < bytes.length;) {
                    const units = new Uint16Array(Math.min(4096, (bytes.length - at) / 2));
                    for (let i = 0; i < units.length; i++, at += 2) units[i] = view.getUint16(at, encoding === 'utf-16le');
                    chunks.push(characters(units));
                }
            }
            const text = chunks.join('');
            if (/[\uD800-\uDFFF]/u.test(text)) throw malformed();
            return text;
        }
        function fileBytes(text, encoding, line) {
            fileEncoding(encoding);
            if (typeof text !== 'string') throw fsError('InvalidArgument', 'Expected text.');
            if (text.length > bufferLimit) throw fsError('BudgetExceeded', 'Text is too large.');
            if (/[\uD800-\uDFFF]/u.test(text)) throw fsError('InvalidEncoding', 'Unpaired surrogate.');
            if (line) text += '\n';
            if (encoding === 'utf-8') {
                const encoder = new TextEncoder(), chunks = [];
                let length = 0;
                for (let at = 0; at < text.length;) {
                    let end = Math.min(at + 8192, text.length);
                    if (end < text.length && (text.charCodeAt(end) & 0xfc00) === 0xdc00) end--;
                    const chunk = encoder.encode(text.slice(at, end));
                    length += chunk.length;
                    if (length > bufferLimit) throw fsError('BudgetExceeded', 'Text is too large.');
                    chunks.push(chunk); at = end;
                }
                const out = new Uint8Array(length);
                let at = 0; for (const chunk of chunks) { out.set(chunk, at); at += chunk.length; }
                return out;
            }
            if (text.length * 2 > bufferLimit) throw fsError('BudgetExceeded', 'Text is too large.');
            const out = new Uint8Array(text.length * 2), view = new DataView(out.buffer);
            for (let i = 0; i < text.length; i++) view.setUint16(i * 2, text.charCodeAt(i), encoding === 'utf-16le');
            return out;
        }
        function lineEnd(bytes, encoding, last, from = 0) {
            const step = encoding === 'utf-8' ? 1 : 2;
            const unit = i => step === 1 ? bytes[i] : encoding === 'utf-16le' ? bytes[i] | bytes[i+1] << 8 : bytes[i] << 8 | bytes[i+1];
            for (let i = from; i + step <= bytes.length; i += step) {
                const ch = unit(i);
                if (ch === 10) return [i, i + step];
                if (ch === 13) {
                    if (!last && i + 2 * step > bytes.length) return null;
                    return [i, i + (i + 2 * step <= bytes.length && unit(i + step) === 10 ? 2 : 1) * step];
                }
            }
            return last ? [bytes.length, bytes.length] : null;
        }
        async function makeFileStream(round, access, opened, buffered) {
            let position = 0, length = opened.length, closed = false, failed = false, busy = false;
            let data = null, dirty = [], cache = null, cacheAt = 0;
            const canRead = access.includes('r'), canWrite = access.includes('w');
            const call = (op, ...args) => fileCall(round, op, opened.id, ...args);
            function clear() {
                if (data) round.buffers -= data.length;
                data = null; cache = null; dirty = []; closed = true;
                round.streams.delete(clear);
            }
            round.streams.add(clear);
            function check(write = false) {
                requireFileRound(round);
                if (buffered && busy) throw fsError('Locked', 'Stream closing.');
                if (closed || failed) throw fsError('InvalidReference', 'The file stream is no longer usable.');
                if (write && !canWrite || !write && !canRead) throw fsError('PermissionDenied', 'File access denied.');
            }
            function resize(size) {
                if (size > bufferLimit) throw fsError('BudgetExceeded', 'Buffered file exceeds 4 MiB.');
                if (data && size <= data.length) return;
                const old = data?.length ?? 0, available = 2 * bufferLimit - round.buffers + old;
                if (size > available) throw fsError('BudgetExceeded', 'Action file buffers exceed 8 MiB.');
                const capacity = Math.min(available, bufferLimit, Math.max(size, old * 2));
                const next = new Uint8Array(capacity);
                if (data) next.set(data);
                round.buffers += capacity - old; data = next;
            }
            async function load(at, count) { return decode64(await call('read', at, count)); }
            if (buffered) {
                try {
                    resize(length);
                    if (canRead) for (let at = 0; at < length; at += fileBlock) data.set(await load(at, Math.min(length-at, fileBlock)), at);
                } catch (error) { clear(); await call('abort'); throw error; }
            }
            function range(view, offset, len) {
                if (!(view instanceof Uint8Array) || !Number.isSafeInteger(offset) || !Number.isSafeInteger(len) ||
                    offset < 0 || len < 0 || len > view.byteLength - offset)
                    throw fsError('InvalidArgument', 'Invalid Uint8Array range.');
                return view.subarray(offset, offset + len);
            }
            function seek(at) {
                requireFileRound(round);
                if (buffered && busy) throw fsError('Locked', 'Stream closing.');
                if (closed || failed) throw fsError('InvalidReference', 'Closed stream.');
                if (!Number.isSafeInteger(at) || at < 0 || at > length) throw fsError('InvalidArgument', 'Invalid byte position.');
                position = at;
            }
            function mark(start, end) {
                if (start === end) return;
                let i = 0;
                while (i < dirty.length && dirty[i][1] < start) i++;
                let j = i;
                while (j < dirty.length && dirty[j][0] <= end) {
                    start = Math.min(start, dirty[j][0]); end = Math.max(end, dirty[j][1]); j++;
                }
                dirty.splice(i, j - i, [start, end]);
                if (dirty.length > 4096) throw fsError('BudgetExceeded', 'Too many disjoint writes.');
            }
            function advance(n) { position += n; length = Math.max(length, position); return n; }
            function writeLocal(bytes) {
                resize(position + bytes.length); data.set(bytes, position);
                mark(position, position + bytes.length); return advance(bytes.length);
            }
            function textResult(bytes, encoding, ends, limit, line) {
                if (ends[0] > limit) throw fsError('BudgetExceeded', 'File text exceeds the read limit.');
                const value = fileText(bytes.subarray(0, ends[0]), encoding, position);
                advance(ends[1]);
                return line && value === '' && position === length && ends[0] === ends[1] ? null : value;
            }
            function localText(encoding, line) {
                check(); fileEncoding(encoding);
                if (line && position === length) return null;
                const block = data.subarray(position, length);
                return textResult(block, encoding, line ? lineEnd(block, encoding, true) : [block.length, block.length], line ? 1048576 : bufferLimit, line);
            }
            async function serial(fn) {
                if (busy) throw fsError('Locked', 'Another stream operation is pending.');
                busy = true;
                try { return await fn(); } finally { busy = false; }
            }
            async function upload(at, bytes) {
                for (let i = 0; i < bytes.length; i += fileBlock)
                    await call('write', at + i, encode64(bytes.subarray(i, i + fileBlock)));
            }
            async function remoteText(encoding, line) {
                check(); fileEncoding(encoding);
                if (line && position === length) return null;
                const limit = line ? 1048576 : bufferLimit, remaining = length - position;
                if (!line && remaining > limit) throw fsError('BudgetExceeded', 'File text exceeds the read limit.');
                let at = position, used = 0, accum = null, scan = 0;
                do {
                    if (!cache || at < cacheAt || at >= cacheAt + cache.length) {
                        cacheAt = at; cache = await load(at, Math.min(fileBlock, length-at));
                    }
                    const part = cache.subarray(at-cacheAt);
                    let bytes = part;
                    if (accum) { accum.set(part, used); bytes = accum.subarray(0, used + part.length); }
                    const last = at + part.length === length;
                    const ends = line ? lineEnd(bytes, encoding, last, scan) : last ? [bytes.length,bytes.length] : null;
                    if (ends) return textResult(bytes, encoding, ends, limit, line);
                    if (bytes.length > limit + 2) throw fsError('BudgetExceeded', 'File line exceeds the read limit.');
                    if (!accum) { accum = new Uint8Array(Math.min(remaining, limit + fileBlock + 4)); accum.set(part); }
                    used += part.length; at += part.length;
                    scan = encoding === 'utf-8' ? Math.max(0, used - 1) : Math.max(0, (used & ~1) - 2);
                } while (at < length);
            }
            async function close() {
                if (closed) return;
                return serial(async () => {
                    try {
                        requireFileRound(round);
                        if (failed) throw fsError('TransactionAborted', 'File draft discarded.');
                        for (const [start,end] of dirty) await upload(start, data.subarray(start,end));
                        await call('close');
                    } catch (error) { if (round === fileRound) await call('abort'); throw error; }
                    finally { clear(); }
                });
            }
            async function abort() {
                if (!closed) await serial(async () => { try { await call('abort'); } finally { clear(); } });
            }
            const result = { close, abort, get length() { return length; }, get position() { return position; } };
            function writing(fn) { check(true); try { return fn(); } catch (error) { failed = true; throw error; } }
            async function remoteWrite(produce) {
                return serial(async () => {
                    check(true);
                    try {
                        const bytes = produce();
                        if (bytes.length > 16 * 1048576) throw fsError('BudgetExceeded', 'Write exceeds the draft budget.');
                        const owned = bytes.slice();
                        await upload(position, owned); cache = null; return advance(owned.length);
                    } catch (error) { failed = true; throw error; }
                });
            }
            Object.assign(result, buffered ? {
                read(view, offset, len) {
                    check();
                    const dest = range(view, offset, len), count = Math.min(len, length-position);
                    dest.set(data.subarray(position,position+count)); return advance(count);
                },
                write(view, offset, len) { return writing(() => writeLocal(range(view, offset, len))); },
                seek, readLine: (enc='utf-8') => localText(enc,true), readAllText: (enc='utf-8') => localText(enc,false),
                writeText: (text,enc='utf-8') => writing(() => { writeLocal(fileBytes(text, enc, false)); }),
                writeLine: (text,enc='utf-8') => writing(() => { writeLocal(fileBytes(text, enc, true)); })
            } : {
                read(view, offset, len) { return serial(async () => {
                    check(); const dest = range(view,offset,len), count = Math.min(len,fileBlock,length-position);
                    const bytes = await load(position,count); dest.set(bytes); return advance(bytes.length);
                }); },
                write: (view, offset, len) => remoteWrite(() => range(view,offset,len)),
                seek: at => serial(() => seek(at)),
                readLine: (enc='utf-8') => serial(() => remoteText(enc,true)),
                readAllText: (enc='utf-8') => serial(() => remoteText(enc,false)),
                writeText: async (text,enc='utf-8') => { await remoteWrite(() => fileBytes(text,enc,false)); },
                writeLine: async (text,enc='utf-8') => { await remoteWrite(() => fileBytes(text,enc,true)); }
            });
            return freeze(result);
        }
        """;
}

