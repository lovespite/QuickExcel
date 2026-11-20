namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    // Capability-driven object model. Every implemented member is declared in the
    // generated manifest; local members are pure handle operations, host members
    // cross the transactional model bridge. Anything else fails with UnsupportedApi
    // before any workbook state is touched.
    private const string WorkbookModelBootstrap = """
        const implemented = capabilities.types;
        const QuickExcel = deepFreeze(capabilities.enums || {});
        Object.defineProperty(globalThis, 'QuickExcel', { value: QuickExcel });
        function apiArgument(value, depth = 0) {
            if (depth > 16) throw new ScriptError('InvalidArgument', '参数嵌套超过限制或存在循环。');
            if (value && typeof value === 'object' && handleData.has(value)) return dataOf(value);
            if (isArray(value)) return value.map(item => apiArgument(item, depth + 1));
            if (value && typeof value === 'object') {
                if (value instanceof Date) throw new ScriptError('UnsupportedApi', '日期验证请传序列值公式；Date 参数尚未支持。');
                const result = Object.create(null);
                for (const key of ownKeys(value)) result[key] = apiArgument(value[key], depth + 1);
                return result;
            }
            return value;
        }
        function apiResult(value) {
            if (isArray(value)) return value.map(apiResult);
            if (value && value.scriptHandle === true && value.type === 'RangeAreas') return makeHandle('RangeAreas', { sheetId: value.data.sheetId, addresses: value.data.addresses.map(modelAddress) });
            if (value && typeof value === 'object' && value.scriptHandle === true)
                return value.type === 'Range' ? range(value.data.sheetId, value.data.address) : makeHandle(value.type, value.data);
            return value === null ? undefined : value;
        }
        const prototypes = Object.create(null);
        const handleData = new WeakMap();
        const actionEntries = [];

        function unsupported(type, member) {
            return new ScriptError('UnsupportedApi', type + '.' + String(member) + ' 尚未实现。', type + '.' + String(member));
        }
        function dataOf(value) {
            const data = handleData.get(value);
            if (!data) throw new ScriptError('InvalidReference', '工作簿对象句柄已失效。');
            return data;
        }
        function model(method, ...args) {
            if (!activeModel) throw new ScriptError('InvalidReference', '工作簿对象只能在操作、事件或控制台请求中使用。');
            const reply = nativeModelCall(JSON.stringify({ method, args }));
            if (!reply || reply.ok !== true) {
                const code = reply && typeof reply.code === 'string' ? reply.code : 'TransactionAborted';
                if (method.startsWith('set') || method === 'clear' || code === 'BudgetExceeded' || code === 'Locked' ||
                    code === 'InvalidArgument' && hostWrote) hostWriteFailed = true;
                throw new ScriptError(code, reply && typeof reply.error === 'string' ? reply.error : '工作簿请求失败。');
            }
            if (method.startsWith('set') || method === 'clear') hostWrote = true;
            return reply.value;
        }
        function modelAddress(address) {
            if (typeof address !== 'string' || address.length > 40 ||
                !/^[A-Za-z]+[1-9][0-9]*(?::[A-Za-z]+[1-9][0-9]*)?$/.test(address))
                throw new ScriptError('InvalidArgument', '需要 A1 单元格或矩形 A1:B2 区域。');
            const parts = address.toUpperCase().split(':');
            const normalized = parts.length === 2 && parts[0] === parts[1] ? parts[0] : parts.join(':');
            rangeSize(normalized); return normalized;
        }
        function rangeSize(address) {
            const parts = address.split(':');
            function point(text) {
                const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(text);
                if (!match) throw new ScriptError('InvalidArgument', 'A1 地址无效。');
                let column = 0;
                for (const c of match[1]) column = column * 26 + c.charCodeAt(0) - 64;
                const row = Number(match[2]);
                if (!Number.isSafeInteger(row) || row > capabilities.limits.rows || column > capabilities.limits.columns)
                    throw new ScriptError('InvalidArgument', 'A1 地址超出范围。');
                return [row, column];
            }
            const a = point(parts[0]), b = point(parts[1] || parts[0]);
            if (b[0] < a[0] || b[1] < a[1]) throw new ScriptError('InvalidArgument', '区域起止顺序颠倒。');
            return { rows: b[0] - a[0] + 1, columns: b[1] - a[1] + 1 };
        }
        function rangeRect(address) {
            const parts = address.split(':');
            function point(text) {
                const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(text);
                if (!match) throw new ScriptError('InvalidArgument', 'A1 地址无效。');
                let column = 0;
                for (const c of match[1]) column = column * 26 + c.charCodeAt(0) - 64;
                return { row: Number(match[2]) - 1, column: column - 1 };
            }
            const a = point(parts[0]), b = point(parts[1] || parts[0]);
            return { top: a.row, left: a.column, bottom: b.row, right: b.column };
        }
        function rangeText(top, left, bottom, right) {
            if (!Number.isInteger(top) || !Number.isInteger(left) || !Number.isInteger(bottom) || !Number.isInteger(right) ||
                top < 0 || left < 0 || bottom >= capabilities.limits.rows || right >= capabilities.limits.columns || top > bottom || left > right)
                throw new ScriptError('InvalidArgument', '区域超出工作表范围。');
            function name(value) {
                let n = value + 1, text = '';
                while (n > 0) { n--; text = String.fromCharCode(65 + n % 26) + text; n = Math.floor(n / 26); }
                return text;
            }
            const first = name(left) + (top + 1), last = name(right) + (bottom + 1);
            return first === last ? first : first + ':' + last;
        }
        function rangeFrom(value, sheetId) {
            if (value && typeof value === 'object' && handleData.has(value)) {
                const data = dataOf(value);
                if (typeof data.address !== 'string') throw new ScriptError('InvalidArgument', '需要 Range。');
                if (sheetId !== undefined && data.sheetId !== sheetId)
                    throw new ScriptError('InvalidArgument', '区域必须属于同一工作表。');
                return { sheetId: data.sheetId, address: data.address, rect: rangeRect(data.address) };
            }
            if (typeof value !== 'string') throw new ScriptError('InvalidArgument', '需要 Range 或 A1 地址。');
            const address = modelAddress(value);
            return { sheetId, address, rect: rangeRect(address) };
        }
        function checkIndex(value, name) {
            if (!Number.isInteger(value) || value < 0) throw new ScriptError('InvalidArgument', name + ' 必须是非负整数。');
            return value;
        }
        function checkCount(value, name) {
            if (!Number.isInteger(value) || value === 0) throw new ScriptError('InvalidArgument', name + ' 必须是非零整数。');
            return value;
        }
        function checkScalar(value) {
            if (typeof value === 'string' && value.length <= 32767) return value;
            if (typeof value === 'number' && isFiniteNumber(value)) return value;
            if (typeof value === 'boolean') return value;
            throw new ScriptError('InvalidArgument', '单元格值必须是字符串、有限数值或布尔值。');
        }
        function checkMatrix(values, address, check) {
            if (!isArray(values)) throw new ScriptError('InvalidArgument', '需要二维数组。');
            const size = rangeSize(address);
            if (values.length !== size.rows) throw new ScriptError('InvalidArgument', '数据行数与区域不一致。');
            let total = 0;
            const rows = [];
            for (const row of values) {
                if (!isArray(row) || row.length !== size.columns)
                    throw new ScriptError('InvalidArgument', '数据列数与区域不一致。');
                total += row.length;
                if (total > 20000) throw new ScriptError('InvalidArgument', '一次写入最多 20000 个单元格。');
                rows.push(row.map(check));
            }
            return rows;
        }
        function checkFormula(formula) {
            if (typeof formula !== 'string' || formula.length < 2 || formula.length > 8192 || formula[0] !== '=')
                throw new ScriptError('InvalidArgument', '公式必须以 = 开头且不超过 8192 个字符。');
            return formula;
        }
        function checkColor(color) {
            if (typeof color !== 'string' || !/^#[0-9a-fA-F]{6}$/.test(color))
                throw new ScriptError('InvalidArgument', '颜色必须是 #RRGGBB。');
            return color;
        }

        function makeHandle(type, data) {
            const target = Object.create(prototypeFor(type));
            const proxy = new Proxy(target, {
                get(value, key) {
                    if (typeof key === 'symbol') {
                        if (key === Symbol.toPrimitive || key === Symbol.iterator || key === Symbol.asyncIterator) return undefined;
                        if (key === Symbol.toStringTag) return type;
                        return value[key];
                    }
                    if (key in value) return value[key];
                    if (key === 'then' || key === 'toJSON') return undefined;
                    throw unsupported(type, key);
                }
            });
            const frozen = freeze(data);
            handleData.set(target, frozen);
            handleData.set(proxy, frozen);
            return freeze(proxy);
        }
        function worksheet(sheetId) { return makeHandle('Worksheet', { sheetId }); }
        function range(sheetId, address) { return makeHandle('Range', { sheetId, address: modelAddress(address) }); }

        const localMembers = {
            Worksheet: {
                getId() { return dataOf(this).sheetId; },
                getCell(address, column) {
                    const data = dataOf(this);
                    let text;
                    if (typeof address === 'number') {
                        if (!Number.isInteger(address) || address < 0 || address >= capabilities.limits.rows ||
                            !Number.isInteger(column) || column < 0 || column >= capabilities.limits.columns)
                            throw new ScriptError('InvalidArgument', '行列索引必须是零基整数。');
                        let name = '', n = column + 1;
                        while (n > 0) { n--; name = String.fromCharCode(65 + n % 26) + name; n = Math.floor(n / 26); }
                        text = name + (address + 1);
                    } else {
                        text = modelAddress(address);
                        if (text.indexOf(':') >= 0) throw new ScriptError('InvalidArgument', 'getCell 只接受单个单元格。');
                    }
                    return range(data.sheetId, text);
                },
                getRange(address) { const data = dataOf(this); return range(data.sheetId, modelAddress(address)); },
                getRangeByIndexes(startRow, startColumn, rowCount, columnCount) {
                    const data = dataOf(this);
                    checkIndex(startRow, 'startRow'); checkIndex(startColumn, 'startColumn');
                    if (!Number.isInteger(rowCount) || rowCount < 1 || !Number.isInteger(columnCount) || columnCount < 1)
                        throw new ScriptError('InvalidArgument', '行数和列数必须是正整数。');
                    return range(data.sheetId, rangeText(startRow, startColumn, startRow + rowCount - 1, startColumn + columnCount - 1));
                },
                getRanges(address) {
                    const data = dataOf(this);
                    if (typeof address !== 'string' || address.trim().length === 0)
                        throw new ScriptError('InvalidArgument', 'getRanges 需要逗号分隔的 A1 地址。');
                    const addresses = address.split(',').map(part => modelAddress(part.trim()));
                    if (addresses.length > 64) throw new ScriptError('InvalidArgument', '一次最多返回 64 个区域。');
                    return makeHandle('RangeAreas', { sheetId: data.sheetId, addresses });
                },
                getFreezePanes() { return makeHandle('WorksheetFreezePanes', dataOf(this)); }
            },
            Range: {
                getAddress() { return dataOf(this).address; },
                getAddressLocal() { return dataOf(this).address; },
                getWorksheet() { return worksheet(dataOf(this).sheetId); },
                getFormat() { return makeHandle('RangeFormat', dataOf(this)); },
                getCell(row, column) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    checkIndex(row, 'row'); checkIndex(column, 'column');
                    const r = rect.top + row, c = rect.left + column;
                    return range(data.sheetId, rangeText(r, c, r, c));
                },
                getRow(row) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    checkIndex(row, 'row');
                    if (rect.top + row > rect.bottom) throw new ScriptError('InvalidArgument', '行索引超出区域。');
                    return range(data.sheetId, rangeText(rect.top + row, rect.left, rect.top + row, rect.right));
                },
                getColumn(column) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    checkIndex(column, 'column');
                    if (rect.left + column > rect.right) throw new ScriptError('InvalidArgument', '列索引超出区域。');
                    return range(data.sheetId, rangeText(rect.top, rect.left + column, rect.bottom, rect.left + column));
                },
                getEntireRow() {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    return range(data.sheetId, rangeText(rect.top, 0, rect.bottom, capabilities.limits.columns - 1));
                },
                getEntireColumn() {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    return range(data.sheetId, rangeText(0, rect.left, capabilities.limits.rows - 1, rect.right));
                },
                getOffsetRange(rowOffset, columnOffset) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    if (!Number.isInteger(rowOffset) || !Number.isInteger(columnOffset))
                        throw new ScriptError('InvalidArgument', '偏移量必须是整数。');
                    return range(data.sheetId, rangeText(rect.top + rowOffset, rect.left + columnOffset,
                        rect.bottom + rowOffset, rect.right + columnOffset));
                },
                getResizedRange(deltaRows, deltaColumns) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    if (!Number.isInteger(deltaRows) || !Number.isInteger(deltaColumns))
                        throw new ScriptError('InvalidArgument', '调整量必须是整数。');
                    return range(data.sheetId, rangeText(rect.top, rect.left, rect.bottom + deltaRows, rect.right + deltaColumns));
                },
                getAbsoluteResizedRange(numRows, numColumns) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    if (!Number.isInteger(numRows) || numRows < 1 || !Number.isInteger(numColumns) || numColumns < 1)
                        throw new ScriptError('InvalidArgument', '行数和列数必须是正整数。');
                    return range(data.sheetId, rangeText(rect.top, rect.left, rect.top + numRows - 1, rect.left + numColumns - 1));
                },
                getBoundingRect(value) {
                    const data = dataOf(this), rect = rangeRect(data.address), other = rangeFrom(value, data.sheetId).rect;
                    return range(data.sheetId, rangeText(Math.min(rect.top, other.top), Math.min(rect.left, other.left),
                        Math.max(rect.bottom, other.bottom), Math.max(rect.right, other.right)));
                },
                getIntersection(value) {
                    const data = dataOf(this), rect = rangeRect(data.address), other = rangeFrom(value, data.sheetId).rect;
                    const top = Math.max(rect.top, other.top), left = Math.max(rect.left, other.left);
                    const bottom = Math.min(rect.bottom, other.bottom), right = Math.min(rect.right, other.right);
                    if (top > bottom || left > right) throw new ScriptError('InvalidArgument', '区域没有交集。');
                    return range(data.sheetId, rangeText(top, left, bottom, right));
                },
                getLastCell() {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    return range(data.sheetId, rangeText(rect.bottom, rect.right, rect.bottom, rect.right));
                },
                getLastRow() {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    return range(data.sheetId, rangeText(rect.bottom, rect.left, rect.bottom, rect.right));
                },
                getLastColumn() {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    return range(data.sheetId, rangeText(rect.top, rect.right, rect.bottom, rect.right));
                },
                getRowsAbove(count) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    count = count === undefined ? 1 : checkCount(count, 'count');
                    return count > 0
                        ? range(data.sheetId, rangeText(rect.top - count, rect.left, rect.top - 1, rect.right))
                        : range(data.sheetId, rangeText(rect.top, rect.left, rect.top - count - 1, rect.right));
                },
                getRowsBelow(count) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    count = count === undefined ? 1 : checkCount(count, 'count');
                    return count > 0
                        ? range(data.sheetId, rangeText(rect.bottom + 1, rect.left, rect.bottom + count, rect.right))
                        : range(data.sheetId, rangeText(rect.bottom + count + 1, rect.left, rect.bottom, rect.right));
                },
                getColumnsBefore(count) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    count = count === undefined ? 1 : checkCount(count, 'count');
                    return count > 0
                        ? range(data.sheetId, rangeText(rect.top, rect.left - count, rect.bottom, rect.left - 1))
                        : range(data.sheetId, rangeText(rect.top, rect.left, rect.bottom, rect.left - count - 1));
                },
                getColumnsAfter(count) {
                    const data = dataOf(this), rect = rangeRect(data.address);
                    count = count === undefined ? 1 : checkCount(count, 'count');
                    return count > 0
                        ? range(data.sheetId, rangeText(rect.top, rect.right + 1, rect.bottom, rect.right + count))
                        : range(data.sheetId, rangeText(rect.top, rect.right + count + 1, rect.bottom, rect.right));
                }
            },
            NamedItem: {
                getScope() {
                    this.getName();
                    return dataOf(this).scopeSheetId ? QuickExcel.NamedItemScope.worksheet : QuickExcel.NamedItemScope.workbook;
                },
                getWorksheet() {
                    this.getName();
                    const data = dataOf(this);
                    return data.scopeSheetId ? worksheet(data.scopeSheetId) : undefined;
                }
            },
            RemoveDuplicatesResult: {
                getRemoved() { return dataOf(this).removed; },
                getUniqueRemaining() { return dataOf(this).uniqueRemaining; }
            },
            RangeFormat: {
                getFill() { return makeHandle('RangeFill', dataOf(this)); }
            },
            RangeAreas: {
                getAddress() { return dataOf(this).addresses.join(','); },
                getAreaCount() { return dataOf(this).addresses.length; },
                getCellCount() {
                    const data = dataOf(this);
                    return data.addresses.reduce((total, address) => {
                        const size = rangeSize(address);
                        return total + size.rows * size.columns;
                    }, 0);
                },
                getAreas() { const data = dataOf(this); return data.addresses.map(address => range(data.sheetId, address)); },
                getWorksheet() { return worksheet(dataOf(this).sheetId); }
            }
        };

        const resultAdapters = {
            'Workbook.getActiveWorksheet': value => worksheet(value.sheetId),
            'Workbook.getActiveCell': value => range(value.sheetId, value.cell),
            'Workbook.getSelectedRange': value => {
                if (value.ranges.length !== 1)
                    throw new ScriptError('InvalidArgument', '当前选区包含多个区域；请使用 getSelectedRanges()。');
                return range(value.sheetId, value.ranges[0]);
            },
            'Workbook.getSelectedRanges': value => makeHandle('RangeAreas', { sheetId: value.sheetId, addresses: value.ranges }),
            'Workbook.getWorksheet': value => value ? worksheet(value.id) : undefined,
            'Workbook.getWorksheets': value => value.map(sheet => worksheet(sheet.id)),
            'Worksheet.getName': value => value.name
        };

        function contextOf(type, data) {
            if (type === 'Workbook') return [];
            if (type === 'Worksheet') return [data.sheetId];
            return [data.sheetId, data.address];
        }
        function hostMember(type, name, spec) {
            return function (...args) {
                try {
                    const data = dataOf(this);
                    if (spec.host.startsWith('getApi:') || spec.host.startsWith('setApi:')) {
                        const result = model(spec.host, data, args.map(value => apiArgument(value)));
                        if (name === 'getCreationDate' && (type === 'Comment' || type === 'CommentReply')) return new Date(result);
                        return result && typeof result === 'object' ? apiResult(result) : result === null &&
                            ['getTable','getColumn','getColumnById','getColumnByName','getUsedRange','getChart','getComment','getCommentReply','getConditionalFormat','getCustom','getCellValue','getTextComparison','getPreset',
                             'getNext','getPrevious','getFirstWorksheet','getLastWorksheet','getNamedItem','getLocation','getHyperlink','find'].includes(name) ? undefined : result;
                    }
                    if (type === 'Range') {
                        if (name === 'setValue') args = [checkScalar(args[0])];
                        else if (name === 'setValues') args = [checkMatrix(args[0], data.address, checkScalar)];
                        else if (name === 'setFormula') args = [checkFormula(args[0])];
                        else if (name === 'setFormulas') args = [checkMatrix(args[0], data.address, checkFormula)];
                    } else if (type === 'RangeFill') {
                        if (name === 'setColor') args = [checkColor(args[0])];
                        else if (name === 'clear') args = [null];
                    }
                    const context = contextOf(type, data);
                    if (args.length + context.length !== spec.arity)
                        throw new ScriptError('InvalidArgument', type + '.' + name + ' 的参数个数不正确。');
                    const value = model(spec.host, ...context, ...args);
                    const adapt = resultAdapters[type + '.' + name];
                    return adapt ? adapt(value) : value;
                } catch (error) {
                    // A rejected argument after a successful host write still
                    // aborts the transaction, even when the script catches it.
                    if (hostWrote && error instanceof ScriptError && error.code === 'InvalidArgument') hostWriteFailed = true;
                    throw error;
                }
            };
        }
        function prototypeFor(type) {
            const existing = prototypes[type];
            if (existing) return existing;
            const definition = implemented[type];
            if (!definition || typeof definition.members !== 'object')
                throw new ScriptError('InvalidReference', '未知的工作簿对象类型：' + type);
            const proto = Object.create(Object.prototype);
            for (const name of ownKeys(definition.members)) {
                const spec = definition.members[name];
                const implementation = spec.kind === 'local'
                    ? localMembers[type] && localMembers[type][name]
                    : spec.kind === 'host' ? hostMember(type, name, spec) : null;
                if (typeof implementation !== 'function')
                    throw new ScriptError('InvalidReference', '能力清单成员缺少实现：' + type + '.' + name);
                Object.defineProperty(proto, name, { value: implementation, writable: false, enumerable: false, configurable: false });
            }
            return prototypes[type] = freeze(proto);
        }

        function registerFunction(name, fn) {
            // Unbound helper modules are shared between contexts; registration is
            // simply inert when the current context does not declare UDF support.
            if (actionMode || sharedMode && !(registrationRole() & 1)) return;
            if (loaded) throw new ScriptError('InvalidArgument', '函数只能在脚本加载期间注册。');
            let descriptor = { apiVersion: 1, minArguments: 0, maxArguments: 255, returnType: 'any', profile: 'conservative', parameters: [] };
            if (typeof name === 'object' && name !== null) {
                const input = name;
                if (input.apiVersion !== 2 || typeof input.name !== 'string')
                    throw new ScriptError('InvalidArgument', 'v2 描述符需要 apiVersion: 2 和 name。');
                name = upper(input.name);
                if (!validV2Name(name)) throw new ScriptError('InvalidArgument', 'v2 函数必须使用 JS 命名空间。');
                const parameters = input.parameters ?? [];
                if (!isArray(parameters) || parameters.length > 255)
                    throw new ScriptError('InvalidArgument', '无效的参数描述。');
                const types = ['any', 'number', 'text', 'boolean', 'blank', 'range'];
                for (const p of parameters)
                    if (!p || typeof p.name !== 'string' || !types.includes(p.type ?? 'any'))
                        throw new ScriptError('InvalidArgument', '无效的参数类型。');
                descriptor = { name, apiVersion: 2, minArguments: parameters.length, maxArguments: parameters.length,
                    returnType: input.returns ?? 'any', profile: input.profile ?? 'conservative',
                    parameters: parameters.map(p => ({ name: p.name, type: p.type ?? 'any' })) };
                if (!['any', 'number', 'text', 'boolean', 'blank'].includes(descriptor.returnType) || descriptor.profile !== 'conservative')
                    throw new ScriptError('InvalidArgument', 'M3 只支持标量返回和保守执行。');
            } else {
                if (typeof name !== 'string') throw new ScriptError('InvalidArgument', '函数名必须是字符串。');
                name = upper(name);
                if (!validName(name)) throw new ScriptError('InvalidArgument', 'v1 函数名只能包含 A-Z。');
            }
            descriptor.name = name;
            if (typeof fn !== 'function') throw new ScriptError('InvalidArgument', '注册的公式必须是函数。');
            if (registry[name]) throw new ScriptError('InvalidArgument', '函数重复注册：' + name);
            registry[name] = fn;
            descriptors[name] = deepFreeze(descriptor);
        }

        let runFiles = null;
        let fileFailed = false;
        function fileContext() {
            if (!activeModel) throw new ScriptError('InvalidReference', '虚拟文件只在该事务内可用。');
            if (!runFiles) runFiles = createFs(model('files'),
                () => { if (!activeModel) throw new ScriptError('InvalidReference', '文件上下文已失效。'); },
                () => { fileFailed = true; });
            return runFiles.fs;
        }

        const workbook = makeHandle('Workbook', {});
        let host = freeze({
            get permissions() { return permissions; },
            get localStorage() { return localStorage; },
            get sessionStorage() { return sessionStorage; },
            functions: freeze({ register(name, fn) { registerFunction(name, fn); } }),
            values: freeze({ getErrorCode(value) {
                const data = dataOf(value);
                if (typeof data.address !== 'string' || typeof data.sheetId !== 'string') throw new ScriptError('InvalidArgument', '需要 Range。');
                return model('errorCode', data.sheetId, data.address);
            } }),
            events,
            get files() { return fileContext(); }
        });

        function runAction(name, request) {
            if (!actionMode || !loaded || actionInvoked) throw new ScriptError('InvalidReference', '操作不可用。');
            if (automatic && name !== '@initialize' && !initialized) throw new ScriptError('InvalidReference', '请先应用脚本，完成主入口初始化。');
            configureNetwork(request);
            actionInvoked = true;
            activeModel = true;
            hostWriteFailed = false;
            hostWrote = false;
            fileFailed = false;
            const baseRevision = request && request.snapshot ? request.snapshot.baseRevision : 0;
            try {
                let returned;
                if (automatic && name === '@initialize') {
                    if (initialized) throw new ScriptError('InvalidReference', '主入口在本次授权中已经运行。');
                    for (const initialize of initializers) {
                        returned = initialize(workbook, host);
                        if (returned !== undefined && returned !== null) throw new ScriptError('InvalidArgument', '初始化主入口必须同步完成且不能返回值。');
                    }
                    initialized = true;
                } else if (request && request.event) {
                    const handlers = eventHandlers[request.event.name] || [];
                    const context = freeze({ workbook, host, event: deepFreeze(request.event) });
                    for (const handler of handlers) {
                        returned = handler(context);
                        if (returned !== undefined && returned !== null)
                            throw new ScriptError('InvalidArgument', '事件处理器必须同步完成且不能返回值。');
                    }
                } else {
                    const main = mains[name];
                    if (typeof main !== 'function') throw new ScriptError('ItemNotFound', '找不到操作入口：' + name);
                    returned = main(workbook, host);
                    if (returned !== undefined && returned !== null)
                        throw new ScriptError('InvalidArgument', 'main 必须同步完成且不能返回值。');
                }
                if (hostWriteFailed || fileFailed)
                    throw new ScriptError('TransactionAborted', '有一次写入失败，本次变更未提交。');
                return { operationId: request.operationId, baseRevision, name, edits: [], fileEdits: runFiles ? runFiles.edits : [] };
            } finally {
                activeModel = false;
                runFiles = null;
            }
        }
        """;
}
