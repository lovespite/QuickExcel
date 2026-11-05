namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private const string StorageBootstrap = """
        function createStorage(scope) {
            function text(value, write) {
                if (typeof value !== 'string') {
                    if (write) hostWriteFailed = true;
                    throw new ScriptError('InvalidArgument', '存储键和值必须是字符串。');
                }
                return value;
            }
            function all(operation) {
                const result = [];
                let after = null;
                for (;;) {
                    const page = model('getStorage', scope, operation, after);
                    for (const item of page.items) result.push(item);
                    if (page.done) return result;
                    after = page.next;
                }
            }
            return freeze({
                get(key) { const value = model('getStorage', scope, 'get', text(key, false)); return value === null ? undefined : value; },
                set(key, value) { model('setStorage', scope, 'set', [text(key, true), text(value, true)]); },
                clear() { model('setStorage', scope, 'clear', null); },
                remove(key) { model('setStorage', scope, 'remove', text(key, true)); },
                keys() { return all('keys'); },
                values() { return all('values'); },
                entries() { return all('entries'); }
            });
        }
        const localStorage = createStorage('local');
        const sessionStorage = createStorage('session');
        Object.defineProperty(globalThis, 'localStorage', { value: localStorage });
        Object.defineProperty(globalThis, 'sessionStorage', { value: sessionStorage });
        """;
}
