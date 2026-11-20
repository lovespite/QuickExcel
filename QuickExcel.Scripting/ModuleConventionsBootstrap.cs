using System.Security.Cryptography;
using System.Text;

namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private static string FunctionPrefix(string id)
    {
        var stem = string.Concat(id.Select(c => char.IsAsciiLetterLower(c) || char.IsAsciiDigit(c) ? c.ToString() : $"_{(int)c:X2}"));
        if (stem.Length > 80) stem = stem[..60] + "_" + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(stem)))[..16];
        if (char.IsAsciiDigit(stem[0])) stem = "_" + stem;
        return "JS." + stem + ".";
    }

    private const string ModuleConventionsBootstrap = """
        const functionTag = Function.prototype.call.bind(Object.prototype.toString);
        function synchronous(fn, name) {
            if (typeof fn !== 'function' || functionTag(fn) !== '[object Function]')
                throw new ScriptError('InvalidArgument', name + ' must be a synchronous function.');
        }
        function actionFunction(fn, name) {
            if (typeof fn !== 'function' || !['[object Function]', '[object AsyncFunction]'].includes(functionTag(fn)))
                throw new ScriptError('InvalidArgument', name + ' must be a function.');
        }
        function bindExports(id, kind, exports, prefix) {
            if (loaded) throw new ScriptError('InvalidArgument', 'Module discovery has finished.');
            if (kind === 'udf') {
                if (actionMode || !(registrationRole() & 1)) throw new ScriptError('InvalidArgument', 'Invalid UDF role.');
                for (const name of ownKeys(exports)) {
                    const fn = exports[name];
                    if (typeof fn !== 'function') continue;
                    synchronous(fn, name);
                    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) || (prefix + name).length > 256 || fn.length > 255)
                        throw new ScriptError('InvalidArgument', 'Invalid exported UDF name or parameter count: ' + name);
                    const metadata = exports.functionDescriptors?.[name];
                    const descriptor = metadata ?? { name: prefix + name, apiVersion: 2,
                        parameters: Array.from({ length: fn.length }, (_, i) => ({ name: 'arg' + (i + 1), type: 'any' })), returns: 'any' };
                    registerFunction(descriptor, fn);
                }
            } else {
                if (!actionMode || !(registrationRole() & (kind === 'initialization' ? 4 : 2))) throw new ScriptError('InvalidArgument', 'Invalid Actions / Events role.');
                if (kind === 'events') {
                    for (const event of eventNames) {
                        const name = 'on' + event.split('.').map(part => part[0].toUpperCase() + part.slice(1)).join('');
                        if (!(name in exports)) continue;
                        const fn = exports[name];
                        synchronous(fn, name);
                        if (fn.length !== 1) throw new ScriptError('InvalidArgument', name + ' must accept exactly one context parameter.');
                        registerEvent(event, fn);
                    }
                    for (const name of ownKeys(exports))
                        if (/^on[A-Z]/.test(name) && !eventNames.some(event => name === 'on' + event.split('.').map(p => p[0].toUpperCase() + p.slice(1)).join('')))
                            throw new ScriptError('InvalidArgument', 'Unknown event handler: ' + name);
                } else if (kind === 'initialization') {
                    if (exports.main !== undefined) actionFunction(exports.main, 'main');
                    initializers.push(exports.main ?? (() => {}));
                    initialized = false;
                } else {
                    actionFunction(exports.main, id + '.main');
                    mains[id] = exports.main;
                }
            }
        }
        """;
}
