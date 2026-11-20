using System.Diagnostics;
using QuickExcel.Scripting;

static void Assert(bool condition, string message)
{
    if (!condition) throw new Exception(message);
}

static void Fails(Action action, string expected)
{
    try { action(); }
    catch (InvalidOperationException error) when (error.Message.Contains(expected, StringComparison.OrdinalIgnoreCase)) { return; }
    throw new Exception("Expected JavaScript error containing: " + expected);
}

using var session = new JscSession("""
    console.log('JSC ready');
    host.functions.register('discount', (price, rate) => price * (1 - rate));
    host.functions.register('RANGESUM', values => values.reduce((sum, value) => sum + value, 0));
    host.functions.register('TEXT', value => '中文 ' + value);
    host.functions.register('BOOL', value => value > 0);
    host.functions.register('EMPTY', () => null);
    host.functions.register('BAD', () => ({}));
    host.functions.register('PROMISE', () => Promise.resolve(42));
    host.functions.register('ENV', () => typeof std + ',' + typeof os + ',' + typeof require + ',' + typeof fetch);
    host.functions.register('NETWORK', () => fetch('https://example.com/'));
    """);

Assert(session.FunctionNames.Contains("DISCOUNT"), "Case normalization");
Assert(Equals(session.Invoke("DISCOUNT", [100d, 0.2d]), 80d), "Discount");
Assert(Equals(session.Invoke("RANGESUM", [new List<object> { 1d, 2d, 3d }]), 6d), "Ranges");
Assert(Equals(session.Invoke("TEXT", ["测试"]), "中文 测试"), "UTF-8");
Assert(Equals(session.Invoke("BOOL", [1d]), true), "Boolean");
Assert(session.Invoke("EMPTY", []) is null, "Null");
Assert(Equals(session.Invoke("ENV", []), "undefined,undefined,undefined,function"), "No OS APIs; fetch has an action-only gate");
Fails(() => session.Invoke("NETWORK", []), "only available in workbook Actions");
Assert(session.GetLogs().Single() == "JSC ready", "Console logs");
Fails(() => session.Invoke("BAD", []), "objects and promises");
Fails(() => session.Invoke("PROMISE", []), "objects and promises");
var timer = Stopwatch.StartNew();
using (var recursive = new JscSession("host.functions.register('F',()=>{function f(){return 1+f()}return f();});"))
{
    Fails(() => recursive.Invoke("F", []), "stack");
    Assert(recursive.LastOperationInvalidated, "Stack overflow invalidates the context");
}
using (var looping = new JscSession("host.functions.register('F',()=>{try {while(true){}} catch(e) {return 1}});"))
{
    timer.Restart();
    Fails(() => looping.Invoke("F", []), "time limit");
    Assert(timer.Elapsed < TimeSpan.FromSeconds(3), "Bounded uncatchable interrupt");
    Assert(looping.LastOperationInvalidated && looping.LastOperationInterrupted, "No reuse after interrupt");
}
await Task.Run(() => Assert(Equals(session.Invoke("DISCOUNT", [100d, 0.2d]), 80d), "Thread switch"));
Fails(() => { using var invalid = new JscSession("host.functions.register('BAD_NAME', () => 1)"); }, "A-Z");
Fails(() => { using var invalid = new JscSession("while (true) {}"); }, "time limit");

timer.Restart();
for (var index = 0; index < 10000; index++) session.Invoke("DISCOUNT", [100d, 0.2d]);
Console.WriteLine($"PASS: JSC runtime smoke checks; 10,000 calls: {timer.ElapsedMilliseconds} ms.");
Console.WriteLine("OPEN: JSC-MEM-001 — 64 MiB engine hard limit deferred; not a progression gate.");
