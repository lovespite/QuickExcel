// Fixed in-memory compiler host. Runs inside a fresh JSC context; no Node, filesystem, DOM or imports.
// The trusted runner supplies fixed standard libraries and QuickExcel declarations as strings.
globalThis.qxCompile = function qxCompile(request) {
    const started = Date.now();
    const diagnostics = [];
    let diagnosticBytes = 2, diagnosticsTruncated = false; // JSON array brackets.
    const files = new Map(), ids = new Map(), originalIds = new Map();
    function bytes(text) {
        let size = 0;
        for (let i = 0; i < text.length; i++) {
            const c = text.charCodeAt(i);
            if (c < 128) size++;
            else if (c < 2048) size += 2;
            else if (c >= 0xd800 && c <= 0xdbff && i + 1 < text.length &&
                text.charCodeAt(i + 1) >= 0xdc00 && text.charCodeAt(i + 1) <= 0xdfff) { size += 4; i++; }
            else size += 3;
        }
        return size;
    }
    function recordDiagnostic(item) {
        if (item.message.length > 4096) item.message = item.message.slice(0, 4096) + '…';
        const size = bytes(JSON.stringify(item)) + (diagnostics.length ? 1 : 0);
        if (diagnostics.length >= 127 || diagnosticBytes + size > 16 * 1024 - 256) {
            diagnosticsTruncated = true; return;
        }
        diagnosticBytes += size; diagnostics.push(item);
    }
    function fail(message, file, node) {
        const position = file && node ? file.getLineAndCharacterOfPosition(node.getStart(file)) : null;
        recordDiagnostic({ code: 'QXTS', message, module: file ? originalIds.get(file.fileName) : null,
            line: position ? position.line + 1 : null, column: position ? position.character + 1 : null });
    }
    function result(outputs = []) {
        if (diagnosticsTruncated) diagnostics.push({ code: 'QXTS', message: 'Additional diagnostics omitted: diagnostic budget exceeded',
            module: null, line: null, column: null });
        return { frontend: ts.version, mode: request?.mode, success: !diagnostics.length,
            milliseconds: Date.now() - started, diagnostics, outputs: diagnostics.length ? [] : outputs };
    }
    if (!request || !['check', 'transpile'].includes(request.mode) || !Array.isArray(request.modules) ||
        request.modules.length < 1 || request.modules.length > 64) {
        fail('Invalid compiler request'); return result();
    }
    let sourceBytes = 0;
    for (const module of request.modules) {
        if (!module || typeof module.id !== 'string' || !/^[A-Za-z0-9_.-]{1,128}$/.test(module.id) ||
            module.id === '.' || module.id === '..' || ids.has(module.id) ||
            !['typescript', 'javascript'].includes(module.language) || typeof module.source !== 'string') {
            fail('Invalid or duplicate module'); return result();
        }
        const name = '/modules/' + ids.size + (module.language === 'typescript' ? '.ts' : '.js');
        sourceBytes += bytes(module.source);
        ids.set(module.id, name); originalIds.set(name, module.id); files.set(name, module.source);
    }
    if (sourceBytes > 1048576) { fail('Source exceeds 1 MiB UTF-8'); return result(); }
    if (!Array.isArray(request.entries) || !request.entries.length || request.entries.length > 64 ||
        new Set(request.entries).size !== request.entries.length || request.entries.some(id => !ids.has(id))) {
        fail('Invalid entry order'); return result();
    }
    for (const [name, source] of Object.entries(request.libraries)) {
        if (!/^lib\.es[\w.]*\.d\.ts$/.test(name) && !/^lib\.decorators(?:\.legacy)?\.d\.ts$/.test(name)) {
            fail('Unapproved standard library'); return result();
        }
        files.set('/lib/' + name, source);
    }
    files.set('/lib/quickexcel.d.ts', request.declarations);
    const options = {
        target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ESNext,
        moduleResolution: ts.ModuleResolutionKind.Bundler, moduleDetection: ts.ModuleDetectionKind.Force,
        strict: true, noEmitOnError: true, allowJs: true, checkJs: false,
        sourceMap: true, inlineSources: true, useDefineForClassFields: true,
        verbatimModuleSyntax: true, isolatedModules: true, types: [],
        skipLibCheck: false, newLine: ts.NewLineKind.LineFeed, outDir: '/out'
    };
    const parsed = new Map();
    for (const [name, source] of files) {
        if (!originalIds.has(name)) continue;
        // Pre-parsed source files must receive the same module detection as the
        // program host. Otherwise files without import/export leak declarations
        // into the global scope and cannot shadow API globals such as `$`.
        const file = ts.createSourceFile(name, source, { languageVersion: options.target,
            setExternalModuleIndicator: ts.getSetExternalModuleIndicator(options) }, true,
            name.endsWith('.ts') ? ts.ScriptKind.TS : ts.ScriptKind.JS);
        parsed.set(name, file);
        if (file.referencedFiles.length || file.typeReferenceDirectives.length || file.libReferenceDirectives.length)
            fail('Reference directives are not supported', file, file);
        function visit(node, functionDepth) {
            if (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) {
                if (node.attributes) fail('Import attributes are not supported', file, node);
                if (node.moduleSpecifier && (!ts.isStringLiteral(node.moduleSpecifier) || !ids.has(node.moduleSpecifier.text)))
                    fail('Imports must name a declared logical module ID', file, node);
            }
            if (ts.isImportEqualsDeclaration(node) || ts.isExportAssignment(node) && node.isExportEquals)
                fail('CommonJS syntax is not supported', file, node);
            if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword)
                fail('Dynamic import is not supported', file, node);
            if (ts.isDecorator(node)) fail('Decorators are not supported', file, node);
            if (!functionDepth && (ts.isAwaitExpression(node) || ts.isForOfStatement(node) && node.awaitModifier))
                fail('Asynchronous module initialization is not supported', file, node);
            ts.forEachChild(node, child => visit(child, functionDepth + (ts.isFunctionLike(node) ? 1 : 0)));
        }
        visit(file, 0);
    }
    if (diagnostics.length) return result();
    function diagnostic(d) {
        const location = d.file && d.start !== undefined ? d.file.getLineAndCharacterOfPosition(d.start) : null;
        recordDiagnostic({ code: d.code, message: ts.flattenDiagnosticMessageText(d.messageText, '\n'),
            module: d.file ? originalIds.get(d.file.fileName) ?? null : null,
            line: location ? location.line + 1 : null, column: location ? location.character + 1 : null });
    }
    const outputs = [];
    if (request.mode === 'transpile') {
        // Deliberately separate from the full checker: success here is NOT a type-check result.
        for (const [name] of parsed) {
            const emitted = ts.transpileModule(files.get(name), { compilerOptions: options, fileName: name, reportDiagnostics: true });
            (emitted.diagnostics ?? []).forEach(diagnostic);
            outputs.push({ id: originalIds.get(name), javascript: emitted.outputText, sourceMap: emitted.sourceMapText });
        }
    } else {
        const emitted = new Map();
        const host = {
            getSourceFile(name, languageVersion) {
                if (parsed.has(name)) return parsed.get(name);
                const source = files.get(name);
                return source === undefined ? undefined : ts.createSourceFile(name, source, languageVersion, true);
            },
            getDefaultLibFileName: () => '/lib/lib.es2022.d.ts',
            getCurrentDirectory: () => '/', getCanonicalFileName: name => name,
            useCaseSensitiveFileNames: () => true, getNewLine: () => '\n',
            fileExists: name => files.has(name), readFile: name => files.get(name),
            directoryExists: name => name === '/' || name === '/modules' || name === '/lib',
            getDirectories: () => [],
            writeFile: (name, content) => emitted.set(name, content),
            resolveModuleNames: names => names.map(id => {
                const name = ids.get(id);
                return name ? { resolvedFileName: name, extension: name.endsWith('.ts') ? ts.Extension.Ts : ts.Extension.Js } : undefined;
            })
        };
        const program = ts.createProgram([...ids.values(), '/lib/quickexcel.d.ts'], options, host);
        ts.getPreEmitDiagnostics(program).forEach(diagnostic);
        if (!diagnostics.length) {
            program.emit().diagnostics.forEach(diagnostic);
            for (const [id, name] of ids) {
                const outputName = '/out/' + name.split('/').pop().replace(/\.(ts|js)$/, '.js');
                const javascript = emitted.get(outputName), sourceMap = emitted.get(outputName + '.map');
                if (javascript === undefined || sourceMap === undefined) fail('Missing compiler output');
                else outputs.push({ id, javascript, sourceMap });
            }
        }
    }
    let jsBytes = 0, mapBytes = 0;
    for (const output of outputs) {
        const map = JSON.parse(output.sourceMap);
        if (map.version !== 3 || map.sources.length !== 1 || map.sourcesContent.length !== 1) {
            fail('Unexpected source map shape'); continue;
        }
        map.file = output.id + '.js'; map.sourceRoot = ''; map.sources = [output.id];
        output.sourceMap = JSON.stringify(map);
        output.javascript = output.javascript.replace(/\/\/# sourceMappingURL=[^\r\n]*$/, '//# sourceMappingURL=' + output.id + '.js.map');
        jsBytes += bytes(output.javascript); mapBytes += bytes(output.sourceMap);
    }
    if (jsBytes > 1048576) fail('Generated JavaScript exceeds execution source budget');
    if (mapBytes > 4 * 1048576) fail('Source maps exceed candidate budget');
    return result(outputs);
};
