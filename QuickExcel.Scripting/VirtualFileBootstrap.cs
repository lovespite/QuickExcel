namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private const string VirtualFileBootstrap = """
        function createFs(input, reject) {
            let entries = new Map(), edits = [];
            const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
            const maxBytes = 8388608;
            function checkActive() {
                if (!activeModel || !entries) throw new ScriptError('InvalidReference', '文件上下文已失效。');
            }
            function path(p, root = false) {
                checkActive();
                if (typeof p !== 'string' || p.length > 248 || !p.startsWith('//') || /[\\:%?#\x00-\x1f\x7f-\x9f]/.test(p)) throw new TypeError('Invalid virtual path.');
                const segments = p.slice(2).split('/');
                if (!['xlsq','temp','mem'].includes(segments[0]) || segments.length < 2) throw new TypeError('Unknown virtual volume.');
                if (root && segments.length === 2 && segments[1] === '') return p;
                if (segments.slice(1).some(s => !s || s === '.' || s === '..' || s.endsWith('.') || s.endsWith(' '))) throw new TypeError('Invalid virtual path segment.');
                return p;
            }
            function decode(s) {
                const bytes = new Uint8Array(s.length / 4 * 3 - (s.endsWith('==') ? 2 : s.endsWith('=') ? 1 : 0));
                let out=0;
                for(let i=0;i<s.length;i+=4) {
                    const n=(alphabet.indexOf(s[i])<<18)|(alphabet.indexOf(s[i+1])<<12)|((alphabet.indexOf(s[i+2])&63)<<6)|(alphabet.indexOf(s[i+3])&63);
                    if(out<bytes.length) bytes[out++]=(n>>16)&255;
                    if(out<bytes.length) bytes[out++]=(n>>8)&255;
                    if(out<bytes.length) bytes[out++]=n&255;
                }
                return bytes;
            }
            function encode(bytes) {
                let text='';
                for(let i=0;i<bytes.length;i+=3) {
                    const n=(bytes[i]<<16)|((bytes[i+1]??0)<<8)|(bytes[i+2]??0);
                    text+=alphabet[(n>>18)&63]+alphabet[(n>>12)&63]+(i+1<bytes.length?alphabet[(n>>6)&63]:'=')+(i+2<bytes.length?alphabet[n&63]:'=');
                }
                return text;
            }
            function size(e) { return e.isDirectory ? 0 : e.base64.length / 4 * 3 - (e.base64.endsWith('==') ? 2 : e.base64.endsWith('=') ? 1 : 0); }
            for (const entry of input) entries.set(entry.path, { ...entry });
            function directory(p) { return p.endsWith('/') || entries.get(p)?.isDirectory || [...entries.keys()].some(k=>k.startsWith(p+'/')); }
            function get(p) {
                path(p); const entry=entries.get(p);
                if (!entry || entry.isDirectory) throw new TypeError('File does not exist: '+p);
                return entry;
            }
            function mutate(operation,p,base64=null) {
                checkActive();
                try {
                    path(p);
                    if (edits.length >= 4096) throw new TypeError('File operation quota exceeded.');
                    if (operation==='remove') {
                        if (!entries.has(p) || [...entries.keys()].some(k=>k.startsWith(p+'/'))) throw new TypeError('Missing file or directory is not empty.');
                        entries.delete(p);
                    } else {
                        if (operation==='write' && directory(p) || operation==='mkdir' && entries.has(p) && !entries.get(p).isDirectory) throw new TypeError('File/directory conflict.');
                        const next={path:p,base64,isDirectory:operation==='mkdir'};
                        let total=size(next);
                        for(const [key,value] of entries) if(key!==p) {
                            total+=size(value);
                            if(key.toLowerCase()===p.toLowerCase() || p.startsWith(key+'/') && !value.isDirectory || key.startsWith(p+'/') && !next.isDirectory) throw new TypeError('Path collision.');
                        }
                        if(total>maxBytes || !entries.has(p) && entries.size>=4096) throw new TypeError('Virtual file quota exceeded.');
                        entries.set(p,next);
                    }
                    edits.push({operation,path:p,base64});
                } catch(error) { reject(); throw error; }
            }
            const fs=freeze({
                exists(p) { path(p,true); return entries.has(p) || directory(p); },
                stat(p) {
                    path(p,true); const e=entries.get(p);
                    if (!e && !directory(p)) throw new TypeError('Path does not exist.');
                    return freeze({path:p,type:e&&!e.isDirectory?'file':'directory',size:e?size(e):0});
                },
                list(p,options={}) {
                    path(p,true); if(!directory(p)) throw new TypeError('Expected directory.');
                    const prefix=p.endsWith('/')?p:p+'/', found=new Map();
                    for(const [key,e] of entries) if(key.startsWith(prefix)) {
                        const tail=key.slice(prefix.length), split=tail.indexOf('/');
                        const name=!options.recursive && split>=0?prefix+tail.slice(0,split):key;
                        if(!found.has(name)) found.set(name,fs.stat(name));
                    }
                    return freeze([...found.values()].sort((a,b)=>a.path<b.path?-1:a.path>b.path?1:0));
                },
                mkdir(p) { mutate('mkdir',p); },
                remove(p) { mutate('remove',p); },
                readBytes(p) { return decode(get(p).base64); },
                writeBytes(p,bytes) {
                    checkActive();
                    try {
                        if (!(bytes instanceof Uint8Array) || bytes.length>maxBytes) throw new TypeError('Expected a bounded Uint8Array.');
                        mutate('write',p,encode(bytes));
                    } catch(error) { reject(); throw error; }
                },
                readText(p) {
                    const bytes=fs.readBytes(p); let escaped='';
                    for(const b of bytes) escaped+='%'+b.toString(16).padStart(2,'0');
                    return decodeURIComponent(escaped);
                },
                writeText(p,text) {
                    checkActive();
                    try {
                        if(typeof text!=='string' || text.length>maxBytes) throw new TypeError('Expected bounded text.');
                        const encoded=encodeURIComponent(text), bytes=[];
                        for(let i=0;i<encoded.length;i++) { if(encoded[i]==='%') { bytes.push(parseInt(encoded.slice(i+1,i+3),16)); i+=2; } else bytes.push(encoded.charCodeAt(i)); }
                        fs.writeBytes(p,new Uint8Array(bytes));
                    } catch(error) { reject(); throw error; }
                },
                readJson(p) { return JSON.parse(fs.readText(p)); },
                writeJson(p,value) { checkActive(); try { fs.writeText(p,JSON.stringify(value)); } catch(error) { reject(); throw error; } },
                copy(from,to) { fs.writeBytes(to,fs.readBytes(from)); },
                move(from,to) { if(from===to) { get(from); return; } fs.copy(from,to); fs.remove(from); }
            });
            return { fs, edits, release() { entries = null; edits = null; } };
        }
        """;
}
