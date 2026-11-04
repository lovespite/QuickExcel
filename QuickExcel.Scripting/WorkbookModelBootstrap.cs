namespace QuickExcel.Scripting;

public sealed partial class JscSession
{
    private const string WorkbookModelBootstrap = """
        let activeModel = false;
        let modelWriteFailed = false;
        function model(method, ...args) {
            if (!activeModel) throw new TypeErrorCtor('Workbook objects are available only during an action, event or console request.');
            const reply = nativeModelCall(JSON.stringify({method,args}));
            if (!reply || reply.ok !== true) throw new TypeErrorCtor(reply?.error ?? 'Workbook request failed.');
            return reply.value;
        }
        function modelScalar(value) {
            if(value===null || typeof value==='boolean' || typeof value==='string' && value.length<=32767 || typeof value==='number' && Number.isFinite(value)) return value;
            if(value && typeof value==='object' && value.kind==='error' && typeof value.code==='string') return {kind:'error',code:value.code};
            throw new TypeErrorCtor('Unsupported cell value.');
        }
        function modelWrite(method,id,address,value) {
            try {
                if(method==='setValue') value=modelScalar(value);
                else if(method==='setValues') {
                    if(!isArray(value)) throw new TypeErrorCtor('Expected a rectangular value matrix.');
                    value=value.map(row=>{if(!isArray(row)) throw new TypeErrorCtor('Expected a row array.');return row.map(modelScalar);});
                } else if(method==='setFill' && value!==null && typeof value!=='string' || method==='setFormula' && typeof value!=='string')
                    throw new TypeErrorCtor('Expected a string.');
                model(method,id,address,value);
            } catch(error) { modelWriteFailed=true; throw error; }
        }
        function modelAddress(address) {
            if (typeof address !== 'string' || address.length > 40 || !/^[A-Za-z]+[1-9][0-9]*(?::[A-Za-z]+[1-9][0-9]*)?$/.test(address))
                throw new TypeErrorCtor('Expected an A1 cell or rectangular A1:B2 range.');
            const parts=address.toUpperCase().split(':');
            return parts.length===2 && parts[0]===parts[1] ? parts[0] : parts.join(':');
        }
        function modelRange(id, input) {
            const address = modelAddress(input);
            const fill = freeze({
                setColor(color) { modelWrite('setFill',id,address,color); },
                getColor() { return model('getFill',id,address); },
                clear() { modelWrite('setFill',id,address,null); }
            });
            const format = freeze({getFill:()=>fill});
            return freeze({
                getAddress:()=>address, getWorksheet:()=>modelSheet(id), getFormat:()=>format,
                getValue() {
                    const cell=model('getCell',id,address);
                    if(cell.state!=='literal' && (cell.state!=='Fresh' && cell.state!=='Error' || !cell.hasValue)) throw new TypeErrorCtor('Formula value is unavailable; a function or dependency may be blocked.');
                    return cell.value;
                },
                getValues() { return model('getValues',id,address); },
                setValue(value) { modelWrite('setValue',id,address,value); },
                setValues(values) { modelWrite('setValues',id,address,values); },
                getFormula() { return model('getFormula',id,address); },
                setFormula(formula) { modelWrite('setFormula',id,address,formula); },
                setFormulas(formulas) { modelWrite('setFormulas',id,address,formulas); }
            });
        }
        function modelSheet(id) {
            return freeze({getId:()=>id, getName:()=>model('sheet',id).name,
                getCell(address, column) {
                    if (typeof address === 'number') {
                        if (!Number.isInteger(address) || address<0 || address>=200000 || !Number.isInteger(column) || column<0 || column>=2048)
                            throw new TypeErrorCtor('Cell indexes must be zero-based row and column indexes.');
                        let name='', n=column+1; while(n>0) { n--; name=String.fromCharCode(65+n%26)+name; n=Math.floor(n/26); }
                        address=name+(address+1);
                    }
                    address=modelAddress(address);
                    if(address.includes(':')) throw new TypeErrorCtor('getCell expects one cell.');
                    return modelRange(id,address);
                },
                getRange:address=>modelRange(id,address)
            });
        }
        const workbookApi = {
            getActiveWorksheet() { return modelSheet(model('active').sheetId); },
            getActiveCell() { const active=model('active'); return modelRange(active.sheetId,active.cell); },
            getActiveRange() {
                const active=model('active');
                if(active.ranges.length!==1) throw new TypeErrorCtor('The selection has multiple areas; use getActiveRanges().');
                return modelRange(active.sheetId,active.ranges[0]);
            },
            getActiveRanges() { const active=model('active'); return active.ranges.map(r=>modelRange(active.sheetId,r)); },
            getWorksheets() { return model('sheets').map(s=>modelSheet(s.id)); },
            getWorksheet(nameOrId) { const sheet=model('sheet',nameOrId); return sheet ? modelSheet(sheet.id) : undefined; }
        };
        """;
}
