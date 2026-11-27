using System.Text;
using QuickExcel.Calculation;
using QuickExcel.Core;
using QuickExcel.Domain;
using QuickExcel.Runtime;
using QuickExcel.Storage;
using static QuickExcel.Core.Test.M3CalculationTests;

namespace QuickExcel.Core.Test;

public class ScriptUserFileTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("quickexcel-file-api-").FullName;
    public void Dispose() => Directory.Delete(_dir, true);
    private WorkbookDocument Open(string action, ScriptFilePicker picker, string initialization = "export function main() {}")
    {
        var modules = new ScriptModuleSet(new Dictionary<string, ScriptModule> {
            ["main"] = new("scripts/main.js", initialization), ["action"] = new("scripts/action.js", action)
        }, ["main", "action"], scriptApiVersion: 4, bindings: [new("main","initialization"),new("action","actions")]);
        var doc = new WorkbookDocument(Book(Sheet("Data")).WithScriptModules(modules),
            new(new TypedWorkbookStorage(), new WorkbookCalculator(new ScriptWorkerRuntimeFactory()), new ScriptWorkerActionExecutor()));
        doc.Session.UserFilePicker = picker;
        doc.EnableScriptsAndCalculate();
        return doc;
    }
    private ScriptFilePicker Select(params string[] paths) => (request, begin, token) =>
    {
        using var pause = begin();
        return Task.FromResult(new ScriptFileSelection(paths));
    };

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void ActiveWorkbookSaveAsPublishesValidatedCurrentDraftBeforeALaterRollback(string extension)
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "export" + extension);
        File.WriteAllText(path, "replace");
        using var doc = Open($$"""
            export async function main(w,h) {
              const [out] = await h.fs.getUserSelectedFiles({ mode:'save', fileMode:'create', fileAccess:'w' });
              const sheet = w.getActiveWorksheet();
              sheet.getRange('A1').setValue(21);
              sheet.getRange('B1').setFormula('=A1*2');
              w.files.writeText('//xlsq/report.txt','report');
              h.localStorage.set('record','stored');
              h.sessionStorage.set('transient','only here');
              const saving = w.saveAs(out,'{{extension[1..]}}');
              sheet.getRange('A1').setValue(99);
              h.localStorage.set('record','changed later');
              w.files.writeText('//xlsq/report.txt','changed later');
              await saving;
              throw Error('AFTER_SAVE_AS');
            }
            """, Select(path));
        Assert.Contains("AFTER_SAVE_AS", Assert.ThrowsAny<Exception>(() => doc.RunAction("action")).Message);
        Assert.Equal("", doc.GetCellValue(0, 0, 0));
        var saved = new TypedWorkbookStorage().Open(path);
        Assert.Equal(21, saved.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal("=A1*2", saved.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        if (extension == ".xlsq")
        {
            Assert.Equal(doc.Snapshot.ScriptModules.Modules["action"].Source, saved.ScriptModules.Modules["action"].Source);
            Assert.Equal("report", Encoding.UTF8.GetString(saved.PackageFiles.Entries["//xlsq/report.txt"].Data.AsSpan()));
            Assert.Equal("stored", saved.LocalStorage.Get("record"));
        }
        else
        {
            Assert.False(TypedWorkbookStorage.RequiresScriptContainer(saved));
        }
    }

    [Theory]
    [InlineData("xlsx", "xlsq", "r")]
    [InlineData("xlsq", "xlsx", "rw")]
    [InlineData("xlsq", "xlsq", "rw")]
    [InlineData("xlsx", "xlsx", "rw")]
    public void SelectedWorkbookSaveAsPublishesCopiesWithoutClosingOrRetargeting(string from, string to, string access)
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "source." + from);
        var copy = Path.Combine(_dir, "copy." + to);
        var storage = new TypedWorkbookStorage();
        var source = Book(Sheet("Items", (0, 0, N(3)), (0, 1, F("=A1*2"))));
        if (from == "xlsq") source = source.WithScriptSource("throw Error('do not run');")
            .WithLocalStorage(ScriptStorage.Empty.Set("key", "value")).WithPackageFiles(new([new("//xlsq/item.txt", [1,2,3])]));
        storage.Save(path, source);
        var original = File.ReadAllBytes(path);
        using var doc = Open($$"""
            export async function main(w,h) {
              const [file] = await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'{{access}}'});
              const book = await file.openWorkbook();
              const sheet = book.workbook.getWorksheet('Items');
              if (sheet.getRange('B1').getValue() !== 6) throw Error('cannot calculate the selected source');
              const [dest] = await h.fs.getUserSelectedFiles({mode:'save',fileMode:'open_or_create',fileAccess:'w'});
              await book.saveAs(dest,'{{to}}');
              if ('{{access}}' === 'rw') sheet.getRange('A1').setValue(5);
              await book.saveAs(dest,'{{to}}');
              if (sheet.getRange('B1').getFormula() !== '=A1*2') throw Error('closed source');
              await book.abort();
            }
            """, (request, _, _) => Task.FromResult(new ScriptFileSelection([request.Save ? copy : path])));
        doc.RunAction("action");
        Assert.Equal(original, File.ReadAllBytes(path));
        var saved = storage.Open(copy);
        Assert.Equal(access == "rw" ? 5 : 3, saved.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal("=A1*2", saved.Sheets[0].GetCell(0, 1)!.Content.FormulaText);
        if (to == "xlsx") Assert.False(TypedWorkbookStorage.RequiresScriptContainer(saved));
        if (to == "xlsq" && from == "xlsq")
        {
            Assert.Equal(source.ScriptSource, saved.ScriptSource);
            Assert.True(source.PackageFiles.ContentEquals(saved.PackageFiles));
            Assert.True(source.LocalStorage.ContentEquals(saved.LocalStorage));
        }
    }

    [Fact]
    public void SaveAsChecksTargetGrantFormatAndVersionBeforePublishing()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "target.bin");
        File.WriteAllText(path, "initial");
        using var doc = Open("""
            export async function main(w,h) {
              const [ro] = await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'r'});
              const [dest] = await h.fs.getUserSelectedFiles({mode:'save',fileMode:'create',fileAccess:'w'});
              async function expect(code, call) {
                try { await call(); } catch(e) { if(e.code === code) return; throw e; }
                throw Error('expected ' + code);
              }
              await expect('PermissionDenied', () => w.saveAs(ro,'xlsx'));
              await expect('InvalidReference', () => w.saveAs({...dest},'xlsx'));
              await expect('InvalidArgument', () => w.saveAs(dest,'csv'));
              await ro.release();
              await expect('InvalidReference', () => w.saveAs(ro,'xlsq'));
              await h.fs.getUserSelectedDirectory();
              await expect('Conflict', () => w.saveAs(dest,'xlsx'));
            }
            """, (request, _, _) => {
                if (!request.Directory) return Task.FromResult(new ScriptFileSelection([path]));
                File.WriteAllText(path, "changed outside");
                return Task.FromResult(new ScriptFileSelection([]));
            });
        doc.RunAction("action");
        Assert.Equal("changed outside", File.ReadAllText(path));
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void SelectedWorkbookEditsPreserveContentAndCalculateWithoutRunningStoredScripts(string extension)
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "selected" + extension);
        var storage = new TypedWorkbookStorage();
        var source = Book(Sheet("Data", (0, 0, N(2)), (0, 1, F("=A1+1", 99))), Sheet("Keep", (2, 2, T("untouched"))));
        if (extension == ".xlsq")
            source = source.WithScriptSource("throw Error('external script ran')").WithLocalStorage(ScriptStorage.Empty.Set("key", "kept"))
                .WithPackageFiles(new([new("//xlsq/attachment.txt", [42])]));
        storage.Save(path, source);
        using var doc = Open("""
            export async function main(w, host) {
              const [file] = await host.fs.getUserSelectedFiles({ mode:'open', fileMode:'open', fileAccess:'rw' });
              const book = await file.openWorkbook();
              const data = book.workbook.getWorksheet('data');
              if (book.workbook.getWorksheets().map(s => s.getName()).join() !== 'Data,Keep' ||
                  data.getUsedRange().getAddress() !== 'A1:B1')
                throw Error('sheet metadata');
              const values = data.getRange('A1:C1').getValues();
              if (JSON.stringify(values) !== '[[2,3,""]]') throw Error('input ' + JSON.stringify(values));
              data.getRange('A1:A1').setValues([[8]]);
              if (data.getRange('B1').getValue() !== 9) throw Error('dependent formula');
              const results = book.workbook.addWorksheet('Results');
              results.getRange('A1:D1').setValues([[true, '=literal', 12.5, 'cleared']]);
              results.getRange('D1').clear(QuickExcel.ClearApplyTo.contents);
              if (results.getRange('D1').getValue() !== '') throw Error('clear');
              results.getRange('C2').setFormulas([['=C1*2']]);
              if (JSON.stringify(results.getRange('B2:C2').getFormulas()) !== '[["","=C1*2"]]') throw Error('formula text');
              w.getActiveCell().setValue(results.getRange('C2').getValue());
              await book.close();
              const again = await file.openWorkbook();
              if (again.workbook.getWorksheet('Results').getRange('B1').getValue() !== '=literal') throw Error('reopen');
              await again.close();
            }
            """, Select(path));
        doc.RunAction("action");
        Assert.Equal("25", doc.GetCellValue(0, 0, 0));
        Assert.True(doc.Undo());
        var saved = storage.Open(path);
        Assert.Equal(8, saved.Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal("untouched", saved.Sheets[1].GetCell(2, 2)!.Value.TextValue);
        Assert.Equal("=C1*2", saved.Sheets[2].GetCell(1, 2)!.Content.FormulaText);
        if (extension == ".xlsq")
        {
            Assert.Equal(source.DocumentId, saved.DocumentId);
            Assert.Equal(source.Sheets[0].Id, saved.Sheets[0].Id);
            Assert.Equal(source.ScriptSource, saved.ScriptSource);
            Assert.True(source.PackageFiles.ContentEquals(saved.PackageFiles));
            Assert.True(source.LocalStorage.ContentEquals(saved.LocalStorage));
        }
    }

    [Fact]
    public void SelectedWorkbookHandlesKeepIdenticalCopiesAndActiveWorkbookIndependent()
    {
        if (!OperatingSystem.IsWindows()) return;
        var first = Path.Combine(_dir, "first.xlsq");
        var second = Path.Combine(_dir, "second.xlsq");
        using var doc = Open("""
            export async function main(w, host) {
              const [one, two] = await host.fs.getUserSelectedFiles({mode:'open', fileMode:'open', fileAccess:'rw', multiSelect:true});
              const left = await one.openWorkbook();
              const right = await two.openWorkbook();
              try { await one.openWorkbook(); throw Error('two writers to the same file'); }
              catch (e) { if (e.code !== 'Locked') throw e; }
              const activeSheet = w.getWorksheet('Data');
              const leftSheet = left.workbook.getWorksheet('Data');
              const rightSheet = right.workbook.getWorksheet('Data');
              if (leftSheet.getId() !== rightSheet.getId() || leftSheet.getId() !== activeSheet.getId())
                throw Error('sheet IDs must match for this test');
              const activeCell = activeSheet.getRange('A1');
              const leftCell = leftSheet.getRange('A1');
              const rightCell = rightSheet.getRange('A1');
              const leftFont = leftCell.getFormat().getFont();
              const rightFont = rightCell.getFormat().getFont();
              function expectInvalidReference(call) {
                try { call(); } catch (e) { if (e.code === 'InvalidReference') return; throw e; }
                throw Error('expected InvalidReference');
              }
              expectInvalidReference(() => leftCell.getIntersection(rightCell));
              expectInvalidReference(() => activeCell.getBoundingRect(leftCell));
              expectInvalidReference(() => rightCell.getBoundingRect(activeCell));
              leftCell.setValue(11);
              rightCell.setValue(22);
              leftSheet.getRange('B1').copyFrom(leftCell);
              rightSheet.getRange('B1').copyFrom(rightCell);
              leftFont.setBold(true);
              activeCell.setValue(33);
              rightFont.setItalic(true);
              if (leftCell.getValue() !== 11 || rightCell.getValue() !== 22 || activeCell.getValue() !== 33 ||
                  !leftFont.getBold() || rightFont.getBold() || leftFont.getItalic() || !rightFont.getItalic())
                throw Error('edits crossed documents');
              await left.close();
              expectInvalidReference(() => leftCell.getAddress());
              expectInvalidReference(() => leftFont.getBold());
              const rejected = await one.openWorkbook();
              const target = rejected.workbook.getWorksheet('Data').getRange('A1');
              target.setValue(99);
              expectInvalidReference(() => target.copyFrom(rightCell));
              try { await rejected.close(); throw Error('published draft with invalid cross-workbook write'); }
              catch (e) { if (e.code !== 'TransactionAborted') throw e; }
              rightCell.setValue(44);
              await right.close();
              if (activeCell.getValue() !== 33) throw Error('active workbook changed');
            }
            """, Select(first, second));
        var storage = new TypedWorkbookStorage();
        storage.Save(first, doc.Snapshot);
        File.Copy(first, second);
        doc.RunAction("action");
        void AssertCell(WorkbookSnapshot snapshot, int value, int? copied, bool bold, bool italic)
        {
            var sheet = snapshot.Sheets.Single();
            Assert.Equal(doc.Snapshot.Sheets[0].Id, sheet.Id);
            Assert.Equal(value, sheet.GetCell(0, 0)!.Value.NumberValue);
            if (copied is { } expected) Assert.Equal(expected, sheet.GetCell(0, 1)!.Value.NumberValue);
            else Assert.Null(sheet.GetCell(0, 1));
            var style = TablePresentation.GetStyle(snapshot, sheet, 0, 0);
            Assert.Equal(bold, style.Bold);
            Assert.Equal(italic, style.Italic);
        }
        AssertCell(storage.Open(first), 11, 11, true, false);
        AssertCell(storage.Open(second), 44, 22, false, true);
        AssertCell(doc.Snapshot, 33, null, false, false);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SelectedWorkbookCellBudgetsAreCumulativeAcrossOpenFiles(bool write)
    {
        if (!OperatingSystem.IsWindows()) return;
        var first = Path.Combine(_dir, "one.xlsx");
        var second = Path.Combine(_dir, "two.xlsx");
        var storage = new TypedWorkbookStorage();
        storage.Save(first, Book(Sheet("Data", (0, 0, N(7)))));
        File.Copy(first, second);
        using var doc = Open($$"""
            export async function main(w, host) {
              const [one, two] = await host.fs.getUserSelectedFiles({mode:'open', fileMode:'open', fileAccess:'rw', multiSelect:true});
              const left = await one.openWorkbook();
              const right = await two.openWorkbook();
              function operate(book) {
                const sheet = book.workbook.getWorksheet('Data');
                {{(write ? "sheet.getRange('A1:A10001').clear(QuickExcel.ClearApplyTo.contents);" : "sheet.getRange('A1:A50001').getFormat().getRowHeight();")}}
              }
              operate(left);
              try { operate(right); throw Error('fresh budget for second workbook'); }
              catch (e) { if (e.code !== 'BudgetExceeded') throw e; }
              await right.abort();
              await left.abort();
            }
            """, Select(first, second));
        doc.RunAction("action");
        Assert.Equal(7, storage.Open(first).Sheets[0].GetCell(0, 0)!.Value.NumberValue);
        Assert.Equal(7, storage.Open(second).Sheets[0].GetCell(0, 0)!.Value.NumberValue);
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsq")]
    public void SelectedWorkbookCreateRequiresSuccessfulExplicitClose(string extension)
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "new" + extension);
        using var doc = Open("""
            export async function main(w, host) {
              const [file] = await host.fs.getUserSelectedFiles({mode:'save',fileMode:'create',fileAccess:'w'});
              const draft = await file.openWorkbook();
              const sheet = draft.workbook.getWorksheet('Sheet1');
              sheet.getRange('A1').setValue(1);
              try { sheet.getRange('A1').getValue(); throw Error('read with write-only grant'); }
              catch (e) { if (e.code !== 'PermissionDenied') throw e; }
              try { sheet.getRange('A1:B1').setValues([[2]]); throw Error('accepted invalid write'); }
              catch (e) { if (e.code !== 'InvalidArgument') throw e; }
              try { await draft.close(); throw Error('saved partial workbook'); }
              catch (e) { if (e.code !== 'TransactionAborted') throw e; }
              if ((await file.stat()).exists) throw Error('created too soon');
              const second = await file.openWorkbook();
              second.workbook.getWorksheet('Sheet1').getRange('A1').setValue(true);
              await second.close();
              try { await file.openWorkbook(); throw Error('read existing content with write-only grant'); }
              catch (e) { if (e.code !== 'PermissionDenied') throw e; }
            }
            """, Select(path));
        doc.RunAction("action");
        var saved = new TypedWorkbookStorage().Open(path);
        Assert.True(saved.Sheets.Single().GetCell(0, 0)!.Value.BooleanValue);
        if (extension == ".xlsq") Assert.Equal(4, saved.ScriptModules.ScriptApiVersion);
    }

    [Fact]
    public void SelectedWorkbookGrantsRejectInvalidFilesAndProtectExistingBytes()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "keep.xlsx");
        var malformed = Path.Combine(_dir, "malformed.xlsq");
        var unsupported = Path.Combine(_dir, "input.txt");
        new TypedWorkbookStorage().Save(path, Book(Sheet("Original")));
        var original = File.ReadAllBytes(path);
        File.WriteAllText(malformed, "broken");
        File.WriteAllText(unsupported, "other");
        using var doc = Open("""
            export async function main(w, h) {
              const [good, malformed, other] = await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'r',multiSelect:true});
              const reader = await good.openWorkbook();
              try { reader.workbook.addWorksheet('Bad'); throw Error('read-only write'); } catch(e) { if(e.code !== 'PermissionDenied') throw e; }
              await reader.abort();
              for (const [file,code] of [[malformed,'FileError'],[other,'UnsupportedApi']]) {
                try { await file.openWorkbook(); throw Error('opened invalid workbook'); } catch(e) { if(e.code !== code) throw e; }
              }
              const [out] = await h.fs.getUserSelectedFiles({mode:'save',fileMode:'create',fileAccess:'rw'});
              const draft = await out.openWorkbook();
              draft.workbook.addWorksheet('Discarded');
              await draft.abort();
              const abandoned = await out.openWorkbook();
              abandoned.workbook.getWorksheet('Sheet1').getRange('A1').setValue(2);
            }
            """, (request, _, _) => Task.FromResult(new ScriptFileSelection(request.Multiple ? [path, malformed, unsupported] : [path])));
        doc.RunAction("action");
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public void SelectedWorkbookRefusesLossySaveOfAnUnmodeledWorksheetFeature()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "custom-views.xlsx");
        var original = Fixtures.WorkbookPackageFixture.Xlsx("", new Dictionary<string, string>
        {
            ["xl/worksheets/sheet1.xml"] = """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><customSheetViews/>
                </worksheet>
                """
        });
        File.WriteAllBytes(path, original);
        using var doc = Open("""
            export async function main(w, h) {
              const [file] = await h.fs.getUserSelectedFiles({mode:'open', fileMode:'open', fileAccess:'rw'});
              const book = await file.openWorkbook();
              const cell = book.workbook.getWorksheet('Sheet1').getRange('A1');
              if (cell.getValue() !== 7) throw Error('original value');
              cell.setValue(8);
              try { await book.close(); throw Error('discarded custom views'); }
              catch (e) { if (e.code !== 'UnsupportedApi') throw e; }
            }
            """, Select(path));
        doc.RunAction("action");
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public void SelectedWorkbookCloseDetectsExternalEdit()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "conflict.xlsx");
        var storage = new TypedWorkbookStorage();
        storage.Save(path, Book(Sheet("Original")));
        using var doc = Open("""
            export async function main(w,h) {
              const [file] = await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'rw'});
              const book = await file.openWorkbook();
              book.workbook.getWorksheet('Original').getRange('A1').setValue(1);
              await h.fs.getUserSelectedDirectory();
              try { await book.close(); throw Error('overwrote external edit'); } catch(e) { if(e.code !== 'Conflict') throw e; }
            }
            """, (request, _, _) => {
                if (!request.Directory) return Task.FromResult(new ScriptFileSelection([path]));
                storage.Save(path, Book(Sheet("External")));
                return Task.FromResult(new ScriptFileSelection([]));
            });
        doc.RunAction("action");
        Assert.Equal("External", storage.Open(path).Sheets.Single().Name);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FilesAndDirectoriesUseActionScopedStreams(bool buffered)
    {
        if (!OperatingSystem.IsWindows()) return;
        var options = buffered ? "" : "{buffering:'stream'}";
        File.WriteAllText(Path.Combine(_dir,"input"), "一\r\n\rZ\n😀", new UnicodeEncoding(false,true,true));
        using var doc = Open($$"""
            export async function main(w, host) {
              const dir = await host.fs.getUserSelectedDirectory({fileAccess:'rw',allowDelete:true});
              if ('path' in dir || dir.type !== 'directory') throw Error('path leak');
              const names=[]; for await(const item of dir.entries()) names.push(item.name);
              if(!names.includes('input')) throw Error('enumeration');
              const f=await dir.getFile('input'), info=await f.stat();
              if(info.size!==18 || info.name!=='input' || !info.modifiedAt) throw Error('metadata '+JSON.stringify(info));
              const s=await f.open({{options}});
              const lines=[]; let line; while((line=await s.readLine('utf-16le'))!==null) lines.push(line);
              await s.close();
              const out=await dir.getDirectory('out',{create:true});
              const dest=await out.getFile('result',{fileMode:'create'});
              const stream=await dest.open({{options}});
              await stream.writeLine(lines.join('|'));
              const input=new Uint8Array([33,65,66,44]).subarray(1);
              await stream.write(input,0,2);
              await stream.seek(0);
              const bytes=new Uint8Array(9).fill(88);
              if (await stream.read(bytes.subarray(2,6),1,2)!==2 || bytes[2]!==88 || bytes[5]!==88) throw Error('subarray');
              await stream.close();
              if(!(await dest.stat()).exists) throw Error('not published');
              const marker=await out.getFile('discard',{fileMode:'open_or_create'});
              const abandoned=await marker.open(); abandoned.writeText('no');
              await abandoned.abort();
              w.getActiveCell().setValue(lines.length);
            }
            """, Select(_dir));
        doc.RunAction("action");
        Assert.Equal("4",doc.GetCellValue(0,0,0));
        Assert.Equal("一||Z|😀\nAB",File.ReadAllText(Path.Combine(_dir,"out","result")));
        Assert.False(File.Exists(Path.Combine(_dir,"out","discard")));
        Assert.True(doc.Undo());
        Assert.True(File.Exists(Path.Combine(_dir,"out","result")));
    }

    [Fact]
    public void NativePickerPauseCancelAndDeniedCombinations()
    {
        if (!OperatingSystem.IsWindows()) return;
        int calls = 0;
        using var doc = Open("""
            export async function main(w,h) {
              try { await h.fs.getUserSelectedFiles({mode:'save',multiSelect:true,fileAccess:'w',fileMode:'create'}); throw Error('accepted'); }
              catch(e) { if(e.code!=='InvalidArgument') throw e; }
              if ((await h.fs.getUserSelectedDirectory())!==null) throw Error('cancel');
              if ((await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'r'})).length) throw Error('cancel');
              w.getActiveCell().setValue('ok');
            }
            """, async (_, begin, token) => {
                using (begin()) { if (++calls == 1) await Task.Delay(2400,token); }
                return new([]);
            });
        doc.Session.Network.Configure(new(ActionTimeoutSeconds: 2));
        doc.RunAction("action");
        Assert.Equal(2,calls);
        Assert.Equal("ok",doc.GetCellValue(0,0,0));
    }

    [Fact]
    public void FilePickerNormalizesOptionalLeadingDotsAndRejectsInvalidExtensions()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path = Path.Combine(_dir, "sample.xlsq");
        File.WriteAllBytes(path, [1, 2, 3]);
        var requests = new List<ScriptFilePickerRequest>();
        using var doc = Open("""
            export async function main(w, h) {
              const options = { mode: 'open', fileMode: 'open', fileAccess: 'r', multiSelect: true };
              for (const extension of ['', '.', '..xlsq', '*.xlsq', 'tar.gz', 'dir/txt', '.' + 'x'.repeat(25)]) {
                try {
                  await h.fs.getUserSelectedFiles({ ...options, filters: [{ name: 'Invalid', extensions: [extension] }] });
                  throw Error('Accepted extension ' + extension);
                } catch (error) {
                  if (error.code !== 'InvalidArgument') throw error;
                }
              }
              const [file] = await h.fs.getUserSelectedFiles({ ...options, filters: [
                { name: 'With a dot', extensions: ['.xlsq', '.' + 'x'.repeat(24)] },
                { name: 'Without a dot', extensions: ['xlsq'] }
              ] });
              w.getActiveCell().setValue((await file.stat()).size);
            }
            """, (request, begin, _) =>
            {
                using var interaction = begin();
                requests.Add(request);
                return Task.FromResult(new ScriptFileSelection([path]));
            });
        doc.RunAction("action");
        var filters = Assert.Single(requests).Filters;
        Assert.Collection(filters,
            filter => Assert.Equal(new[] { "xlsq", new string('x', 24) }, filter.Extensions),
            filter => Assert.Equal(new[] { "xlsq" }, filter.Extensions));
        Assert.Equal("3", doc.GetCellValue(0, 0, 0));
    }

    [Fact]
    public void ReopenPreservesContentAndPublicationSurvivesLaterFailure()
    {
        if (!OperatingSystem.IsWindows()) return;
        var path=Path.Combine(_dir,"save.txt"); File.WriteAllText(path,"abcdef");
        using var doc=Open("""
            export async function main(w,h) {
              const [file]=await h.fs.getUserSelectedFiles({mode:'save',fileMode:'open_or_create',fileAccess:'w'});
              const s=await file.open();
              try { s.readLine(); throw Error('read granted'); } catch(e) {if(e.code!=='PermissionDenied') throw e;}
              s.seek(1); s.writeText('X'); s.seek(4); s.writeText('Y');
              await s.close();
              const next=await file.open(); next.seek(6); next.writeText('Z'); await next.close();
              w.getActiveCell().setValue('rolled back'); throw Error('AFTER_PUBLICATION');
            }
            """,Select(path));
        Assert.Contains("AFTER_PUBLICATION",Assert.ThrowsAny<Exception>(()=>doc.RunAction("action")).Message);
        Assert.Equal("aXcdYfZ",File.ReadAllText(path));
        Assert.Equal("",doc.GetCellValue(0,0,0));
    }

    [Fact]
    public void DirectoryDeleteRequiresPermissionAndAChildName()
    {
        if (!OperatingSystem.IsWindows()) return;
        Directory.CreateDirectory(Path.Combine(_dir,"nested","child"));
        File.WriteAllText(Path.Combine(_dir,"nested","child","test"),"keep until authorized");
        using var doc=Open("""
            export async function main(w,h) {
              const r=await h.fs.getUserSelectedDirectory();
              try { await r.removeEntry('nested',{recursive:true}); throw Error('unexpected delete'); } catch(e) { if(e.code!=='PermissionDenied')throw e; }
              try { await r.getDirectory('..'); throw Error('unexpected parent'); } catch(e) { if(e.code!=='InvalidArgument')throw e; }
              const d=await h.fs.getUserSelectedDirectory({allowDelete:true});
              await d.removeEntry('nested',{recursive:true});
            }
            """,Select(_dir));
        doc.RunAction("action");
        Assert.False(Directory.Exists(Path.Combine(_dir,"nested")));
    }

    [Fact]
    public void PublicationDetectsLaterEditsAndProtectsTheOpenWorkbook()
    {
        if (!OperatingSystem.IsWindows()) return;
        var workbook = Path.Combine(_dir, "active.xlsq");
        var existing = Path.Combine(_dir, "existing.txt");
        var newFile = Path.Combine(_dir, "new.txt");
        File.WriteAllText(workbook, "active workbook");
        File.WriteAllText(existing, "before");
        int choices = 0;
        using var doc = Open("""
            export async function main(w,h) {
              const d = await h.fs.getUserSelectedDirectory({fileAccess:'rw',allowDelete:true});
              const active = await d.getFile('active.xlsq');
              try { await active.open(); throw Error('workbook writable'); } catch(e) { if(e.code!=='PermissionDenied') throw e; }
              try { await d.removeEntry('active.xlsq'); throw Error('workbook deletable'); } catch(e) { if(e.code!=='PermissionDenied') throw e; }
              const existing = await (await d.getFile('existing.txt')).open();
              const fresh = await (await d.getFile('new.txt',{fileMode:'create'})).open();
              existing.writeText('edited'); fresh.writeText('created');
              await h.fs.getUserSelectedDirectory();
              for (const stream of [existing,fresh]) {
                try { await stream.close(); throw Error('overwrote external file'); } catch(e) { if(e.code!=='Conflict') throw e; }
              }
            }
            """, (_, begin, _) =>
            {
                using var interaction = begin();
                if (++choices == 1) return Task.FromResult(new ScriptFileSelection([_dir], workbook));
                File.WriteAllText(existing, "changed outside");
                File.WriteAllText(newFile, "created outside");
                return Task.FromResult(new ScriptFileSelection([], workbook));
            });
        using (File.Open(workbook, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            doc.RunAction("action");
        Assert.Equal("active workbook", File.ReadAllText(workbook));
        Assert.Equal("changed outside", File.ReadAllText(existing));
        Assert.Equal("created outside", File.ReadAllText(newFile));
    }

    [Fact]
    public void DirectoryJunctionsAndHardLinkedWritesAreExcludedFromTheGrant()
    {
        if (!OperatingSystem.IsWindows()) return;
        var root = Directory.CreateDirectory(Path.Combine(_dir, "selected"));
        var outside = Directory.CreateDirectory(Path.Combine(_dir, "outside"));
        var link = Path.Combine(root.FullName, "linked");
        File.WriteAllText(Path.Combine(outside.FullName, "keep.txt"), "keep");
        var command = new System.Diagnostics.ProcessStartInfo("powershell.exe") { UseShellExecute = false, CreateNoWindow = true };
        command.Environment["QE_TEST_JUNCTION"] = link;
        command.Environment["QE_TEST_TARGET"] = outside.FullName;
        command.Environment["QE_TEST_HARDLINK"] = Path.Combine(root.FullName, "alias");
        foreach (var arg in new[] { "-NoProfile", "-NonInteractive", "-Command", """
            $ErrorActionPreference='Stop'
            New-Item -ItemType Junction -Path $env:QE_TEST_JUNCTION -Target $env:QE_TEST_TARGET | Out-Null
            New-Item -ItemType HardLink -Path $env:QE_TEST_HARDLINK -Target (Join-Path $env:QE_TEST_TARGET 'keep.txt') | Out-Null
            """ })
            command.ArgumentList.Add(arg);
        using var setup = System.Diagnostics.Process.Start(command)!;
        Assert.True(setup.WaitForExit(10000));
        Assert.Equal(0, setup.ExitCode);
        try
        {
            using var doc = Open("""
                export async function main(w,h) {
                  const root = await h.fs.getUserSelectedDirectory({fileAccess:'rw',allowDelete:true});
                  for await(const item of root.entries()) if(item.name==='linked') throw Error('listed junction');
                  try { await root.getDirectory('linked'); throw Error('opened junction'); } catch(e) { if(e.code!=='UnsupportedApi') throw e; }
                  try { await root.removeEntry('linked',{recursive:true}); throw Error('deleted junction'); } catch(e) { if(e.code!=='UnsupportedApi') throw e; }
                  const alias = await root.getFile('alias');
                  try { await alias.open(); throw Error('writable hard link'); } catch(e) { if(e.code!=='PermissionDenied') throw e; }
                  try { await root.removeEntry('alias'); throw Error('deleted hard link'); } catch(e) { if(e.code!=='PermissionDenied') throw e; }
                }
                """, Select(root.FullName));
            doc.RunAction("action");
            Assert.Equal("keep", File.ReadAllText(Path.Combine(outside.FullName, "keep.txt")));
        }
        finally { Directory.Delete(link); }
    }

    [Fact]
    public void MultipleSelectionsAndRetainedObjectsAcrossActions()
    {
        if (!OperatingSystem.IsWindows()) return;
        var first = Path.Combine(_dir, "first");
        var second = Path.Combine(_dir, "second");
        File.WriteAllText(first, "first"); File.WriteAllText(second, "second");
        using var doc = Open("""
            let old, buffered;
            export async function main(w,h) {
              if (old) {
                try { await old.stat(); throw Error('old grant'); } catch(e) { if(e.code!=='InvalidReference') throw e; }
                try { buffered.readLine(); throw Error('old buffer'); } catch(e) { if(e.code!=='InvalidReference') throw e; }
                w.getActiveCell().setValue('expired'); return;
              }
              const files = await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'r',multiSelect:true});
              if(files.length!==2) throw Error('selection count');
              old=files[1]; buffered=await old.open();
              w.getActiveCell().setValue(buffered.readLine());
            }
            """, Select(first, second));
        doc.RunAction("action");
        Assert.Equal("second", doc.GetCellValue(0,0,0));
        doc.RunAction("action");
        Assert.Equal("expired", doc.GetCellValue(0,0,0));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void LineReadingHonorsBlockBoundariesBomAndInvalidEncoding(bool utf16)
    {
        if (!OperatingSystem.IsWindows()) return;
        var firstLine = new string('x', utf16 ? 32766 : 65532);
        Encoding encoding = utf16 ? new UnicodeEncoding(true, true) : new UTF8Encoding(true);
        var file = Path.Combine(_dir, "lines");
        File.WriteAllText(file, firstLine + "\r\n界😀\n", encoding);
        using (var output = new FileStream(file, FileMode.Append)) output.WriteByte(0xff);
        var bom = Path.Combine(_dir, "bom");
        File.WriteAllBytes(bom, encoding.GetPreamble());
        using var doc = Open($$"""
            export async function main(w,h) {
              const [f,bom]=await h.fs.getUserSelectedFiles({mode:'open',fileMode:'open',fileAccess:'r',multiSelect:true});
              for (const options of [{buffering:'stream'},undefined]) {
                const s=await f.open(options), enc='{{(utf16 ? "utf-16be" : "utf-8")}}';
                if((await s.readLine(enc)).length!=={{firstLine.Length}} || await s.readLine(enc)!=='界😀') throw Error('split line');
                const pos=s.position;
                try { await s.readAllText(enc); throw Error('accepted invalid encoding'); } catch(e) {if(e.code!=='InvalidEncoding')throw e;}
                if(s.position!==pos) throw Error('advanced on failure');
                const bytes=new Uint8Array(2);
                if(await s.read(bytes,1,1)!==1 || bytes[1]!==255 || await s.readLine(enc)!==null) throw Error('last byte');
                await s.close();
                const empty=await bom.open(options);
                if(await empty.readLine(enc)!==null || empty.position!==empty.length) throw Error('BOM-only file');
                await empty.close();
              }
            }
            """, Select(file, bom));
        doc.RunAction("action");
    }

    [Fact]
    public void DirectoryMultiSelectAndInitializationPermissions()
    {
        if (!OperatingSystem.IsWindows()) return;
        var one = Directory.CreateDirectory(Path.Combine(_dir, "one")).FullName;
        var two = Directory.CreateDirectory(Path.Combine(_dir, "two")).FullName;
        int calls = 0;
        using var doc = Open("""
            export async function main(w,h) {
              const roots = await h.fs.getUserSelectedDirectory({multiSelect:true, title:'Read directories'});
              if(roots.length!==2 || !Object.isFrozen(roots)) throw Error('directory selection');
              const names=[];
              for (const root of roots) {
                names.push((await root.stat()).name);
                await root.release();
                try { await root.stat(); throw Error('released grant'); } catch(e) { if(e.code!=='InvalidReference') throw e; }
              }
              w.getActiveCell().setValue(names.join(','));
            }
            """, (request, begin, _) =>
            {
                using var interaction = begin();
                Assert.True(request.Directory && request.Multiple && !request.AllowDelete && request.Access == "r");
                Assert.Equal("Read directories", request.Title);
                calls++;
                return Task.FromResult(new ScriptFileSelection([one, two]));
            }, """
            export async function main(w,h) {
              try { await h.fs.getUserSelectedDirectory(); throw Error('initialization picked a directory'); }
              catch(e) { if(e.code!=='InvalidReference') throw e; }
            }
            """);
        Assert.Equal(0, calls);
        doc.RunAction("action");
        Assert.Equal(1, calls);
        Assert.Equal("one,two", doc.GetCellValue(0,0,0));
    }

    [Fact]
    public async Task CancellingAnActionWhilePickingDiscardsItsDraftAndLateSelection()
    {
        if (!OperatingSystem.IsWindows()) return;
        var output = Path.Combine(_dir, "pending");
        var entered = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var dismiss = new TaskCompletionSource<ScriptFileSelection>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var doc = Open("""
            export async function main(w,h) {
              const [file] = await h.fs.getUserSelectedFiles({mode:'save',fileMode:'create',fileAccess:'w'});
              const stream = await file.open(); stream.writeText('pending');
              await h.fs.getUserSelectedDirectory();
              await stream.close();
              w.getActiveCell().setValue('finished');
            }
            """, async (request, begin, _) =>
            {
                using var interaction = begin();
                if (request.Save) return new([output]);
                entered.TrySetResult();
                return await dismiss.Task;
            });
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        var action = Task.Run(() => doc.RunAction("action", cancellation.Token));
        try
        {
            await entered.Task.WaitAsync(cancellation.Token);
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => await action.WaitAsync(TimeSpan.FromSeconds(5)));
        }
        finally { dismiss.TrySetResult(new([_dir])); }
        Assert.False(File.Exists(output));
        Assert.Equal("", doc.GetCellValue(0,0,0));
    }
}

