using Avalonia.Controls;
using Avalonia.Threading;
using QuickExcel.Core;
using QuickExcel.Storage;

namespace QuickExcel.Desktop.Avalonia;

public sealed partial class MainWindow
{
    private WorkbookRecoveryStore _recoveryStore = null!;
    private WorkbookRecoveryStore.RecoveryLease? _recoveryLease;
    private readonly DispatcherTimer _recoveryTimer = new() { Interval = TimeSpan.FromSeconds(30) };
    private readonly TextBlock _recoveryStatus = new() { Margin = new(8, 5), MaxWidth = 360 };
    private Guid _recoverySession;
    private long _recoveryRevision = -1;
    private Task _recoveryWrite = Task.CompletedTask;
    private RecoveryMetadata? _lastRecovery;
    private RecoveryCandidate? _recoveredFrom;
    private bool _recoveryEnabled, _recoveryStopping, _recoveryAgain;

    private void InstallRecovery()
    {
        _recoveryStore = new(Path.Combine(_userDataDirectory, "recovery"));
        var menu = new MenuItem { Header = UiText.Get("自动恢复草稿…") };
        menu.Click += async (_, _) => await PerformAsync(() => ShowRecoveryAsync(false)); ((MenuItem)_menu.Items[0]!).Items.Add(menu);
        ToolTip.SetTip(_recoveryStatus, UiText.Get("每 30 秒保存已提交的工作簿编辑。单元格、公式栏、富文本窗口和脚本编辑器中尚未确认的输入不包含在内。"));
        _recoveryTimer.Tick += async (_, _) => await CaptureRecoveryAsync();
        Closed += (_, _) => { _recoveryTimer.Stop(); _recoveryLease?.Dispose(); };
    }
    private void StartRecovery() { _recoveryEnabled = true; _recoveryTimer.Start(); }
    private Task CaptureRecoveryAsync()
    {
        if (_recoveryStopping || !_book.IsVisible || Document.IsReadOnly || !Document.IsDirty) return Task.CompletedTask;
        if (!_recoveryWrite.IsCompleted) { _recoveryAgain = true; return _recoveryWrite; }
        _recoveryWrite = WriteRecoveryAsync(); return _recoveryWrite;
    }
    private async Task WriteRecoveryAsync()
    {
        do
        {
            _recoveryAgain = false;
            if (_recoveryStopping || !Document.IsDirty) break;
            var snapshot = Document.Snapshot; var session = Document.Session.SessionId;
            var source = _workspace.FilePath; var version = _workspace.Version; var captured = DateTimeOffset.UtcNow;
            if (_recoverySession == session && _recoveryRevision == snapshot.Revision) break;
            try
            {
                if (_recoverySession != session)
                {
                    // An unexpected document replacement preserves the previous orphan for recovery.
                    _recoveryLease?.Dispose(); _recoveryLease = null; _recoveryRevision = -1; _lastRecovery = null;
                }
                _recoverySession = session;
                var lease = _recoveryLease ??= await Task.Run(() => _recoveryStore.Start(session, snapshot.DocumentId));
                _recoveryStatus.Text = UiText.Get("正在写入恢复草稿…");
                var completed = await Task.Run(() => lease.Capture(snapshot, source, version, captured));
                if (Document.Session.SessionId != session) continue;
                _recoveryRevision = snapshot.Revision; _lastRecovery = completed;
                _recoveryStatus.Text = UiText.Get("恢复草稿：") + completed.CapturedUtc.ToLocalTime().ToString("HH:mm:ss") +
                    (Document.Snapshot.Revision != snapshot.Revision || DateTimeOffset.UtcNow - captured > TimeSpan.FromSeconds(60) ? UiText.Get(" · 有未记录编辑") : "");
            }
            catch (Exception ex)
            {
                _recoveryStatus.Text = UiText.Get("恢复草稿写入失败：") + ex.Message;
                ToolTip.SetTip(_recoveryStatus, _recoveryStatus.Text + "\n" + (_lastRecovery?.CapturedUtc.ToLocalTime().ToString("O") ?? UiText.Get("尚无完整恢复副本")));
                break;
            }
        } while (_recoveryAgain && !_recoveryStopping);
    }
    private async Task ClearRecoveryAsync()
    {
        _recoveryStopping = true; _recoveryTimer.Stop();
        try
        {
            await _recoveryWrite;
            var lease = _recoveryLease; _recoveryLease = null;
            if (lease is not null) { try { await Task.Run(lease.Discard); } finally { lease.Dispose(); } }
            if (_recoveredFrom is { } source) { await Task.Run(() => _recoveryStore.Discard(source)); _recoveredFrom = null; }
            _recoverySession = Guid.Empty; _recoveryRevision = -1; _lastRecovery = null; _recoveryStatus.Text = "";
        }
        catch (Exception ex) { _recoveryStatus.Text = UiText.Get("旧恢复副本清理失败：") + ex.Message; }
        finally { _recoveryStopping = false; if (_recoveryEnabled && !_windowClosed) _recoveryTimer.Start(); }
    }
    private async Task ShowRecoveryAsync(bool startup)
    {
        RecoveryScan scan;
        try { scan = await Task.Run(_recoveryStore.Scan); }
        catch (Exception ex) { _recoveryStatus.Text = UiText.Get("恢复列表读取失败：") + ex.Message; return; }
        if (scan.Items.Count == 0)
        {
            if (!startup) await new FieldsDialog(UiText.Get("自动恢复草稿")).Note(UiText.Get("没有可恢复的离线草稿。其他运行实例的草稿不会显示。") + "\n" + string.Join("\n", scan.Notices)).ShowAsync(this);
            else if (scan.Notices.Count > 0) _recoveryStatus.Text = UiText.Get("部分恢复副本不可读取，请查看恢复列表。");
            return;
        }
        var dialog = new FieldsDialog(UiText.Get("自动恢复草稿"))
            .Note(UiText.Get("恢复到新文档并要求另存，脚本保持停用。仅恢复已提交编辑；所列时间之后及编辑器内尚未确认的输入可能丢失。原文件不会被覆盖。"))
            .Choice("item", UiText.Get("选择恢复副本"), scan.Items.Select(c => (c.Metadata.SourcePath ?? UiText.Get("未命名")) + " · " + c.Metadata.CapturedUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + " · r" + c.Metadata.Revision + (c.DamagedGeneration ? UiText.Get(" · 已回退到完整副本") : "")).ToArray())
            .Check("remove", UiText.Get("永久删除所选会话的恢复草稿"));
        if (scan.Notices.Count > 0) dialog.Note(string.Join("\n", scan.Notices));
        if (!await dialog.ShowAsync(this)) return;
        var candidate = scan.Items[dialog.Index("item")];
        if (dialog.Boolean("remove")) { await Task.Run(() => _recoveryStore.Discard(candidate)); return; }
        var snapshot = await Task.Run(() => _recoveryStore.Recover(candidate));
        if (!await ConfirmDiscardAsync()) return;
        _workspace.Recover(snapshot); _recoveredFrom = candidate; ShowBook();
        _status.Text = UiText.Get("已恢复为未保存的新文档；请另存。脚本未授权。");
        if (candidate.Metadata.SourcePath is { } path)
        {
            try
            {
                var current = await Task.Run(() => FileVersions.Read(path));
                if (current.Sha256 != candidate.Metadata.SourceSha256) _status.Text += UiText.Get(" 原文件已改变或不存在，请核对恢复内容。");
            }
            catch (IOException) { _status.Text += UiText.Get(" 无法核对原文件版本，请另存后核对。"); }
            catch (UnauthorizedAccessException) { _status.Text += UiText.Get(" 无法核对原文件版本，请另存后核对。"); }
        }
        await CaptureRecoveryAsync();
    }
}
