param([Parameter(Mandatory)][string]$OutputDirectory,[string]$WorkbookPath,[string]$OraclePath)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent;$target=[IO.Path]::GetFullPath($OutputDirectory,$repo)
$root=[IO.Path]::GetFullPath((Join-Path $repo 'artifacts'))+[IO.Path]::DirectorySeparatorChar
if(-not $target.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $target)){throw 'Use a new artifacts directory.'}
[void](New-Item -ItemType Directory -Path $target)
function Release-Com($o){if($null -ne $o){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}}
function Capture($book){
    $records=[Collections.Generic.List[object]]::new()
    for($s=1;$s -le $book.Worksheets.Count;$s++){
        $sheet=$book.Worksheets.Item($s);$comments=$sheet.CommentsThreaded
        try{for($i=1;$i -le $comments.Count;$i++){
            $comment=$comments.Item($i);$cell=$comment.Parent;$replies=$comment.Replies;$texts=[Collections.Generic.List[string]]::new()
            try{for($j=1;$j -le $replies.Count;$j++){$reply=$replies.Item($j);try{$texts.Add([string]$reply.Text())}finally{Release-Com $reply}}
                $records.Add([pscustomobject]@{Sheet=[string]$sheet.Name;Cell=[string]$cell.Address($false,$false);Text=[string]$comment.Text();Replies=$texts.ToArray()})
            }finally{Release-Com $replies;Release-Com $cell;Release-Com $comment}
        }}finally{Release-Com $comments;Release-Com $sheet}
    }
    return ,@($records|Sort-Object Sheet,Cell)
}
$excel=$null;$book=$null
try{
    $excel=New-Object -ComObject Excel.Application;$excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.AskToUpdateLinks=$false;$excel.AutomationSecurity=3;$excel.EnableEvents=$false
    if(-not ('ThreadExcelOwner' -as [type])){Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class ThreadExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process); }'}
    [uint32]$owned=0;[void][ThreadExcelOwner]::GetWindowThreadProcessId([intptr]$excel.Hwnd,[ref]$owned)
    @{ProcessId=$owned;Started=(Get-Process -Id $owned).StartTime;Owner='test-excel-threaded-comments.ps1'}|ConvertTo-Json|Set-Content (Join-Path $target 'excel-owner.json') -Encoding utf8
    if($WorkbookPath){
        $source=[IO.Path]::GetFullPath($WorkbookPath,$repo)
        if(-not $source.StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or -not $OraclePath){throw 'Verification requires an artifacts source and oracle.'}
        $hash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash;$book=$excel.Workbooks.Open($source,0,$true)
    }else{
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        try{$sheet.Name='Data';$cell=$sheet.Range('A1');try{$cell.Value2=10.0;$comment=$cell.AddCommentThreaded('Fixture root');try{$reply=$comment.AddReply('Fixture reply');Release-Com $reply}finally{Release-Com $comment}}finally{Release-Com $cell}
            $cell=$sheet.Range('C3');try{$cell.Value2='Unchanged';$comment=$cell.AddCommentThreaded('Second local thread');Release-Com $comment}finally{Release-Com $cell}
        }finally{Release-Com $sheet}
    }
    $before=Capture $book;$path=Join-Path $target 'threads.xlsx';$book.SaveAs($path,51);$book.Close($false);Release-Com $book;$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$after=Capture $book;$book.Close($false);Release-Com $book;$book=$null
    $stable=($before|ConvertTo-Json -Depth 7 -Compress) -ceq ($after|ConvertTo-Json -Depth 7 -Compress)
    $matches=if($WorkbookPath){$oracle=Get-Content -LiteralPath $OraclePath -Raw|ConvertFrom-Json;($oracle.Comments|ConvertTo-Json -Depth 7 -Compress) -ceq ($after|ConvertTo-Json -Depth 7 -Compress)}else{$true}
    $unchanged=-not $WorkbookPath -or $hash -eq (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    [ordered]@{Version=[string]$excel.Version;Build=[string]$excel.Build;Comments=$after;Stable=$stable;ReferenceMatches=$matches;SourceUnchanged=$unchanged;SourceSha256=$hash;OutputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}|
        ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $target 'threaded-comments.json') -Encoding utf8
    if(-not $stable -or -not $matches -or -not $unchanged){throw 'Threaded comment verification failed.'}
    Write-Output "PASS: $($after.Count) local threaded comments and replies; save/reopen stable."
}finally{if($book){$book.Close($false);Release-Com $book};[GC]::Collect();[GC]::WaitForPendingFinalizers();if($excel){$excel.Quit();Release-Com $excel}}
