using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using QuickExcel.Core;

namespace QuickExcel.Printing;

public sealed record CupsDestination(string Name,string? Instance,bool IsDefault)
{
    public string DisplayName=>Instance is null?Name:Name+"/"+Instance;
}
public sealed record CupsMedia(string Name,int Width,int Height,int Left,int Right,int Top,int Bottom,bool Landscape);

/// <summary>System CUPS capability query and bounded PDF spooling; does not own the PDF file.</summary>
public static class CupsPrintService
{
    private static readonly Lazy<Api> Native=new(()=>new Api());
    public static async Task<IReadOnlyList<CupsDestination>> GetDestinationsAsync(CancellationToken cancellation=default)
        =>await Task.Run(()=>WithDestinations((api,http,count,dests)=>
        {
            var result=new List<CupsDestination>();
            for(var i=0;i<count;i++)
            {
                var dest=Marshal.PtrToStructure<Destination>(dests+i*Marshal.SizeOf<Destination>());
                var name=Marshal.PtrToStringUTF8(dest.Name);var instance=Marshal.PtrToStringUTF8(dest.Instance);
                if(name is null || name.Length>127 || name.Any(char.IsControl) || instance?.Length>127)throw new InvalidDataException("打印机名称无效。");
                result.Add(new(name,instance,dest.IsDefault!=0));
            }
            return (IReadOnlyList<CupsDestination>)result;
        },cancellation),cancellation).ConfigureAwait(false);

    public static Task<CupsMedia> ValidateMediaAsync(CupsDestination destination,PrintLayout layout,CancellationToken cancellation=default)
        =>Task.Run(()=>WithDestinations((api,http,count,dests)=>
        {
            var dest=api.GetDest(destination.Name,destination.Instance,count,dests);
            if(dest==0)throw new IOException("所选打印机已不可用。");
            var info=api.CopyInfo(http,dest);if(info==0)throw new IOException("无法读取打印机纸张与硬边距，任务未提交。");
            try
            {
                var countMedia=api.MediaCount(http,dest,info,0);
                if(countMedia is <1 or >4096)throw new IOException("打印机没有提供有界的纸张能力列表。");
                var width=(int)Math.Round(layout.PaperWidth/96*2540);var height=(int)Math.Round(layout.PaperHeight/96*2540);
                CupsMedia? selected=null;
                for(var i=0;i<countMedia;i++)
                {
                    cancellation.ThrowIfCancellationRequested();
                    if(api.MediaByIndex(http,dest,info,i,0,out var size)==0)continue;
                    var normal=Math.Abs(size.Width-width)<=1 && Math.Abs(size.Length-height)<=1;
                    var landscape=Math.Abs(size.Width-height)<=1 && Math.Abs(size.Length-width)<=1;
                    if(!normal && !landscape)continue;
                    if(size.Width<=0 || size.Length<=0 || new[]{size.Left,size.Right,size.Top,size.Bottom}.Any(m=>m<0))continue;
                    var margin=landscape&&!normal?Math.Max(Math.Max(size.Left,size.Right),Math.Max(size.Top,size.Bottom)):-1;
                    var media=new CupsMedia(size.Media,size.Width,size.Length,margin<0?size.Left:margin,margin<0?size.Right:margin,margin<0?size.Top:margin,margin<0?size.Bottom:margin,!normal);
                    if(layout.Margins.Left*2540+1<media.Left || layout.Margins.Right*2540+1<media.Right || layout.Margins.Top*2540+1<media.Top || layout.Margins.Bottom*2540+1<media.Bottom)continue;
                    selected=media;break;
                }
                return selected ?? throw new InvalidOperationException("打印机纸张或不可打印边缘与当前预览不匹配；请调整页面设置并重新生成预览。");
            }
            finally{api.FreeInfo(info);}
        },cancellation),cancellation);

    public static string NormalizePageRanges(string text,int pageCount)
    {
        if(pageCount is <1 or >PrintingLayout.MaxPages || text.Length>256)throw new ArgumentException("打印页范围无效。");
        if(string.IsNullOrWhiteSpace(text))return "1-"+pageCount.ToString(CultureInfo.InvariantCulture);
        var pages=new SortedSet<int>();var ranges=text.Split(',');
        if(ranges.Length>16)throw new ArgumentException("最多输入 16 个页范围。");
        foreach(var range in ranges)
        {
            var parts=range.Trim().Split('-');
            if(parts.Length is <1 or >2 || !int.TryParse(parts[0],NumberStyles.None,CultureInfo.InvariantCulture,out var first))throw new ArgumentException("页范围示例：1-3,5。");
            var last=first;
            if(parts.Length==2 && !int.TryParse(parts[1],NumberStyles.None,CultureInfo.InvariantCulture,out last))throw new ArgumentException("页范围无效。");
            if(first<1 || last<first || last>pageCount)throw new ArgumentException("页范围超出当前预览。");
            for(var page=first;page<=last;page++)pages.Add(page);
        }
        var normalized=new List<string>();var ordered=pages.ToArray();
        for(var i=0;i<ordered.Length;i++){var first=ordered[i];while(i+1<ordered.Length && ordered[i+1]==ordered[i]+1)i++;normalized.Add(first==ordered[i]?first.ToString(CultureInfo.InvariantCulture):$"{first}-{ordered[i]}");}
        return string.Join(',',normalized);
    }
    public static async Task<string> SubmitAsync(CupsDestination destination,CupsMedia media,string pdfPath,string pageRanges,int copies,CancellationToken cancellation=default)
    {
        if(!OperatingSystem.IsLinux()&&!OperatingSystem.IsMacOS())throw new PlatformNotSupportedException();
        if(copies is <1 or >999 || !Path.IsPathFullyQualified(pdfPath) || !File.Exists(pdfPath) || !File.Exists("/usr/bin/lp"))throw new ArgumentException("打印参数、PDF 或系统 CUPS 客户端不可用。");
        if(string.IsNullOrEmpty(media.Name)||media.Name.Length>127||media.Name.Any(char.IsControl))throw new InvalidDataException("打印机纸张标识无效。");
        pageRanges=NormalizePageRanges(pageRanges,PrintingLayout.MaxPages);
        var start=new ProcessStartInfo("/usr/bin/lp"){UseShellExecute=false,RedirectStandardOutput=true,RedirectStandardError=true};
        start.Environment["LC_ALL"]="C";
        string[] args=["-d",destination.DisplayName,"-n",copies.ToString(CultureInfo.InvariantCulture),"-t","QuickExcel", "-o","media="+media.Name,
            "-o","document-format=application/pdf","-o","print-scaling=none","-o","fit-to-page=false","-o","number-up=1","-o","sides=one-sided","-o","job-sheets=none",
            "-o","orientation-requested="+(media.Landscape?"4":"3"),"-o","page-ranges="+pageRanges,"--",pdfPath];
        foreach(var argument in args)start.ArgumentList.Add(argument);
        using var timeout=CancellationTokenSource.CreateLinkedTokenSource(cancellation);timeout.CancelAfter(TimeSpan.FromSeconds(30));
        cancellation.ThrowIfCancellationRequested();
        using var process=Process.Start(start)??throw new IOException("无法启动系统打印客户端。");
        try
        {
            var output=process.StandardOutput.ReadToEndAsync(timeout.Token);var error=process.StandardError.ReadToEndAsync(timeout.Token);
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
            var response=(await output.ConfigureAwait(false)).Trim();var failure=(await error.ConfigureAwait(false)).Trim();
            if(process.ExitCode!=0)throw new IOException("系统打印服务拒绝任务："+failure[..Math.Min(failure.Length,2048)]);
            if(response.Length>2048 || !response.StartsWith("request id is ",StringComparison.Ordinal))throw new IOException("打印服务响应未能确认任务编号，请检查系统队列；不会自动重试。");
            return response;
        }
        catch(OperationCanceledException ex){throw new IOException("打印提交已中断；任务可能已进入系统队列，请先检查队列再决定是否重试。",ex);}
        finally{if(!process.HasExited){process.Kill();await process.WaitForExitAsync().ConfigureAwait(false);}}
    }
    private static T WithDestinations<T>(Func<Api,nint,int,nint,T> action,CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();var api=Native.Value;var cancel=0;
        var server=Marshal.PtrToStringUTF8(api.Server())??throw new IOException("CUPS server 未配置。");
        var http=api.Connect(server,api.Port(),0,0,api.Encryption(),1,3000,ref cancel);
        if(http==0)throw new IOException("无法连接 CUPS 打印服务。请检查系统打印设置。");
        try
        {
            api.Timeout(http,3,0,0);var count=api.GetDests(http,out var dests);
            try{if(count is <0 or >1024)throw new InvalidDataException("打印机列表超过限制。");cancellation.ThrowIfCancellationRequested();return action(api,http,count,dests);}
            finally{if(dests!=0)api.FreeDests(count,dests);}
        }
        finally{api.Close(http);}
    }
    [StructLayout(LayoutKind.Sequential)] private struct Destination {public nint Name,Instance;public int IsDefault,OptionsCount;public nint Options;}
    [StructLayout(LayoutKind.Sequential,CharSet=CharSet.Ansi)] private struct Size { [MarshalAs(UnmanagedType.ByValTStr,SizeConst=128)] public string Media;public int Width,Length,Bottom,Left,Right,Top; }
    private sealed class Api
    {
        private readonly nint _library;
        internal readonly StringValue Server;internal readonly IntegerValue Port,Encryption;internal readonly HttpConnect Connect;internal readonly HttpTimeout Timeout;internal readonly Release Close,FreeInfo;
        internal readonly GetDestinations GetDests;internal readonly FreeDestinations FreeDests;internal readonly GetDestination GetDest;internal readonly CopyDestinationInfo CopyInfo;internal readonly GetMediaCount MediaCount;internal readonly GetMediaByIndex MediaByIndex;
        internal Api()
        {
            if(!OperatingSystem.IsLinux()&&!OperatingSystem.IsMacOS())throw new PlatformNotSupportedException();
            _library=NativeLibrary.Load(OperatingSystem.IsMacOS()?"/usr/lib/libcups.2.dylib":"libcups.so.2");
            T Load<T>(string name) where T:Delegate=>Marshal.GetDelegateForFunctionPointer<T>(NativeLibrary.GetExport(_library,name));
            Server=Load<StringValue>("cupsServer");Port=Load<IntegerValue>("ippPort");Encryption=Load<IntegerValue>("cupsEncryption");Connect=Load<HttpConnect>("httpConnect2");Timeout=Load<HttpTimeout>("httpSetTimeout");Close=Load<Release>("httpClose");
            GetDests=Load<GetDestinations>("cupsGetDests2");FreeDests=Load<FreeDestinations>("cupsFreeDests");GetDest=Load<GetDestination>("cupsGetDest");CopyInfo=Load<CopyDestinationInfo>("cupsCopyDestInfo");FreeInfo=Load<Release>("cupsFreeDestInfo");MediaCount=Load<GetMediaCount>("cupsGetDestMediaCount");MediaByIndex=Load<GetMediaByIndex>("cupsGetDestMediaByIndex");
        }
    }
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate nint StringValue();
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate int IntegerValue();
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate nint HttpConnect([MarshalAs(UnmanagedType.LPUTF8Str)]string host,int port,nint addresses,int family,int encryption,int blocking,int milliseconds,ref int cancel);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate void HttpTimeout(nint http,double seconds,nint callback,nint context);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate void Release(nint value);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate int GetDestinations(nint http,out nint destinations);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate void FreeDestinations(int count,nint destinations);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate nint GetDestination([MarshalAs(UnmanagedType.LPUTF8Str)]string name,[MarshalAs(UnmanagedType.LPUTF8Str)]string? instance,int count,nint destinations);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate nint CopyDestinationInfo(nint http,nint destination);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate int GetMediaCount(nint http,nint destination,nint info,uint flags);
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)] private delegate int GetMediaByIndex(nint http,nint destination,nint info,int index,uint flags,out Size size);
}
