// Development-only Win32 clipboard capture/publication for synthetic interoperability evidence.
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class ClipboardNative
{
    [DllImport("user32.dll")] static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")] static extern bool CloseClipboard();
    [DllImport("user32.dll")] static extern uint EnumClipboardFormats(uint format);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClipboardFormatName(uint format, StringBuilder name, int count);
    [DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint format);
    [DllImport("kernel32.dll")] static extern UIntPtr GlobalSize(IntPtr handle);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr handle);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr handle);
    [DllImport("user32.dll")] static extern bool EmptyClipboard();
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern uint RegisterClipboardFormat(string name);
    [DllImport("user32.dll")] static extern IntPtr SetClipboardData(uint format, IntPtr handle);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalAlloc(uint flags, UIntPtr size);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalFree(IntPtr handle);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr CreateWindowEx(uint ex, string cls, string title, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
    [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr window);
    static void Open(IntPtr owner) { for(int i=0;i<20;i++) { if(OpenClipboard(owner)) return; System.Threading.Thread.Sleep(25); } throw new InvalidOperationException("Clipboard busy"); }
    public static Dictionary<string, byte[]> Read()
    {
        var result = new Dictionary<string,byte[]>(); Open(IntPtr.Zero);
        try { uint f=0; while((f=EnumClipboardFormats(f))!=0) { var s=new StringBuilder(256); GetClipboardFormatName(f,s,256); string name=s.Length==0 ? "CF_"+f : s.ToString();
            // Never dereference bitmap/metafile/OLE handles; only capture bounded HGLOBAL text formats.
            if(name!="XML Spreadsheet" && name!="HTML Format" && name!="Biff8" && f!=13) { result[name]=new byte[0]; continue; }
            var h=GetClipboardData(f); ulong size=GlobalSize(h).ToUInt64(); if(size>16*1024*1024) throw new InvalidOperationException("Payload too large");
            var p=GlobalLock(h); if(p==IntPtr.Zero) continue; try { var b=new byte[(int)size]; Marshal.Copy(p,b,0,b.Length); result[name]=b; } finally { GlobalUnlock(h); }
        } } finally { CloseClipboard(); } return result;
    }
    public static void Write(string[] names, byte[][] payloads)
    {
        var window=CreateWindowEx(0,"STATIC","QuickExcel synthetic clipboard probe",0,0,0,0,0,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
        if(window==IntPtr.Zero) throw new InvalidOperationException("Clipboard owner failed");
        try { Open(window); try { if(!EmptyClipboard()) throw new InvalidOperationException("EmptyClipboard failed");
            for(int i=0;i<names.Length;i++) { var h=GlobalAlloc(0x42,(UIntPtr)payloads[i].Length); var p=GlobalLock(h); if(p==IntPtr.Zero) throw new InvalidOperationException("Allocation failed");
                Marshal.Copy(payloads[i],0,p,payloads[i].Length); GlobalUnlock(h);
                if(SetClipboardData(names[i]=="CF_13"?13:RegisterClipboardFormat(names[i]),h)==IntPtr.Zero) { GlobalFree(h); throw new InvalidOperationException("SetClipboardData failed"); }
            }
        } finally { CloseClipboard(); } } finally { DestroyWindow(window); }
    }
}
