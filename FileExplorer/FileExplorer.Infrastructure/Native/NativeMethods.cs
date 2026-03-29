// Copyright (c) FileExplorer contributors. All rights reserved.
// AOT Note: All P/Invoke signatures use static extern with LibraryImport where possible.
// For AOT, COM interfaces require [DynamicDependency] on the consuming service class.

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace FileExplorer.Infrastructure.Native;

/// <summary>
/// Win32 and COM P/Invoke declarations for file system operations, shell integration,
/// context menus, drag-drop, and Cloud Filter API.
/// </summary>
internal static partial class NativeMethods
{
    #region Kernel32 — File Operations

    internal const uint MOVEFILE_REPLACE_EXISTING = 0x00000001;
    internal const uint MOVEFILE_COPY_ALLOWED = 0x00000002;
    internal const uint MOVEFILE_WRITE_THROUGH = 0x00000008;

    [LibraryImport("kernel32.dll", EntryPoint = "MoveFileExW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool MoveFileEx(string lpExistingFileName, string lpNewFileName, uint dwFlags);

    [DllImport("kernel32.dll", EntryPoint = "GetFileAttributesExW", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool GetFileAttributesEx(string lpFileName, int fInfoLevelId, out WIN32_FILE_ATTRIBUTE_DATA lpFileInformation);

    [LibraryImport("kernel32.dll", EntryPoint = "CreateFileW", SetLastError = true, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial nint CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        nint lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        nint hTemplateFile);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool GetFileInformationByHandleEx(
        nint hFile,
        int fileInformationClass,
        nint lpFileInformation,
        uint dwBufferSize);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool CloseHandle(nint hObject);

    internal const uint FILE_READ_ATTRIBUTES = 0x80;
    internal const uint FILE_SHARE_READ = 0x00000001;
    internal const uint FILE_SHARE_WRITE = 0x00000002;
    internal const uint FILE_SHARE_DELETE = 0x00000004;
    internal const uint OPEN_EXISTING = 3;
    internal const uint FILE_FLAG_BACKUP_SEMANTICS = 0x02000000;
    internal const int FileAttributeTagInfo = 9;
    internal static readonly nint INVALID_HANDLE_VALUE = new(-1);

    [DllImport("kernel32.dll", EntryPoint = "FormatMessageW", SetLastError = true, CharSet = CharSet.Unicode)]
    internal static extern int FormatMessage(
        uint dwFlags,
        nint lpSource,
        uint dwMessageId,
        uint dwLanguageId,
        [Out] StringBuilder lpBuffer,
        int nSize,
        nint Arguments);

    internal const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;
    internal const uint FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;

    /// <summary>Formats a Win32 error code into a human-readable message.</summary>
    internal static string FormatWin32Error(int errorCode)
    {
        var sb = new StringBuilder(512);
        int len = FormatMessage(
            FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            nint.Zero, (uint)errorCode, 0, sb, sb.Capacity, nint.Zero);
        return len > 0 ? sb.ToString(0, len).TrimEnd('\r', '\n') : $"Unknown error (0x{errorCode:X8})";
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct WIN32_FILE_ATTRIBUTE_DATA
    {
        public FileAttributes dwFileAttributes;
        public FILETIME ftCreationTime;
        public FILETIME ftLastAccessTime;
        public FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;

        public readonly long FileSize => ((long)nFileSizeHigh << 32) | nFileSizeLow;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct FILE_ATTRIBUTE_TAG_INFO
    {
        public FileAttributes FileAttributes;
        public uint ReparseTag;
    }

    #endregion

    #region Shell32 — SHGetFileInfo

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint SHGetFileInfo(
        string pszPath,
        uint dwFileAttributes,
        ref SHFILEINFO psfi,
        uint cbFileInfo,
        uint uFlags);

    internal const uint SHGFI_ICON = 0x000000100;
    internal const uint SHGFI_SMALLICON = 0x000000001;
    internal const uint SHGFI_LARGEICON = 0x000000000;
    internal const uint SHGFI_TYPENAME = 0x000000400;
    internal const uint SHGFI_SYSICONINDEX = 0x000004000;
    internal const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;

    internal const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;
    internal const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct SHFILEINFO
    {
        public nint hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    [LibraryImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static partial bool DestroyIcon(nint hIcon);

    #endregion

    #region Shell32 — SHFileOperation (Recycle Bin)

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    internal static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);

    internal const int FO_MOVE = 0x0001;
    internal const int FO_COPY = 0x0002;
    internal const int FO_DELETE = 0x0003;
    internal const int FO_RENAME = 0x0004;

    internal const ushort FOF_MULTIDESTFILES = 0x0001;
    internal const ushort FOF_SILENT = 0x0004;
    internal const ushort FOF_NOCONFIRMATION = 0x0010;
    internal const ushort FOF_ALLOWUNDO = 0x0040;
    internal const ushort FOF_NOERRORUI = 0x0400;
    internal const ushort FOF_NOCONFIRMMKDIR = 0x0200;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct SHFILEOPSTRUCT
    {
        public nint hwnd;
        public int wFunc;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string pFrom;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pTo;
        public ushort fFlags;
        [MarshalAs(UnmanagedType.Bool)]
        public bool fAnyOperationsAborted;
        public nint hNameMappings;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpszProgressTitle;
    }

    #endregion

    #region Shell32 — Known Folders

    [DllImport("shell32.dll")]
    internal static extern int SHGetKnownFolderPath(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rfid,
        uint dwFlags,
        nint hToken,
        [MarshalAs(UnmanagedType.LPWStr)] out string ppszPath);

    internal static readonly Guid FOLDERID_Desktop = new("B4BFCC3A-DB2C-424C-B029-7FE99A87C641");
    internal static readonly Guid FOLDERID_Documents = new("FDD39AD0-238F-46AF-ADB4-6C85480369C7");
    internal static readonly Guid FOLDERID_Downloads = new("374DE290-123F-4565-9164-39C4925E467B");
    internal static readonly Guid FOLDERID_Music = new("4BD8D571-6D19-48D3-BE97-422220080E43");
    internal static readonly Guid FOLDERID_Pictures = new("33E28130-4E1E-4676-835A-98395C3BC3BB");
    internal static readonly Guid FOLDERID_Videos = new("18989B1D-99B5-455B-841C-AB7C74E4DDFC");
    internal static readonly Guid FOLDERID_Profile = new("5E6C858F-0E22-4760-9AFE-EA3317B67173");

    #endregion

    #region Shell32 — Context Menu COM Interfaces

    /// <summary>IShellFolder GUID.</summary>
    internal static readonly Guid IID_IShellFolder = new("000214E6-0000-0000-C000-000000000046");

    /// <summary>IContextMenu GUID.</summary>
    internal static readonly Guid IID_IContextMenu = new("000214E4-0000-0000-C000-000000000046");

    /// <summary>IContextMenu2 GUID.</summary>
    internal static readonly Guid IID_IContextMenu2 = new("000214F4-0000-0000-C000-000000000046");

    /// <summary>IContextMenu3 GUID.</summary>
    internal static readonly Guid IID_IContextMenu3 = new("BCFCE0A0-EC17-11D0-8D10-00A0C90F2719");

    [DllImport("shell32.dll")]
    internal static extern int SHGetDesktopFolder(out IShellFolder ppshf);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    internal static extern int SHParseDisplayName(
        string pszName,
        nint pbc,
        out nint ppidl,
        uint sfgaoIn,
        out uint psfgaoOut);

    [DllImport("shell32.dll")]
    internal static extern int SHBindToParent(
        nint pidl,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IShellFolder ppv,
        out nint ppidlLast);

    [DllImport("ole32.dll")]
    internal static extern void CoTaskMemFree(nint pv);

    [DllImport("user32.dll")]
    internal static extern nint CreatePopupMenu();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DestroyMenu(nint hMenu);

    [DllImport("user32.dll")]
    internal static extern int TrackPopupMenuEx(
        nint hMenu,
        uint uFlags,
        int x,
        int y,
        nint hwnd,
        nint lptpm);

    internal const uint TPM_RETURNCMD = 0x0100;
    internal const uint TPM_LEFTALIGN = 0x0000;

    /// <summary>IShellFolder COM interface.</summary>
    [ComImport, Guid("000214E6-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IShellFolder
    {
        void ParseDisplayName(nint hwnd, nint pbc, [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, out uint pchEaten, out nint ppidl, ref uint pdwAttributes);
        void EnumObjects(nint hwnd, uint grfFlags, out nint ppenumIDList);
        void BindToObject(nint pidl, nint pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out nint ppv);
        void BindToStorage(nint pidl, nint pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out nint ppv);
        [PreserveSig]
        int CompareIDs(nint lParam, nint pidl1, nint pidl2);
        void CreateViewObject(nint hwndOwner, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out nint ppv);
        void GetAttributesOf(uint cidl, [MarshalAs(UnmanagedType.LPArray)] nint[] apidl, ref uint rgfInOut);
        void GetUIObjectOf(nint hwndOwner, uint cidl, [MarshalAs(UnmanagedType.LPArray)] nint[] apidl, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, ref uint rgfReserved, out nint ppv);
        void GetDisplayNameOf(nint pidl, uint uFlags, out nint pName);
        void SetNameOf(nint hwnd, nint pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, uint uFlags, out nint ppidlOut);
    }

    /// <summary>IContextMenu COM interface.</summary>
    [ComImport, Guid("000214E4-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IContextMenu
    {
        [PreserveSig]
        int QueryContextMenu(nint hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
        void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);
        void GetCommandString(nuint idCmd, uint uType, nint pReserved, nint pszName, uint cchMax);
    }

    /// <summary>IContextMenu2 COM interface for owner-draw menu support.</summary>
    [ComImport, Guid("000214F4-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IContextMenu2
    {
        [PreserveSig]
        int QueryContextMenu(nint hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
        void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);
        void GetCommandString(nuint idCmd, uint uType, nint pReserved, nint pszName, uint cchMax);
        [PreserveSig]
        int HandleMenuMsg(uint uMsg, nint wParam, nint lParam);
    }

    /// <summary>IContextMenu3 COM interface for extended menu message handling.</summary>
    [ComImport, Guid("BCFCE0A0-EC17-11D0-8D10-00A0C90F2719"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IContextMenu3
    {
        [PreserveSig]
        int QueryContextMenu(nint hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
        void InvokeCommand(ref CMINVOKECOMMANDINFOEX pici);
        void GetCommandString(nuint idCmd, uint uType, nint pReserved, nint pszName, uint cchMax);
        [PreserveSig]
        int HandleMenuMsg(uint uMsg, nint wParam, nint lParam);
        [PreserveSig]
        int HandleMenuMsg2(uint uMsg, nint wParam, nint lParam, out nint plResult);
    }

    internal const uint CMF_NORMAL = 0x00000000;
    internal const uint CMF_EXPLORE = 0x00000004;
    internal const uint CMF_CANRENAME = 0x00000010;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct CMINVOKECOMMANDINFOEX
    {
        public int cbSize;
        public uint fMask;
        public nint hwnd;
        public nint lpVerb;
        [MarshalAs(UnmanagedType.LPStr)]
        public string? lpParameters;
        [MarshalAs(UnmanagedType.LPStr)]
        public string? lpDirectory;
        public int nShow;
        public uint dwHotKey;
        public nint hIcon;
        [MarshalAs(UnmanagedType.LPStr)]
        public string? lpTitle;
        public nint lpVerbW;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpParametersW;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpDirectoryW;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpTitleW;
        public POINT ptInvoke;
    }

    internal const uint CMIC_MASK_UNICODE = 0x00004000;
    internal const uint CMIC_MASK_PTINVOKE = 0x20000000;
    internal const int SW_SHOWNORMAL = 1;

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINT
    {
        public int X;
        public int Y;
    }

    #endregion

    #region Shell32 — ShellExecute

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern nint ShellExecute(
        nint hwnd,
        string lpOperation,
        string lpFile,
        string? lpParameters,
        string? lpDirectory,
        int nShowCmd);

    internal const int SW_SHOW = 5;

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SHObjectProperties(
        nint hwnd,
        uint shopObjectType,
        string pszObjectName,
        string? pszPropertyPage);

    internal const uint SHOP_PRINTERNAME = 0x00000001;
    internal const uint SHOP_FILEPATH = 0x00000002;
    internal const uint SHOP_VOLUMEGUID = 0x00000004;

    #endregion

    #region Cloud Filter API (cfapi)

    // AOT Note: CfApi functions are safe for AOT as they are plain P/Invoke with no COM dispatch.

    [DllImport("cldapi.dll", CharSet = CharSet.Unicode)]
    internal static extern int CfGetPlaceholderStateFromFileInfo(
        nint InfoBuffer,
        int InfoClass);

    [DllImport("cldapi.dll", CharSet = CharSet.Unicode)]
    internal static extern int CfGetPlaceholderInfo(
        nint FileHandle,
        int InfoClass,
        nint InfoBuffer,
        int InfoBufferLength,
        out int ReturnedLength);

    /// <summary>Opens a placeholder file handle for cloud filter operations.</summary>
    [DllImport("cldapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int CfOpenFileWithOplock(
        string FilePath,
        int Flags,
        out nint ProtectedHandle);

    [DllImport("cldapi.dll")]
    internal static extern void CfCloseHandle(nint ProtectedHandle);

    /// <summary>Hydrates (downloads) a cloud-only file.</summary>
    [DllImport("cldapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int CfHydratePlaceholder(
        nint FileHandle,
        long StartingOffset,
        long Length,
        int HydrateFlags,
        nint Overlapped);

    /// <summary>Dehydrates a locally available file back to cloud-only.</summary>
    [DllImport("cldapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int CfDehydratePlaceholder(
        nint FileHandle,
        long StartingOffset,
        long Length,
        int DehydrateFlags,
        nint Overlapped);

    /// <summary>Sets the pin state of a placeholder file.</summary>
    [DllImport("cldapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern int CfSetPinState(
        nint FileHandle,
        int PinState,
        int PinFlags,
        nint Overlapped);

    /// <summary>Gets sync root info to detect OneDrive paths.</summary>
    [DllImport("cldapi.dll", CharSet = CharSet.Unicode)]
    internal static extern int CfGetSyncRootInfoByPath(
        string FilePath,
        int InfoClass,
        nint InfoBuffer,
        int InfoBufferLength,
        out int ReturnedLength);

    internal const int CF_PIN_STATE_PINNED = 1;
    internal const int CF_PIN_STATE_UNPINNED = 2;
    internal const int CF_PIN_STATE_UNSPECIFIED = 0;

    internal const int CF_PLACEHOLDER_STATE_NO_STATES = 0;
    internal const int CF_PLACEHOLDER_STATE_PLACEHOLDER = 1;
    internal const int CF_PLACEHOLDER_STATE_SYNC_ROOT = 2;
    internal const int CF_PLACEHOLDER_STATE_ESSENTIAL_PROP_PRESENT = 4;
    internal const int CF_PLACEHOLDER_STATE_IN_SYNC = 8;
    internal const int CF_PLACEHOLDER_STATE_PARTIAL = 16;
    internal const int CF_PLACEHOLDER_STATE_PARTIALLY_ON_DISK = 32;

    #endregion

    #region Drag and Drop COM Interfaces

    /// <summary>IDropTarget COM interface for accepting drag-and-drop operations.</summary>
    [ComImport, Guid("00000122-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDropTarget
    {
        [PreserveSig]
        int DragEnter(IDataObject pDataObj, uint grfKeyState, POINT pt, ref uint pdwEffect);
        [PreserveSig]
        int DragOver(uint grfKeyState, POINT pt, ref uint pdwEffect);
        [PreserveSig]
        int DragLeave();
        [PreserveSig]
        int Drop(IDataObject pDataObj, uint grfKeyState, POINT pt, ref uint pdwEffect);
    }

    internal const uint DROPEFFECT_NONE = 0;
    internal const uint DROPEFFECT_COPY = 1;
    internal const uint DROPEFFECT_MOVE = 2;
    internal const uint DROPEFFECT_LINK = 4;

    /// <summary>IDragSourceHelper for producing standard Windows drag ghost images.</summary>
    [ComImport, Guid("DE5BF786-477A-11D2-839D-00C04FD918D0"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDragSourceHelper
    {
        void InitializeFromBitmap(ref SHDRAGIMAGE pshdi, IDataObject pDataObject);
        void InitializeFromWindow(nint hwnd, ref POINT ppt, IDataObject pDataObject);
    }

    /// <summary>IDropTargetHelper for drawing drag images during drop operations.</summary>
    [ComImport, Guid("4657278B-411B-11D2-839A-00C04FD918D0"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDropTargetHelper
    {
        void DragEnter(nint hwndTarget, IDataObject pDataObject, ref POINT ppt, uint dwEffect);
        void DragLeave();
        void DragOver(ref POINT ppt, uint dwEffect);
        void Drop(IDataObject pDataObject, ref POINT ppt, uint dwEffect);
        void Show([MarshalAs(UnmanagedType.Bool)] bool fShow);
    }

    internal static readonly Guid CLSID_DragDropHelper = new("4657278A-411B-11D2-839A-00C04FD918D0");

    [StructLayout(LayoutKind.Sequential)]
    internal struct SHDRAGIMAGE
    {
        public SIZE sizeDragImage;
        public POINT ptOffset;
        public nint hbmpDragImage;
        public uint crColorKey;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SIZE
    {
        public int cx;
        public int cy;
    }

    #endregion

    #region Long Path Support

    /// <summary>Ensures a path uses the \\?\ prefix for long path support.</summary>
    internal static string EnsureLongPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return path;
        if (path.StartsWith(@"\\?\", StringComparison.Ordinal)) return path;
        if (path.StartsWith(@"\\", StringComparison.Ordinal))
            return @"\\?\UNC\" + path[2..];
        return @"\\?\" + path;
    }

    #endregion

    #region Window Message Constants

    internal const uint WM_COMMAND = 0x0111;
    internal const uint WM_DRAWITEM = 0x002B;
    internal const uint WM_MEASUREITEM = 0x002C;
    internal const uint WM_INITMENUPOPUP = 0x0117;
    internal const uint WM_MENUCHAR = 0x0120;

    #endregion
}
