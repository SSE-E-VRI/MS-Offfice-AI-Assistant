using System;
using System.Runtime.InteropServices;

namespace MistralOfficeAddin
{
    // ---------------------------------------------------------------------
    // Hand-declared COM interop interfaces (no Office PIAs required).
    //
    // GUIDs verified against the Office Core type library (MSO.DLL) and
    // the Extensibility type library.  InterfaceType set to match each
    // interface's real nature:
    //   - IDTExtensibility2 is IUnknown-based (vtable, not IDispatch)
    //   - Office Core interfaces (Ribbon, CTP) are IDispatch-based
    // ---------------------------------------------------------------------

    public enum ext_ConnectMode
    {
        ext_cm_CommandLine = 0,
        ext_cm_External = 1,
        ext_cm_Solution = 2,
        ext_cm_Startup = 3,
        ext_cm_AfterStartup = 4
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UIClosed = 2
    }

    /// <summary>
    /// Extensibility.IDTExtensibility2 — the shared COM add-in entry interface.
    /// This is a vtable (IUnknown-based) interface, NOT IDispatch.
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnDisconnection(
            ext_DisconnectMode RemoveMode,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnAddInsUpdate(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnStartupComplete(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnBeginShutdown(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>
    /// Microsoft.Office.Core.IRibbonExtensibility — supplies custom Ribbon XML.
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    /// <summary>
    /// Microsoft.Office.Core.ICustomTaskPaneConsumer — the host hands the CTP
    /// factory to the add-in through this interface at startup.
    /// GUID: 000C033E (verified against Office type library).
    /// </summary>
    [ComImport]
    [Guid("000C033E-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICustomTaskPaneConsumer
    {
        [DispId(1)]
        void CTPFactoryAvailable([MarshalAs(UnmanagedType.IDispatch)] object CTPFactoryInst);
    }

    /// <summary>
    /// Microsoft.Office.Core.ICTPFactory — creates custom task panes.
    /// GUID: 000C033D (verified against Office type library).
    /// </summary>
    [ComImport]
    [Guid("000C033D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICTPFactory
    {
        [DispId(1)]
        object CreateCTP(
            [MarshalAs(UnmanagedType.BStr)] string CTPAxID,
            [MarshalAs(UnmanagedType.BStr)] string CTPTitle,
            [MarshalAs(UnmanagedType.Struct)] object CTPParentWindow);
    }

    /// <summary>
    /// Microsoft.Office.Core._CustomTaskPane — the task pane object returned
    /// by ICTPFactory.CreateCTP. Member order matches the PIA exactly:
    /// Title, Application, Window, Visible, ContentControl, Height, Width,
    /// DockPosition, DockPositionRestrict, Delete.
    /// DockPosition values (MsoCTPDockPosition): 0=Left 1=Top 2=Right 3=Bottom 4=Floating.
    /// </summary>
    [ComImport]
    [Guid("000C033B-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _CustomTaskPane
    {
        [DispId(1)]
        string Title { get; }
        [DispId(2)]
        object Application { get; }
        [DispId(3)]
        object Window { get; }
        [DispId(4)]
        bool Visible { get; set; }
        [DispId(5)]
        object ContentControl { get; }
        [DispId(6)]
        int Height { get; set; }
        [DispId(7)]
        int Width { get; set; }
        [DispId(8)]
        int DockPosition { get; set; }
        [DispId(9)]
        int DockPositionRestrict { get; set; }
        [DispId(10)]
        void Delete();
    }

    /// <summary>
    /// IObjectSafety — Office hosts query this interface on ActiveX controls
    /// before allowing them to be hosted in a CTP.  We declare both
    /// INTERFACESAFE_FOR_UNTRUSTED_CALLER and _DATA to satisfy the host.
    /// </summary>
    [ComImport]
    [Guid("CB5BDC81-93C1-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        void GetInterfaceSafetyOptions(
            ref Guid riid,
            out int pdwSupportedOptions,
            out int pdwEnabledOptions);

        void SetInterfaceSafetyOptions(
            ref Guid riid,
            int dwOptionSetMask,
            int dwEnabledOptions);
    }
}
