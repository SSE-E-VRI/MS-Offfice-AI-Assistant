using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.Addin
{
    #region COM Interfaces for IDTExtensibility2 and IRibbonExtensibility

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
    /// Microsoft.VisualStudio.CommandBars / Extensibility IDTExtensibility2
    /// Canonical COM Definition: Dual interface with DispIds 1..5
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            ext_DisconnectMode RemoveMode,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>
    /// Microsoft.Office.Core.IRibbonExtensibility
    /// Canonical COM Definition: Dual interface with DispId 1
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    #endregion

    /// <summary>
    /// The primary COM Add-in entry point for Microsoft Office (Word, Excel, PowerPoint, Outlook).
    /// AutoDual enables Office to discover ribbon callback methods (OnToggleSidebar, etc.) via IDispatch.
    /// </summary>
    [Guid("2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78")]
    [ProgId("MistralAI.Addin")]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private object _appObj;
        private string _hostType = "Office";
        private CustomTaskPaneManager _taskPaneManager;
        private RibbonCallback _ribbonCallback;

        static Connect()
        {
            try
            {
                // Ensure TLS 1.2 is active across the Office host process for modern HTTPS APIs
                System.Net.ServicePointManager.SecurityProtocol |= (System.Net.SecurityProtocolType)3072 | (System.Net.SecurityProtocolType)12288 | System.Net.SecurityProtocolType.Tls12;
            }
            catch { }

            try
            {
                AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
            }
            catch { }
        }

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                string location = typeof(Connect).Assembly.Location;
                if (string.IsNullOrEmpty(location)) return null;

                string folder = Path.GetDirectoryName(location);
                string simpleName = new AssemblyName(args.Name).Name;
                string candidate = Path.Combine(folder, simpleName + ".dll");
                if (File.Exists(candidate))
                    return Assembly.LoadFrom(candidate);
            }
            catch { }
            return null;
        }

        public Connect()
        {
            try
            {
                _taskPaneManager = new CustomTaskPaneManager();
                _ribbonCallback = new RibbonCallback(_taskPaneManager);
                Logger.Info("Connect constructor completed.");
            }
            catch (Exception ex)
            {
                Logger.Error("Connect constructor error", ex);
            }
        }

        #region IDTExtensibility2 Implementation

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Logger.Info(string.Format("OnConnection ENTERED (Mode: {0})", ConnectMode));
            try
            {
                _appObj = Application;

                if (Application != null)
                {
                    try
                    {
                        Type t = Application.GetType();
                        string name = t.InvokeMember("Name",
                            System.Reflection.BindingFlags.GetProperty, null, Application, null) as string;

                        Logger.Info(string.Format("Host application: {0}", name ?? "(unknown)"));

                        if (name != null)
                        {
                            if (name.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0)
                                _hostType = "Word";
                            else if (name.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0)
                                _hostType = "Excel";
                            else if (name.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
                                _hostType = "PowerPoint";
                            else if (name.IndexOf("Outlook", StringComparison.OrdinalIgnoreCase) >= 0)
                                _hostType = "Outlook";
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("Could not read Application.Name: {0}", ex.Message));
                    }
                }

                if (_taskPaneManager == null)
                {
                    _taskPaneManager = new CustomTaskPaneManager();
                    _ribbonCallback = new RibbonCallback(_taskPaneManager);
                }

                _taskPaneManager.SetHost(_appObj, _hostType);
                Logger.Info(string.Format("OnConnection completed OK. Host={0}", _hostType));
            }
            catch (Exception ex)
            {
                Logger.Error("OnConnection error (swallowed)", ex);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Logger.Info(string.Format("OnDisconnection (Mode: {0})", RemoveMode));
            try
            {
                if (_taskPaneManager != null)
                    _taskPaneManager.Cleanup();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OnDisconnection error: {0}", ex.Message));
            }
            finally
            {
                _appObj = null;
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            Logger.Info("OnStartupComplete: Office fully loaded.");
            try
            {
                if (_appObj != null)
                {
                    var ver = VersionDetector.DetectVersion(_appObj);
                    Logger.Info(string.Format("Detected Office version: {0}", ver));
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("OnStartupComplete error: {0}", ex.Message));
            }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Logger.Info("OnBeginShutdown.");
            try
            {
                if (_taskPaneManager != null)
                    _taskPaneManager.Cleanup();
            }
            catch { }
        }

        #endregion

        #region IRibbonExtensibility Implementation

        public string GetCustomUI(string RibbonID)
        {
            Logger.Info(string.Format("GetCustomUI: {0}", RibbonID));
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("MistralOfficeAddin.Addin.Ribbon.xml"))
                {
                    if (stream != null)
                    {
                        using (var reader = new StreamReader(stream))
                            return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("GetCustomUI failed to load Ribbon.xml", ex);
            }

            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""tabMistralAI"" label=""AI Assistant"">
        <group id=""grpChat"" label=""AI Chat"">
          <button id=""btnToggleSidebar"" label=""Open Chat"" imageMso=""HappyFace"" size=""large"" onAction=""OnToggleSidebar""/>
          <button id=""btnSettings"" label=""Configure"" imageMso=""AdpDiagramTableRelationships"" size=""large"" onAction=""OnOpenSettings""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        #endregion

        #region ICustomTaskPaneConsumer Implementation

        public void CTPFactoryAvailable(object CTPFactoryInst)
        {
            Logger.Info(string.Format("CTPFactoryAvailable: {0}",
                CTPFactoryInst != null ? CTPFactoryInst.GetType().FullName : "null"));
            try
            {
                if (_taskPaneManager == null)
                {
                    _taskPaneManager = new CustomTaskPaneManager();
                    _ribbonCallback = new RibbonCallback(_taskPaneManager);
                    _taskPaneManager.SetHost(_appObj, _hostType);
                }
                _taskPaneManager.SetFactory(CTPFactoryInst);
            }
            catch (Exception ex)
            {
                Logger.Error("CTPFactoryAvailable error", ex);
            }
        }

        #endregion

        #region Ribbon Callback Dispatchers (called by Office via IDispatch)

        public void OnToggleSidebar(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnToggleSidebar(control); }
            catch (Exception ex) { Logger.Error("OnToggleSidebar error", ex); }
        }

        public void OnNewChat(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnNewChat(control); }
            catch (Exception ex) { Logger.Error("OnNewChat error", ex); }
        }

        public void OnGenerate(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnGenerate(control); }
            catch (Exception ex) { Logger.Error("OnGenerate error", ex); }
        }

        public void OnContinueWriting(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnContinueWriting(control); }
            catch (Exception ex) { Logger.Error("OnContinueWriting error", ex); }
        }

        public void OnSummarize(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnSummarize(control); }
            catch (Exception ex) { Logger.Error("OnSummarize error", ex); }
        }

        public void OnRewrite(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnRewrite(control); }
            catch (Exception ex) { Logger.Error("OnRewrite error", ex); }
        }

        public void OnExpand(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnExpand(control); }
            catch (Exception ex) { Logger.Error("OnExpand error", ex); }
        }

        public void OnShorten(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnShorten(control); }
            catch (Exception ex) { Logger.Error("OnShorten error", ex); }
        }

        public void OnTranslate(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnTranslate(control); }
            catch (Exception ex) { Logger.Error("OnTranslate error", ex); }
        }

        public void OnOpenSettings(object control)
        {
            try { if (_ribbonCallback != null) _ribbonCallback.OnOpenSettings(control); }
            catch (Exception ex) { Logger.Error("OnOpenSettings error", ex); }
        }

        #endregion
    }
}
