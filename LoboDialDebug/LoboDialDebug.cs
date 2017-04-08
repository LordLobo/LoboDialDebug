using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Windows.UI.Input;
using Microsoft.VisualStudio.Shell.Interop;
using Windows.UI.Notifications;
using System.Linq;

namespace LoboDialDebug
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(LoboDialDebug.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class LoboDialDebug : Package
    {
        /// <summary>
        /// LoboDialDebug GUID string. DO NOT TOUCH
        /// </summary>
        public const string PackageGuidString = "ce965617-426e-4b6b-8635-af4e416e61b0";

        private DTE _dte;
        private RadialController _radialController;
        private List<RadialControllerMenuItem> _menuItems;

        /// <summary>
        /// Initializes a new instance of the <see cref="LoboDialDebug"/> class.
        /// </summary>
        public LoboDialDebug()
        {
            Log.dl("LoblDialDeubg INIT");
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            _dte = GetService(typeof(DTE)) as DTE;

            if (_dte == null)
            {
                throw new NullReferenceException("DTE is null");
            }

            CreateController();
            CreateMenuItem();
            HookUpEvents();
        }

        private void CreateController()
        {
            IRadialControllerInterop interop = (IRadialControllerInterop)System.Runtime.InteropServices.WindowsRuntime.WindowsRuntimeMarshal.GetActivationFactory(typeof(RadialController));
            Guid guid = typeof(RadialController).GetInterface("IRadialController").GUID;

            _radialController = interop.CreateForWindow(new IntPtr(_dte.ActiveWindow.HWnd), ref guid);
        }

        private void CreateMenuItem()
        {
            _menuItems = new List<RadialControllerMenuItem>
            {
                RadialControllerMenuItem.CreateFromKnownIcon("LoboDebug", RadialControllerMenuKnownIcon.InkColor)
            };

            foreach (var item in _menuItems)
            {
                _radialController.Menu.Items.Add(item);
            }
        }

        private void HookUpEvents()
        {
            _radialController.RotationResolutionInDegrees = 25;

            _radialController.RotationChanged += OnRotationChanged;
            _radialController.ButtonClicked += OnButtonClicked;

            _radialController.ScreenContactStarted += ScreenContactStarted;

            // probably don't want this
            //_radialController.ControlLost += _radialController_ControlLost;

            _dte.Events.SolutionEvents.AfterClosing += () =>
            {
                Log.dl("AfterClosing time!");
                _radialController.Menu.Items.Clear();
            };
        }

        private void _radialController_ControlLost(RadialController sender, object args)
        {
            Log.dl("Control Lost");
            _radialController.Menu.Items.Clear();
        }

        private void ScreenContactStarted(RadialController sender, RadialControllerScreenContactStartedEventArgs args)
        {
            Log.dl("************* SCREEN CONTACT!");
        }

        private void OnButtonClicked(RadialController sender, RadialControllerButtonClickedEventArgs args)
        {
            if (_dte.Application.Debugger.CurrentMode == dbgDebugMode.dbgRunMode)
            {
                _dte.Application.Debugger.Stop();

                ShowToast("LoboDialDebug", "STOP OH GODS THE PAIN");
            }
            else
            {
                _dte.Application.Debugger.Go();

                ShowToast("LoboDialDebug", "DEBUG GO");
            }
        }

        private void OnRotationChanged(RadialController sender, RadialControllerRotationChangedEventArgs args)
        {
            if (args.RotationDeltaInDegrees > 0)
            {
                Log.dl("Clockwise");

                _dte.Application.Debugger.StepInto();

                //ShowToast("LoboDialDebug", "Step Into");
            }
            else
            {
                Log.dl("CounterClockwise");

                _dte.Application.Debugger.StepOver();

                //ShowToast("LoboDialDebug", "Step Over");
            }
        }

        void ShowToast(string title, string message)
        {
            var toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastText04);
            
            toastXml.GetElementsByTagName("text").First().AppendChild(toastXml.CreateTextNode(title));
            toastXml.GetElementsByTagName("text").Last().AppendChild(toastXml.CreateTextNode(message));
            
            var dte = GetGlobalService(typeof(DTE)) as DTE;
            var notifier = ToastNotificationManager.CreateToastNotifier(EditionToAppUserModelId(dte.Edition, dte.Version));
            notifier.Show(new ToastNotification(toastXml));
        }

        string EditionToAppUserModelId(string edition, string version)
        {
            switch (edition)
            {
                case "WD Express":
                    return "VWDExpress." + version;
                case "Desktop Express":
                    return "WDExpress." + version;
                case "VSWin Express":
                    return "VSWinExpress." + version;
                case "PD Express":
                    return "VPDExpress." + version;
            }
            // detect AppUserModelId
            var s = Shell32.GetCurrentProcessExplicitAppUserModelID();
            return !string.IsNullOrEmpty(s) ? s : "VisualStudio." + version;
        }
    }

    static class Shell32
    {
        [DllImport("Shell32.dll")]
        public static extern IntPtr GetCurrentProcessExplicitAppUserModelID(out IntPtr AppID);

        public static string GetCurrentProcessExplicitAppUserModelID()
        {
            IntPtr pv;
            GetCurrentProcessExplicitAppUserModelID(out pv);
            if (pv == null) return null;
            var s = Marshal.PtrToStringAuto(pv);
            Ole32.CoTaskMemFree(pv);
            return s;
        }
    }

    static class User32
    {
        /// <summary>The GetForegroundWindow function returns a handle to the foreground window.</summary>
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }

    static class Ole32
    {
        [DllImport("ole32.dll")]
        public static extern void CoTaskMemFree(IntPtr pv);
    }
}
