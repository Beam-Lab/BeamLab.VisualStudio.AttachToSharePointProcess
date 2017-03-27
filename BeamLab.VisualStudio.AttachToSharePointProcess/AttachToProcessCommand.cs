//------------------------------------------------------------------------------
// <copyright file="AttachToProcessCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System.Runtime.InteropServices;

namespace BeamLab.VisualStudio.AttachToSharePointProcess
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AttachToProcessCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("1649511e-f9ff-49e0-9887-daaf7ee9c712");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AttachToProcessCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private AttachToProcessCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AttachToProcessCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new AttachToProcessCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var _debugger = this.ServiceProvider.GetService(typeof(SVsShellDebugger)) as IVsDebugger;
            var _dte = this.ServiceProvider.GetService(typeof(SDTE)) as EnvDTE80.DTE2;

            Window w = (Window)_dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
            w.Visible = true;
            OutputWindow ow = (OutputWindow)w.Object;
            OutputWindowPane owp = ow.OutputWindowPanes.Add("Attach to SharePoint Processes");
            owp.Activate();

            owp.OutputString("Finding processes..." + Environment.NewLine);

            var processes = _dte.Debugger.LocalProcesses;
            owp.OutputString(string.Format("Total processes: {0}", processes.Count) + Environment.NewLine);
            foreach (Process proc in processes)
            {
                if (proc.Name.IndexOf("Notepad2.exe") != -1)
                {
                    proc.Attach();
                    var p = System.Diagnostics.Process.GetProcessById(proc.ProcessID);

                    owp.OutputString(string.Format("Attach to process: {0} - {1}", p.ProcessName,  proc.ProcessID) + Environment.NewLine);
                }
            }

            //// Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
