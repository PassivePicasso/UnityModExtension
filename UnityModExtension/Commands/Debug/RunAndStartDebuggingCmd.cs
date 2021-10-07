using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace UnityModExtension.Commands.Debug
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class RunAndStartDebuggingCmd
    {
        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static RunAndStartDebuggingCmd Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("2bfd9814-8bf4-4d5a-828e-1976305175b3");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public string TargetPath
        {
            get
            {
                GeneralOptionPage page = (GeneralOptionPage)package.GetDialogPage(typeof(GeneralOptionPage));
                return page.TargetPath;
            }
        }

        public string TargetArguments
        {
            get
            {
                GeneralOptionPage page = (GeneralOptionPage)package.GetDialogPage(typeof(GeneralOptionPage));
                return page.TargetArguments;
            }
        }
        public string WorkingDirectory
        {
            get
            {
                GeneralOptionPage page = (GeneralOptionPage)package.GetDialogPage(typeof(GeneralOptionPage));
                return page.WorkingDirectory;
            }
        }

        public int TargetPort
        {
            get
            {
                GeneralOptionPage page = (GeneralOptionPage)package.GetDialogPage(typeof(GeneralOptionPage));
                return page.TargetPort;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RunAndStartDebuggingCmd"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private RunAndStartDebuggingCmd(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);

            menuItem.ParametersDescription = "$";
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in RunAndStartDebuggingCmd's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new RunAndStartDebuggingCmd(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var eventArgs = e as OleMenuCmdEventArgs;
            var args = string.Empty;
            if (eventArgs.InValue != null)
            {
                args = eventArgs.InValue.ToString();
            }

            ToolsForUnity.Init();
            if (ToolsForUnity.Loaded)
            {
                try
                {
                    if (string.IsNullOrEmpty(args))
                    {
                        RunTarget(TargetPath, WorkingDirectory, TargetArguments);
                        LaunchDebugger(TargetPort);
                    }
                    else
                    {
                        int port = TargetPort;
                        string arguments = TargetArguments;
                        string path = TargetPath;
                        string workingDirectory = WorkingDirectory;
                        foreach (var arg in args.SplitCommandLine())
                            switch (arg)
                            {
                                case var argument when arg.StartsWith($"-{nameof(TargetPath)}=", StringComparison.OrdinalIgnoreCase):
                                    path = argument.TrimPrefix($"-{nameof(TargetPath)}=", StringComparison.OrdinalIgnoreCase).TrimMatchingQuotes('"');
                                    break;
                                case var argument when arg.StartsWith($"-{nameof(TargetArguments)}=", StringComparison.OrdinalIgnoreCase):
                                    arguments = argument.TrimPrefix($"-{nameof(TargetArguments)}=", StringComparison.OrdinalIgnoreCase).TrimMatchingQuotes('"');
                                    arguments = Regex.Unescape(arguments);
                                    break;
                                case var argument when arg.StartsWith($"-{nameof(WorkingDirectory)}=", StringComparison.OrdinalIgnoreCase):
                                    workingDirectory = argument.TrimPrefix($"-{nameof(WorkingDirectory)}=", StringComparison.OrdinalIgnoreCase).TrimMatchingQuotes('"');
                                    break;
                                case var argument when arg.StartsWith($"-{nameof(TargetPort)}=", StringComparison.OrdinalIgnoreCase):
                                    port = int.Parse(argument.TrimPrefix($"-{nameof(TargetPort)}=", StringComparison.OrdinalIgnoreCase));
                                    break;
                            }

                        RunTarget(path, workingDirectory, arguments);
                        LaunchDebugger(port);
                    }
                }
                catch (Exception ex)
                {
                    UnityModExtension.Debug.Print(ex.ToString());
                }
            }
            else
            {

            }
        }

        private void RunTarget(string path, string workingDirectory, string arguments)
        {
            if (string.IsNullOrEmpty(workingDirectory))
                workingDirectory = Path.GetDirectoryName(path);

            var processStartInfo = new ProcessStartInfo(path)
            {
                WorkingDirectory = workingDirectory,
                Arguments = arguments,
                UseShellExecute = true
            };

            Process.Start(processStartInfo);
        }

        private void LaunchDebugger(int port)
        {
            var unityProcess = ToolsForUnity.CreateUnityProcess(port);
            ToolsForUnity.LaunchDebugger(unityProcess);
        }

    }
}
