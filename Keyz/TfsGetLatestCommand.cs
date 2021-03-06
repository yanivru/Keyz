using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace Keyz
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TfsGetLatestCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4433;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("2b04abce-197a-4ba3-9f60-68e4f526a5e1");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TfsGetLatestCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private TfsGetLatestCommand(AsyncPackage package, OleMenuCommandService commandService, TfsService tfsService, OutputWindowLogger outputLogger)
        {
            _tfsService = tfsService;
            _outputLogger = outputLogger;

            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TfsGetLatestCommand Instance
        {
            get;
            private set;
        }

        private readonly TfsService _tfsService;
        private readonly OutputWindowLogger _outputLogger;

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package, TfsService tfsService, OutputWindowLogger outputLogger)
        {
            // Switch to the main thread - the call to AddCommand in TfsGetLatestCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new TfsGetLatestCommand(package, commandService, tfsService, outputLogger);
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
            dynamic status = _tfsService.GetLatestForSoltution();
            if (status == null)
            {
                _outputLogger.Write("No status returned", true);
            }
            else
            {
                _outputLogger.Write("Status: ");
                _outputLogger.Write($"   No action needed: {status.NoActionNeeded}");
                _outputLogger.Write($"   #Bytes: {status.NumBytes}");
                _outputLogger.Write($"   #Conflicts: {status.NumConflicts}");
                _outputLogger.Write($"   #Failures: {status.NumFailures}");
                _outputLogger.Write($"   #Files: {status.NumFiles}");
                _outputLogger.Write($"   #Operations: {status.NumOperations}");
                _outputLogger.Write($"   #ResolvedConflicts: {status.NumResolvedConflicts}");
                _outputLogger.Write($"   #Updated: {status.NumUpdated}");
                _outputLogger.Write($"   #Warnings: {status.NumWarnings}");
            }
        }
    }
}
