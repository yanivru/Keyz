using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace Keyz
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
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(KeyzPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class KeyzPackage : AsyncPackage
    {
        /// <summary>
        /// KeyzPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "574fe2f0-46e3-4f55-aaaf-06ea0262e775";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            OleMenuCommandService commandService = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            DTE dte = (DTE)await GetServiceAsync(typeof(DTE));
            //object b = dte.GetObject("Microsoft.VisualStudio.TeamFoundation.VersionControl.VersionControlExt");
            //object sw = b.GetType().GetProperty("SolutionWorkspace").GetValue(b);
            //var w = (Type)sw.GetType();
            //var versionSpec = w.Assembly.GetType("Microsoft.TeamFoundation.VersionControl.Client.VersionSpec");
            //var latest = versionSpec.GetProperty("Latest");
            //var recursion = w.Assembly.GetType("Microsoft.TeamFoundation.VersionControl.Client.RecursionType");
            //var enums = (Array)recursion.GetEnumValues();
            //var getOptions = w.Assembly.GetType("Microsoft.TeamFoundation.VersionControl.Client.GetOptions");
            //var enumValues = (Array)getOptions.GetEnumValues();

            //var dir = Path.GetDirectoryName(dte.Solution.FullName);

            //string[] dirs = { dir };

            //var get = w.GetMethod("Get", BindingFlags.Public | BindingFlags.Instance, null, new Type[] {typeof(string[]), versionSpec, enums.GetValue(2).GetType(), enumValues.GetValue(0).GetType() }, null);

            //var l = latest.GetValue(sw);
            //var o = get.Invoke(sw, BindingFlags.Public | BindingFlags.Instance, null, new[] { dirs, l, enums.GetValue(2), enumValues.GetValue(0) }, CultureInfo.InvariantCulture);


            //var c = b.SolutionWorkspace.Get(new string[]{ dte.Solution.FullName }, VersionSpec.Latest, RecursionType.Full, GetOptions.None);
            IVsActivityLog activityLog = await GetServiceAsync(typeof(SVsActivityLog)) as IVsActivityLog;

            var shell = new Shell();

            await OpenOutputCommand.InitializeAsync(this, shell, new StartupProjectProvider(dte, activityLog));
            await OpenSolutionFolderCommand.InitializeAsync(this, dte, shell);
        }

        #endregion
    }
}
