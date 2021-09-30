using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace Keyz
{
    class StartupProjectProvider
    {
        private readonly DTE _dte;
        private readonly IVsActivityLog _log;

        public StartupProjectProvider(DTE dte, Microsoft.VisualStudio.Shell.Interop.IVsActivityLog activityLog)
        {
            _dte = dte;
            _log = activityLog;
        }

        public Project GetStartupProject()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var startupProjects = (Array)_dte.Solution?.SolutionBuild?.StartupProjects;

            _log.LogEntry((uint)__ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION, "KeyZ", $"# of startup projects {startupProjects?.Length}");
            if (startupProjects?.Length > 0)
            {
                foreach (Project project in _dte.Solution.Projects)
                {
                    if (project.UniqueName.Equals(startupProjects.GetValue(0)))
                    {
                        return project;
                    }
                }
            }

            return null;
        }
    }
}
