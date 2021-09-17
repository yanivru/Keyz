using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;

namespace Keyz
{
    class StartupProjectProvider
    {
        private DTE _dte;

        public StartupProjectProvider(DTE dte)
        {
            _dte = dte;
        }

        public Project GetStartupProject()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var startupProjects = (Array)_dte.Solution.SolutionBuild.StartupProjects;

            if (startupProjects.Length > 0)
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
