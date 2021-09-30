using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Keyz
{
    internal class TfsService
    {
        private DTE _dte;
        private readonly OutputWindowLogger _outputLogger;

        public TfsService(DTE dte, OutputWindowLogger outputLogger)
        {
            _dte = dte;
            _outputLogger = outputLogger;
        }

        internal object GetLatestForSoltution()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Invoking Get method on SolutionWorkspace since it's hard to refernce the correct
            // Microsoft.TeamFoundation.VersionControl dll for each visual studio version.
            object solutionWorkspace = GetSolutionWorkspace();
            if (solutionWorkspace == null)
            {
                _outputLogger.Write("No solution workspace");
                return null;
            }

            var solutionWorkspaceType = solutionWorkspace.GetType();

            object latestVersion = GetVersionSpecLatest(solutionWorkspace, solutionWorkspaceType);

            object recursionFull = GetFullRecursionEnumValue(solutionWorkspaceType);

            object getOptionsNone = GetGetOptionsNone(solutionWorkspaceType);

            var solutionDirectory = Path.GetDirectoryName(_dte.Solution.FullName);

            return InvokeMethodDynamically(solutionWorkspace, "Get", new string[]{ solutionDirectory}, latestVersion, recursionFull, getOptionsNone);
        }

        private object InvokeMethodDynamically(object targetObject, string methodName, params object[] parameters)
        {
            var parametersTypes = parameters.Select(x => x.GetType()).ToArray();

            var methodInfo = targetObject.GetType().GetMethod(
                methodName,
                BindingFlags.Public | BindingFlags.Instance,
                null,
                parametersTypes,
                null);

            return methodInfo.Invoke(
                targetObject,
                BindingFlags.Public | BindingFlags.Instance,
                null,
                parameters,
                CultureInfo.InvariantCulture);
        }

        private static object GetGetOptionsNone(Type solutionWorkspaceType)
        {
            var getOptionsType = solutionWorkspaceType.Assembly.GetType("Microsoft.TeamFoundation.VersionControl.Client.GetOptions");
            var getOptionsEnumValues = getOptionsType.GetEnumValues();
            object getOptionsNone = getOptionsEnumValues.GetValue(0);
            return getOptionsNone;
        }

        private static object GetFullRecursionEnumValue(Type solutionWorkspaceType)
        {
            var recursionType = solutionWorkspaceType.Assembly.GetType("Microsoft.TeamFoundation.VersionControl.Client.RecursionType");
            var recursionEnumValues = recursionType.GetEnumValues();
            object recursionFull = recursionEnumValues.GetValue(2);
            return recursionFull;
        }

        private static object GetVersionSpecLatest(object solutionWorkspace, Type solutionWorkspaceType)
        {
            var versionSpecType = solutionWorkspaceType.Assembly.GetType("Microsoft.TeamFoundation.VersionControl.Client.VersionSpec");
            var latestProperty = versionSpecType.GetProperty("Latest");
            var latestVersion = latestProperty.GetValue(solutionWorkspace);
            return latestVersion;
        }

        private object GetSolutionWorkspace()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            object versionControlExt = _dte.GetObject("Microsoft.VisualStudio.TeamFoundation.VersionControl.VersionControlExt");
            object solutionWorkspace = versionControlExt.GetType().GetProperty("SolutionWorkspace").GetValue(versionControlExt);
            return solutionWorkspace;
        }
    }
}