using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace Keyz
{
    internal class OutputWindowLogger
    {
        private Guid _guid = new Guid("B96F0767-86D9-471D-BA85-146C45B2DB40");
        public void Write(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

            string customTitle = "KeyZ";
            outWindow.CreatePane(ref _guid, customTitle, 1, 1);

            IVsOutputWindowPane customPane;
            outWindow.GetPane(ref _guid, out customPane);

            customPane.OutputString(message + Environment.NewLine);
            customPane.Activate(); // Brings this pane into view
        }
    }
}