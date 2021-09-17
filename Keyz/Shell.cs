namespace Keyz
{
    class Shell
    {
        public void OpenFolder(string folderName)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = folderName,
                UseShellExecute = true,
                Verb = "open"
            });
        }
    }
}
