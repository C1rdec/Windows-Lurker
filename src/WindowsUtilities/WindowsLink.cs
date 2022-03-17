using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using static WindowsUtilities.Native;

namespace WindowsUtilities
{
    public class WindowsLink
    {
        #region Fields

        private string _linkName;
        private string _description;

        #endregion

        #region Constructors

        public WindowsLink(string linkName, string description)
        {
            _linkName = linkName;
            _description = description;
        }

        #endregion

        #region Properties

        private string ApplicationDataFolderPath => Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

        private string ProgramsFolderPath => Path.Combine(ApplicationDataFolderPath, @"Microsoft\Windows\Start Menu\Programs");

        private string StartupFolderPath => Path.Combine(ProgramsFolderPath, "Startup");

        private string StartupFilePath => Path.Combine(StartupFolderPath, _linkName);

        private string MenuFilePath => Path.Combine(ProgramsFolderPath, _linkName);

        #endregion

        #region Methods

        public void AddStartWithWindows() => CreateLink(StartupFilePath);

        public void RemoveStartWithWindows() => RemoveLink(StartupFilePath);

        public void AddStartMenu() => CreateLink(MenuFilePath);

        public void RemoveStartMenu() => RemoveLink(MenuFilePath);

        private void RemoveLink(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private void CreateLink(string path)
        {
            RemoveLink(path);

            var link = (IShellLink)new ShellLink();
            link.SetDescription(_description);
            var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            link.SetPath(exePath);
            var file = (IPersistFile)link;
            file.Save(path, false);
        }

        #endregion
    }
}