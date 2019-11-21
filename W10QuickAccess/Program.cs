using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Diagnostics;
using System.IO;

namespace CriarLinkAcessoRapido
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var arguments = "/J" ;
            var quickAccessFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "QuickAccess");

            if (!Directory.Exists(quickAccessFolder))
            {
                Directory.CreateDirectory(quickAccessFolder);
            }

            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;
                dialog.ShowDialog();

                if (IsUnc(dialog.FileName))
                {
                    arguments = "/D";
                }

                Console.WriteLine("Digite o nome do atalho");
                var shortcutName = Console.ReadLine();
                var shortcutPath = Path.Combine(quickAccessFolder, shortcutName);

                Directory.CreateDirectory(shortcutPath);
                Execute(null, "powershell.exe", "-ExecutionPolicy ByPass -File AddQuickAccess.ps1 \"" + shortcutPath + "\"");

                Directory.Delete(shortcutPath);
                Execute(quickAccessFolder, "cmd.exe", "/C mklink " + arguments + " " + "\"" + shortcutName + "\" " + "\"" + dialog.FileName + "\"");
            }

        }

        private static void Execute(string workingDirectory, string fileName, string arguments)
        {
            var startInfo = new ProcessStartInfo();
            startInfo.WorkingDirectory = workingDirectory;
            startInfo.FileName = fileName;
            startInfo.Arguments = arguments;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            var process = new Process();
            process.StartInfo = startInfo;
            process.Start();

            var output = process.StandardOutput.ReadToEnd();
            var errors = process.StandardError.ReadToEnd();

            Console.WriteLine(output);

            if (errors.Length > 0)
            {
                Console.WriteLine("Error: ");
                Console.WriteLine(errors);
                Console.ReadLine();
            }
        }

        private static bool IsUnc(string path)
        {
            var root = Path.GetPathRoot(path);

            // Check if root starts with "\\", clearly an UNC
            if (root.StartsWith(@"\\"))
            {
                return true;
            }

            // Check if the drive is a network drive
            var drive = new DriveInfo(root);
            return drive.DriveType == DriveType.Network;
        }
    }
}
