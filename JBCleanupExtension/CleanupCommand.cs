using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using Task = System.Threading.Tasks.Task;

namespace JBCleanupExtension
{
    internal sealed class CleanupCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("a60f18ee-381d-4e57-8500-33dc3a30708c");

        private readonly AsyncPackage _package;

        private CleanupCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static CleanupCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync<IMenuCommandService, OleMenuCommandService>();
            Instance = new CleanupCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            if (dte == null) return;

            var solutionPath = dte.Solution?.FullName;
            if (string.IsNullOrEmpty(solutionPath))
            {
                ShowMessage("No solution is open!");
                return;
            }

            var repoRoot = System.IO.Path.GetDirectoryName(solutionPath);

            var activeDoc = dte.ActiveDocument;
            if (activeDoc == null)
            {
                ShowMessage("No active file!");
                return;
            }

            if (!activeDoc.FullName.EndsWith(".cs", StringComparison.OrdinalIgnoreCase))
            {
                ShowMessage("Active file is not a .cs file!");
                return;
            }

            // Get options
            var options = (CleanupOptions)_package.GetDialogPage(typeof(CleanupOptions));
            var profileName = options.ProfileName;

            // If DotSettingsFileName is empty — auto-detect based on .sln file name
            var dotSettingsFileName = string.IsNullOrEmpty(options.DotSettingsFileName)
                ? System.IO.Path.GetFileName(solutionPath) + ".DotSettings"
                : options.DotSettingsFileName;

            activeDoc.Save();
            RunCleanup(repoRoot, activeDoc.FullName, profileName, dotSettingsFileName);
        }

        private void RunCleanup(string repoRoot, string filePath, string profileName, string dotSettingsFileName)
        {
            var jbPath = FindExecutableInPath("jb");
            if (jbPath == null)
            {
                ShowMessage(
                    "JetBrains CLI tool 'jb' was not found in PATH.\n\n" +
                    "Please install it via JetBrains Toolbox or manually:\n" +
                    "https://www.jetbrains.com/help/resharper/ReSharper_Command_Line_Tools.html");
                return;
            }

            var settingsPath = System.IO.Path.Combine(repoRoot, dotSettingsFileName);

            var args = $"cleanupcode \"{filePath}\"" +
                       $" --settings=\"{settingsPath}\"" +
                       $" --no-build";

            if (!string.IsNullOrEmpty(profileName))
            {
                args += $" --profile={profileName}";
            }

            var psi = new ProcessStartInfo
            {
                FileName = "jb",
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            ShowStatusBar("🔄 JetBrains Cleanup running...");

            var process = new System.Diagnostics.Process { StartInfo = psi };

            process.EnableRaisingEvents = true;
            process.Exited += (s, ev) =>
            {
                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                var exitCode = process.ExitCode;

#pragma warning disable VSTHRD110
                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    if (exitCode == 0)
                    {
                        ShowStatusBar("✅ Cleanup completed successfully!");
                        ReloadActiveDocument();
                    }
                    else
                    {
                        ShowStatusBar("❌ Cleanup failed");
                        ShowMessage($"Error:\n{error}\n\nOutput:\n{output}");
                    }
                }).FileAndForget("JBCleanupExtension/cleanup");
#pragma warning restore VSTHRD110
            };

            process.Start();
        }

        private void ReloadActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            dte?.ExecuteCommand("File.ReloadAllFiles");
        }

        private void ShowStatusBar(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var statusBar = (IVsStatusbar)Package.GetGlobalService(typeof(SVsStatusbar));
            statusBar?.SetText(message);
        }

        private void ShowMessage(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            VsShellUtilities.ShowMessageBox(
                _package,
                message,
                "JetBrains Cleanup",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private static string FindExecutableInPath(string executable)
        {
            var pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (string.IsNullOrEmpty(pathEnv)) return null;

            foreach (var dir in pathEnv.Split(System.IO.Path.PathSeparator))
            {
                var fullPath = System.IO.Path.Combine(dir, executable + ".exe");
                if (System.IO.File.Exists(fullPath))
                    return fullPath;
            }

            return null;
        }
    }
}