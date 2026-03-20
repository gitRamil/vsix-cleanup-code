using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace JBCleanupExtension
{
    internal sealed class CleanupCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("a60f18ee-381d-4e57-8500-33dc3a30708c");

        // Хардкоженные пути — как в твоём PS скрипте
        private const string SolutionPath = @"C:\code\pki-ca\PkiCa.sln";
        private const string RepoRoot = @"C:\code\pki-ca";
        private const string ProfileName = "PkiCaLight";
        private const string DotSettingsFileName = "PkiCa.sln.DotSettings";

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

            var dte = (DTE2)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE));
            if (dte == null) return;

            // Получаем активный файл
            var activeDoc = dte.ActiveDocument;
            if (activeDoc == null)
            {
                ShowMessage("Нет активного файла!");
                return;
            }

            // Проверяем что это .cs файл
            if (!activeDoc.FullName.EndsWith(".cs", StringComparison.OrdinalIgnoreCase))
            {
                ShowMessage("Активный файл не является .cs файлом!");
                return;
            }

            // Сохраняем перед форматированием
            activeDoc.Save();

            RunCleanup(activeDoc.FullName);
        }

        private void RunCleanup(string filePath)
        {
            var settingsPath = System.IO.Path.Combine(RepoRoot, DotSettingsFileName);

            // Собираем аргументы — точно как в PS скрипте:
            // jb cleanupcode <file> --profile=... --settings=... --no-build
            var args = $"cleanupcode \"{filePath}\"" +
                       $" --profile={ProfileName}" +
                       $" --settings=\"{settingsPath}\"" +
                       $" --no-build";

            var psi = new ProcessStartInfo
            {
                FileName = "jb",
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            ShowStatusBar("🔄 JetBrains Cleanup запущен...");

            var process = new System.Diagnostics.Process { StartInfo = psi };

            process.EnableRaisingEvents = true;
            process.Exited += (s, ev) =>
            {
                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                var exitCode = process.ExitCode;

                ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    if (exitCode == 0)
                    {
                        ShowStatusBar("✅ Cleanup завершён успешно!");
                        ReloadActiveDocument();
                    }
                    else
                    {
                        ShowStatusBar("❌ Cleanup завершился с ошибкой");
                        ShowMessage($"Ошибка:\n{error}\n\nВывод:\n{output}");
                    }
                }).FileAndForget("JBCleanupExtension/cleanup");
            };

            process.Start();
        }

        private void ReloadActiveDocument()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = (DTE2)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE));
            dte?.ExecuteCommand("File.ReloadAllFiles");
        }

        private void ShowStatusBar(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var statusBar = (IVsStatusbar)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsStatusbar));
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
    }
}