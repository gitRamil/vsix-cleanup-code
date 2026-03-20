using Microsoft.VisualStudio.Shell;
using System.ComponentModel;

namespace JBCleanupExtension
{
    public class CleanupOptions : DialogPage
    {
        [Category("JetBrains Cleanup")]
        [DisplayName("Profile Name")]
        [Description("Cleanup profile name (e.g.'Built-in: Full Cleanup')")]
        public string ProfileName { get; set; } = "";

        [Category("JetBrains Cleanup")]
        [DisplayName("DotSettings File Name")]
        [Description("Settings file name relative to the solution folder. Leave empty to auto-detect based on the .sln file name (e.g. MyProject.sln → MyProject.sln.DotSettings)")]
        public string DotSettingsFileName { get; set; } = "";
    }
}