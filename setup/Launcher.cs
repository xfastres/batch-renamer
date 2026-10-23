using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class Launcher
{
    [STAThread]
    private static int Main()
    {
        string folder = AppDomain.CurrentDomain.BaseDirectory;
        string script = Path.Combine(folder, "BatchRenamer.ps1");
        if (!File.Exists(script))
        {
            MessageBox.Show("找不到 BatchRenamer.ps1，请与启动程序放在同一文件夹。", "批量重命名工具",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return 1;
        }

        ProcessStartInfo info = new ProcessStartInfo();
        info.FileName = Path.Combine(Environment.SystemDirectory, "WindowsPowerShell\\v1.0\\powershell.exe");
        if (!File.Exists(info.FileName))
        {
            info.FileName = "powershell.exe";
        }
        info.Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" + script + "\"";
        info.UseShellExecute = false;
        info.CreateNoWindow = true;
        info.WorkingDirectory = folder;
        Process.Start(info);
        return 0;
    }
}
