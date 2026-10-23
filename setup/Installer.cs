using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        string installDir = DefaultInstallDir();
        bool desktopShortcut = true;
        bool silent = false;

        foreach (string arg in args)
        {
            if (arg.StartsWith("/DIR=", StringComparison.OrdinalIgnoreCase))
            {
                installDir = arg.Substring(5).Trim().Trim('"');
            }
            else if (string.Equals(arg, "/SILENT", StringComparison.OrdinalIgnoreCase))
            {
                silent = true;
            }
            else if (string.Equals(arg, "/NODESKTOP", StringComparison.OrdinalIgnoreCase))
            {
                desktopShortcut = false;
            }
        }

        if (silent)
        {
            try
            {
                Install(installDir, desktopShortcut);
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }

        Application.Run(new InstallForm(installDir, desktopShortcut));
        return 0;
    }

    internal static string DefaultInstallDir()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "批量重命名工具");
    }

    internal static void Install(string installDir, bool desktopShortcut)
    {
        if (string.IsNullOrWhiteSpace(installDir))
        {
            throw new InvalidOperationException("请选择安装目录。");
        }

        installDir = Path.GetFullPath(installDir.Trim());
        Directory.CreateDirectory(installDir);
        ExtractPayload(installDir);

        string vbs = Path.Combine(installDir, "启动.vbs");
        string core = Path.Combine(installDir, "BatchRenamer.Core.ps1");
        string main = Path.Combine(installDir, "BatchRenamer.ps1");
        if (!File.Exists(vbs) || !File.Exists(core) || !File.Exists(main))
        {
            throw new InvalidOperationException("安装包不完整，缺少启动文件。");
        }

        WriteUninstall(installDir, desktopShortcut);
        if (desktopShortcut)
        {
            CreateDesktopShortcut(vbs, installDir);
        }
    }

    private static void ExtractPayload(string installDir)
    {
        Assembly asm = Assembly.GetExecutingAssembly();
        using (Stream src = asm.GetManifestResourceStream("Payload.zip"))
        {
            if (src == null)
            {
                throw new InvalidOperationException("安装包缺少程序文件。");
            }

            string stage = Path.Combine(Path.GetTempPath(), "br_setup_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stage);
            string zipPath = Path.Combine(stage, "payload.zip");
            string unpack = Path.Combine(stage, "unpack");
            try
            {
                using (FileStream dst = File.Create(zipPath))
                {
                    src.CopyTo(dst);
                }
                ZipFile.ExtractToDirectory(zipPath, unpack);
                foreach (string file in Directory.GetFiles(unpack, "*", SearchOption.TopDirectoryOnly))
                {
                    string name = Path.GetFileName(file);
                    File.Copy(file, Path.Combine(installDir, name), true);
                }
            }
            finally
            {
                try { Directory.Delete(stage, true); } catch { }
            }
        }
    }

    private static void CreateDesktopShortcut(string vbsPath, string workingDirectory)
    {
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string lnk = Path.Combine(desktop, "批量重命名工具.lnk");
        Type shellType = Type.GetTypeFromProgID("WScript.Shell");
        object shell = Activator.CreateInstance(shellType);
        object shortcut = shellType.InvokeMember(
            "CreateShortcut",
            BindingFlags.InvokeMethod,
            null,
            shell,
            new object[] { lnk });
        Type shortcutType = shortcut.GetType();
        shortcutType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, new object[] { vbsPath });
        shortcutType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortcut, new object[] { workingDirectory });
        shortcutType.InvokeMember("WindowStyle", BindingFlags.SetProperty, null, shortcut, new object[] { 1 });
        shortcutType.InvokeMember("Description", BindingFlags.SetProperty, null, shortcut, new object[] { "批量重命名工具" });
        shortcutType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
    }

    private static void WriteUninstall(string installDir, bool desktopShortcut)
    {
        string path = Path.Combine(installDir, "卸载.vbs");
        string lines =
            "Set fso = CreateObject(\"Scripting.FileSystemObject\")" + "\r\n" +
            "Set sh = CreateObject(\"WScript.Shell\")" + "\r\n" +
            "installDir = fso.GetParentFolderName(WScript.ScriptFullName)" + "\r\n" +
            "answer = MsgBox(\"确定卸载批量重命名工具吗？\", 4 + 32, \"卸载\")" + "\r\n" +
            "If answer <> 6 Then WScript.Quit" + "\r\n";
        if (desktopShortcut)
        {
            lines +=
                "desk = sh.SpecialFolders(\"Desktop\")" + "\r\n" +
                "lnk = desk & \"\\批量重命名工具.lnk\"" + "\r\n" +
                "If fso.FileExists(lnk) Then fso.DeleteFile lnk, True" + "\r\n";
        }
        lines +=
            "cmd = \"cmd.exe /c ping 127.0.0.1 -n 2 >nul & rd /s /q \"\"\" & installDir & \"\"\"\"" + "\r\n" +
            "sh.Run cmd, 0, False" + "\r\n";
        File.WriteAllText(path, lines, System.Text.Encoding.Default);
    }
}

internal sealed class InstallForm : Form
{
    private readonly TextBox txtDir;
    private readonly CheckBox chkDesktop;
    private readonly Button btnInstall;

    public InstallForm(string initialDir, bool desktopShortcut)
    {
        Text = "安装 批量重命名工具";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(520, 188);
        Font = new Font("Microsoft YaHei UI", 9F);

        Label title = new Label();
        title.Text = "选择安装目录，完成后可从桌面快捷方式打开。";
        title.Location = new Point(20, 16);
        title.AutoSize = true;
        Controls.Add(title);

        Label lblDir = new Label();
        lblDir.Text = "安装到";
        lblDir.Location = new Point(20, 52);
        lblDir.AutoSize = true;
        Controls.Add(lblDir);

        txtDir = new TextBox();
        txtDir.Location = new Point(20, 76);
        txtDir.Size = new Size(380, 26);
        txtDir.Text = initialDir;
        Controls.Add(txtDir);

        Button btnBrowse = new Button();
        btnBrowse.Text = "浏览...";
        btnBrowse.Location = new Point(410, 74);
        btnBrowse.Size = new Size(90, 28);
        btnBrowse.Click += delegate
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择安装目录";
                dialog.SelectedPath = txtDir.Text;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    txtDir.Text = dialog.SelectedPath;
                }
            }
        };
        Controls.Add(btnBrowse);

        chkDesktop = new CheckBox();
        chkDesktop.Text = "在桌面创建快捷方式";
        chkDesktop.Location = new Point(20, 114);
        chkDesktop.AutoSize = true;
        chkDesktop.Checked = desktopShortcut;
        Controls.Add(chkDesktop);

        btnInstall = new Button();
        btnInstall.Text = "安装";
        btnInstall.Location = new Point(314, 146);
        btnInstall.Size = new Size(90, 28);
        btnInstall.Click += OnInstall;
        AcceptButton = btnInstall;
        Controls.Add(btnInstall);

        Button btnCancel = new Button();
        btnCancel.Text = "取消";
        btnCancel.Location = new Point(410, 146);
        btnCancel.Size = new Size(90, 28);
        btnCancel.Click += delegate { Close(); };
        CancelButton = btnCancel;
        Controls.Add(btnCancel);
    }

    private void OnInstall(object sender, EventArgs e)
    {
        btnInstall.Enabled = false;
        try
        {
            Program.Install(txtDir.Text, chkDesktop.Checked);
            MessageBox.Show(this, "安装完成。", "批量重命名工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            btnInstall.Enabled = true;
        }
    }
}
