using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FormatHelperInstaller
{
    class Program
    {
        private static string ProgId = "VSTOFormatHelper.Addin";
        private static string InstallDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordFormatHelper");
        private static string DllName = "VSTOFormatHelper.dll";
        private static string DllPath = Path.Combine(InstallDir, DllName);
        private static string ResourceName = "FormatHelperInstaller.Resources.VSTOFormatHelper.dll";

        [STAThread]
        static void Main(string[] args)
        {
            bool isUninstall = false;
            foreach (string a in args)
            {
                if (a.ToLower() == "/uninstall" || a.ToLower() == "-uninstall")
                {
                    isUninstall = true;
                    break;
                }
            }

            if (isUninstall)
            {
                DoUninstall();
            }
            else
            {
                DoInstall();
            }
        }

        static void DoInstall()
        {
            try
            {
                // Check Word is not running
                Process[] wordProcs = Process.GetProcessesByName("WINWORD");
                if (wordProcs.Length > 0)
                {
                    DialogResult dr = MessageBox.Show(
                        "检测到Word正在运行，需要关闭Word才能继续安装。是否关闭Word？",
                        "格式调整助手安装",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        foreach (Process p in wordProcs)
                        {
                            try { p.Kill(); } catch { }
                        }
                        Thread.Sleep(1000);
                    }
                    else if (dr == DialogResult.No)
                    {
                        // Continue anyway
                    }
                    else
                    {
                        return;
                    }
                }

                // Create install directory
                if (!Directory.Exists(InstallDir))
                {
                    Directory.CreateDirectory(InstallDir);
                }

                // Extract DLL from resources
                byte[] dllBytes = null;
                Assembly asm = Assembly.GetExecutingAssembly();
                foreach (string name in asm.GetManifestResourceNames())
                {
                    if (name == ResourceName || name.EndsWith("." + DllName))
                    {
                        using (Stream s = asm.GetManifestResourceStream(name))
                        {
                            dllBytes = new byte[s.Length];
                            s.Read(dllBytes, 0, dllBytes.Length);
                        }
                        break;
                    }
                }

                if (dllBytes == null)
                {
                    MessageBox.Show("无法找到内嵌的DLL资源", "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Write DLL
                File.WriteAllBytes(DllPath, dllBytes);

                // Register COM component with RegAsm
                string regAsmPath = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "RegAsm.exe");
                if (!File.Exists(regAsmPath))
                {
                    // Try framework64
                    regAsmPath = Path.Combine(
                        RuntimeEnvironment.GetRuntimeDirectory().Replace("\\Framework\\", "\\Framework64\\"),
                        "RegAsm.exe");
                }

                if (File.Exists(regAsmPath))
                {
                    Process regProc = new Process();
                    regProc.StartInfo.FileName = regAsmPath;
                    regProc.StartInfo.Arguments = "/codebase \"" + DllPath + "\"";
                    regProc.StartInfo.UseShellExecute = false;
                    regProc.StartInfo.RedirectStandardOutput = true;
                    regProc.StartInfo.RedirectStandardError = true;
                    regProc.StartInfo.CreateNoWindow = true;
                    regProc.Start();
                    regProc.WaitForExit(30000);
                }

                // Write Word Addins registry
                RegistryKey addinKey = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Office\Word\Addins\" + ProgId);
                addinKey.SetValue("FriendlyName", "格式调整助手");
                addinKey.SetValue("Description", "Word文档格式调整插件");
                addinKey.SetValue("LoadBehavior", 3);
                addinKey.Close();

                // Write uninstall registry
                RegistryKey uninstallKey = Registry.CurrentUser.CreateSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProgId);
                uninstallKey.SetValue("DisplayName", "格式调整助手");
                uninstallKey.SetValue("DisplayVersion", "3.0");
                uninstallKey.SetValue("Publisher", "FormatHelper");
                uninstallKey.SetValue("UninstallString", "\"" + DllPath.Replace(DllName, "SetupFormatHelper_v3.exe") + "\" /uninstall");
                uninstallKey.SetValue("NoModify", 1);
                uninstallKey.SetValue("NoRepair", 1);
                uninstallKey.Close();

                // Copy installer to install dir for uninstall
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                string exeName = Path.GetFileName(exePath);
                string destExe = Path.Combine(InstallDir, exeName);
                try { File.Copy(exePath, destExe, true); } catch { }

                // Update uninstall string to point to copied exe
                uninstallKey = Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProgId, true);
                if (uninstallKey != null)
                {
                    uninstallKey.SetValue("UninstallString", "\"" + destExe + "\" /uninstall");
                    uninstallKey.Close();
                }

                MessageBox.Show("安装成功！请重新启动Word查看\"格式调整\"选项卡。", "格式调整助手安装",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("安装失败: " + ex.Message + "\r\n\r\n" + ex.StackTrace,
                    "格式调整助手安装", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void DoUninstall()
        {
            try
            {
                // Check Word is not running
                Process[] wordProcs = Process.GetProcessesByName("WINWORD");
                if (wordProcs.Length > 0)
                {
                    DialogResult dr = MessageBox.Show(
                        "检测到Word正在运行，需要关闭Word才能继续卸载。是否关闭Word？",
                        "格式调整助手卸载",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        foreach (Process p in wordProcs)
                        {
                            try { p.Kill(); } catch { }
                        }
                        Thread.Sleep(1000);
                    }
                    else if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                }

                // Unregister COM
                string regAsmPath = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "RegAsm.exe");
                if (!File.Exists(regAsmPath))
                {
                    regAsmPath = Path.Combine(
                        RuntimeEnvironment.GetRuntimeDirectory().Replace("\\Framework\\", "\\Framework64\\"),
                        "RegAsm.exe");
                }

                if (File.Exists(regAsmPath) && File.Exists(DllPath))
                {
                    try
                    {
                        Process regProc = new Process();
                        regProc.StartInfo.FileName = regAsmPath;
                        regProc.StartInfo.Arguments = "/unregister \"" + DllPath + "\"";
                        regProc.StartInfo.UseShellExecute = false;
                        regProc.StartInfo.CreateNoWindow = true;
                        regProc.Start();
                        regProc.WaitForExit(30000);
                    }
                    catch { }
                }

                // Delete Word Addins registry
                try
                {
                    Registry.CurrentUser.DeleteSubKey(
                        @"Software\Microsoft\Office\Word\Addins\" + ProgId);
                }
                catch { }

                // Delete uninstall registry
                try
                {
                    Registry.CurrentUser.DeleteSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + ProgId);
                }
                catch { }

                // Delete files
                try
                {
                    if (Directory.Exists(InstallDir))
                    {
                        Directory.Delete(InstallDir, true);
                    }
                }
                catch { }

                MessageBox.Show("卸载成功！", "格式调整助手卸载",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Self-delete via batch
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                try
                {
                    string batPath = Path.Combine(Path.GetTempPath(), "uninstall_cleanup.bat");
                    File.WriteAllText(batPath,
                        "@echo off\r\n" +
                        "ping 127.0.0.1 -n 3 > nul\r\n" +
                        "del \"" + exePath + "\"\r\n" +
                        "del \"%~f0\"\r\n");
                    Process.Start(new ProcessStartInfo(batPath) { CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden });
                }
                catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show("卸载失败: " + ex.Message,
                    "格式调整助手卸载", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
