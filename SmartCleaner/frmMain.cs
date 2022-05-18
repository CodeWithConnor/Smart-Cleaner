using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Security.Principal;
using System.Security.AccessControl;
using Microsoft.Win32;
using System.ServiceProcess;

namespace SmartCleaner
{
    public partial class frmMain : Form
    {
        
        // Recycle bin values
        enum RecycleFlags : int
        {
            SHRB_NOCONFIRMATION = 0x00000001,
            SHRB_NOPROGRESSUI = 0x00000002,
            SHRB_NOSOUND = 0x00000004,
        }

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        static extern uint SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, RecycleFlags dwFlags);

        // Initialize stopwatch
        private Stopwatch stopWatch;

        // Used for concatenating path strings
        string computerUsername = Environment.UserName;

        private void Form1_Load(object sender, EventArgs e)
        {
            stopWatch = new Stopwatch();
        }

        public frmMain()
        {
            InitializeComponent();

            // Adjust home panel size
            this.Height = 655;

            // Set current panel's visiblity to true and rest to false
            pnlHome.Visible = true;
            pnlFAQ.Visible = false;
            pnlLicense.Visible = false;
        }
        
        private void btnFAQ_Click(object sender, EventArgs e)
        {
            this.Height = 303;

            pnlFAQ.Visible = true;
            pnlHome.Visible = false;
            pnlLicense.Visible = false;
        }

        private void btnHome_Click(object sender, EventArgs e)
        {
            this.Height = 655;

            pnlHome.Visible = true;
            pnlFAQ.Visible = false;
            pnlLicense.Visible = false;
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            this.Height = 460;

            pnlLicense.Visible = true;
            pnlFAQ.Visible = false;
            pnlHome.Visible = false;
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void ElapsedTimeTimer_Tick(object sender, EventArgs e)
        {
            this.lblElapsedTimeCounter.Text = string.Format("{0:hh\\:mm\\:ss}", stopWatch.Elapsed);
        }

        private void btnCleanMyPC_Click(object sender, EventArgs e)
        {
            stopWatch.Start();
            this.Height = 558;
            prgProgressBar.Visible = true;

            // Junk file cleanup
            if (cbRecycleBin.Checked)
            {
                ClearRecycleBin();
            }
            if (cbWindowsTempFiles.Checked)
            {
                ClearTempFiles();
            }
            if (cbClipboard.Checked)
            {
                Clipboard.Clear();
            }
            if (cbApplicationCrashDumps.Checked)
            {
                ClearApplicationCrashDumps();
            }
            if (cbDNSCache.Checked)
            {
                ClearDNSCache();
            }
            if (cbThumbnailCache.Checked)
            {
                ClearThumbnailCache();
            }
            if (cbOldPrefetchData.Checked)
            {
                ClearOldPrefetchData();
            }
            if (cbMenuOrderCache.Checked)
            {
                ClearMenuOrderCache();
            }
            if (cbWindowsSizeLocationCache.Checked)
            {
                ClearWindowsSizeLocationCache();
            }
            if (cbChkdskFileFragments.Checked)
            {
                ClearChkdskFileFragments();
            }
            if (cbUserAssistHistory.Checked)
            {
                ClearUserAssistHistory();
            }
            if (cbWindowsLogFiles.Checked)
            {
                ClearWindowsLogFiles();
            }
            if (cbTrayNotificationsCache.Checked)
            {
                ClearTrayNotificationsCache();
            }
            if (cbWindowsErrorReports.Checked)
            {
                ClearWindowsErrorReports();
            }
            if (cbTaskbarJumpList.Checked)
            {
                ClearTaskbarJumpList();
            }
            if (cbMemoryDumps.Checked)
            {
                ClearMemoryDumps();
            }
            if (cbWindowsFontCache.Checked)
            {
                ClearWindowsFontCache();
            }
            if (cbEventLogs.Checked)
            {
                ClearEventLogs();
            }
            if (cbWindowsDefenderLogs.Checked)
            {
                ClearWindowsDefenderLogs();
            }
            if (cbWindowsIconCache.Checked)
            {
                ClearWindowsIconCache();
            }
            if (cbWindowsShortcuts.Checked)
            {
                ClearWindowsShortcuts();
            }

            // Browser cleanup
            if (cbGoogleChrome.Checked)
            {
                GoogleChromeCleanup();
            }
            if (cbInternetExplorer.Checked)
            {
                InternetExplorerCleanup();
            }
            if (cbMicrosoftEdge.Checked)
            {
                MicrosoftEdgeCleanup();
            }
            if (cbMozillaFirefox.Checked)
            {
                MozillaFirefoxCleanup();
            }
            if (cbOpera.Checked)
            {
                OperaCleanup();
            }

            // App cleanup
            if (cbWordpadHistory.Checked)
            {
                ClearWordpadHistory();
            }
        }

        public static void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }
            Directory.Delete(target_dir, false);
        }

        // Junk File Cleanup
        private void ClearRecycleBin()
        {
            // Empty recycle bin
            SHEmptyRecycleBin(IntPtr.Zero, null, 0);
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared recycled bin.");

            // Update progress bar
            //prgProgressBar.Value = 10;
        }

        private void ClearTempFiles()
        {
            Process.Start("cmd.exe", "/c del /s /q \"C:\\Windows\\Temp\"");
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared temp files.");
        }

        private void ClearApplicationCrashDumps()
        {
            Process.Start("cmd.exe", "/c del /s /q \"%LOCALAPPDATA%\\CrashDumps");
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared application crash dumps.");
        }

        private void ClearDNSCache()
        {
            Process.Start("cmd.exe", "ipconfig /flushdns");
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared DNS dache.");
        }

        private void ClearThumbnailCache()
        {
            try
            {
                string computerUsername = Environment.UserName; ;
                string path = @"C:\Users\" + computerUsername + @"\AppData\Local\Microsoft\Windows\Explorer";

                // Only delete .db files containing 'thumbcache' in their filenames
                string filesToDelete = @"*thumbcache*.db";

                // Save all file paths from directory in 'fileNames' array
                string[] fileList = System.IO.Directory.GetFiles(path, filesToDelete);

                // Ensure we have high enough privileges to delete file(s)
                File.SetAttributes(path, FileAttributes.Normal);

                foreach (string file in fileList)
                {
                    System.IO.File.Delete(file);
                }

                rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared thumbnail cache.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        private void ClearOldPrefetchData()
        {
            string prefetchPath = @"C:\Windows\Prefetch";

            string[] fileList = System.IO.Directory.GetFiles(prefetchPath);
            File.SetAttributes(prefetchPath, FileAttributes.Normal);

            foreach (string file in fileList)
            {
                System.IO.File.Delete(file);
            }

            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared old prefetch data.");
        }

        private void ClearMenuOrderCache()
        {
            string regPath = @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer";

            using (RegistryKey explorerKey =
                Registry.CurrentUser.OpenSubKey(regPath, writable: true))
            {
                if (explorerKey != null)
                {
                    explorerKey.DeleteSubKeyTree("MenuOrder");
                }
            }
        }

        private void ClearWindowsSizeLocationCache()
        {
            string regPath = @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer";

            // Delete both registry folders
            using (RegistryKey explorerKey =
                Registry.CurrentUser.OpenSubKey(regPath, writable: true))
            {
                if (explorerKey != null)
                {
                    explorerKey.DeleteSubKeyTree("StreamMRU");
                    explorerKey.DeleteSubKeyTree("Streams");
                }
            }

            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleaned up Windows size/location cache.");
        }

        private void ClearChkdskFileFragments()
        {
            Process.Start("cmd.exe", "/c del /s /q \"%LOCALAPPDATA%\\CrashDumps");
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared Application Crash Dumps.");
        }

        private void ClearUserAssistHistory()
        {
            string[] regPath = { @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\{5E6AB780-7743-11CF-A12B-00AA004AE837}", @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\{75048700-EF1F-11D0-9888-006097DEACF9}", @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\{CEBFF5CD-ACE2-4F4F-9178-9926F41749EA}", @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\UserAssist\{F4E57C4B-2036-45F0-A9AB-443BCFE33D9F}" };

            // Delete all registry folders
            foreach (var reg in regPath)
            {
                using (RegistryKey explorerKey =
                Registry.CurrentUser.OpenSubKey(reg, writable: true))
                {
                    if (explorerKey != null)
                    {
                        explorerKey.DeleteSubKeyTree("Count");
                    }
                }
            }
            rtxCleaningLogs.AppendText(Environment.NewLine + "Deleted user assist history.");
        }

        private void ClearWindowsLogFiles()
        {
            // Get current dir and go back two parent levels
            string currentDir = System.IO.Directory.GetCurrentDirectory();
            currentDir = currentDir.Substring(0, currentDir.Length - 10);

            // Path to script
            currentDir = currentDir + @"\Modules\Clear Windows Logs.cmd";

            // Execute script
            System.Diagnostics.Process.Start(currentDir);

            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared Windows log files.");
        }

        private void ClearTrayNotificationsCache()
        {
            // Path to registry folder
            string keyName = @"Software\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\TrayNotify";
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyName, true))
            {
                try
                {
                    key.DeleteValue("IconStreams");
                    key.DeleteValue("PastIconsStream");

                    rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared tray notifications cache.");
                }
                catch (System.ArgumentException)
                {
                    rtxCleaningLogs.AppendText(Environment.NewLine + "Skipping: Registry key(s) not found whilst attempting to clear tray notification cache.");
                }
            }
        }

        private void ClearWindowsErrorReports()
        {
            string[] regPaths = { @"C:\ProgramData\Microsoft\Windows\WER\ReportArchive\", @"C:\ProgramData\Microsoft\Windows\WER\ReportQueue\" };
            
            /*foreach (var reg in regPaths)
            {
                // Delete all files and folders in both directories
                System.IO.DirectoryInfo reg = new DirectoryInfo();

                foreach (FileInfo file in reg.GetFiles())
                {
                    file.Delete();
                }
                foreach (DirectoryInfo dir in reg.GetDirectories())
                {
                    dir.Delete(true);
                }
            }*/

            rtxCleaningLogs.AppendText(Environment.NewLine + "Feature temporarily removed.");
        }

        private void ClearTaskbarJumpList()
        {
            // Concatenate account username in directory path
            string computerUsername = Environment.UserName;
            System.IO.DirectoryInfo directory = new DirectoryInfo(@"C:\Users\" + computerUsername + @"\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations");

            // Delete all files and folders within directory
            foreach (FileInfo file in directory.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in directory.GetDirectories())
            {
                dir.Delete(true);
            }
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared taskbar jump list items.");
        }

        private void ClearMemoryDumps()
        {
            Process.Start("cmd.exe", @"del /f /s /q %systemroot%\Minidump\*.*");
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared memory dump files.");
        }

        private void ClearWindowsFontCache()
        {
            void StopService(string serviceName)
            {
                // Toggle service
                ServiceController sc = new ServiceController(serviceName);

                if ((sc.Status.Equals(ServiceControllerStatus.Running)))
                {
                    // Stop service
                    sc.Stop();

                    rtxCleaningLogs.AppendText(Environment.NewLine + "Stopping "+ serviceName + "...");

                    // Alow some time for the service to fully stop
                    Thread.Sleep(1000);

                    // Refresh and display current service status
                    sc.Refresh();
                    string serviceStatus = sc.Status.ToString();
                    rtxCleaningLogs.AppendText(Environment.NewLine + "The " + serviceName + " is now " + serviceStatus.ToLower() + ".");
                }
                else
                {
                    rtxCleaningLogs.AppendText(Environment.NewLine + "Skipping: " + serviceName + " isn't running.");
                }
            }
            StopService("Windows Font Cache Service");
            StopService("Windows Presentation Foundation Font Cache 3.0.0.0");

            // Clear the font cache
            try
            {
                int counter = 0;
                string path = @"C:\Windows\ServiceProfiles\LocalService\AppData\Local";

                // Only delete .db files containing 'thumbcache' in their filenames
                string filesToDelete = @"*FontCache*.dat";

                // Save all file paths from directory in 'fileList' array
                string[] fileList = System.IO.Directory.GetFiles(path, filesToDelete);

                // Ensure we have high enough privileges to delete file(s)
                File.SetAttributes(path, FileAttributes.Normal);

                foreach (string file in fileList)
                {
                    System.IO.File.Delete(file);
                    counter++;
                }

                if (counter == 0)
                {
                    // If there are no files to delete, skip
                    rtxCleaningLogs.AppendText(Environment.NewLine + "Skipping: There are no files available to delete in your font cache.");
                }
                else
                {
                    rtxCleaningLogs.AppendText(Environment.NewLine + "Deleted " + counter + " files from font cache.");
                }
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }

        }

        private void ClearEventLogs()
        {
            Process.Start("cmd.exe", "for /F \"tokens = *\" %1 in ('wevtutil.exe el') DO wevtutil.exe cl \" % 1\"");
            rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared event viewer logs");
        }

        private void ClearWindowsDefenderLogs()
        {
            try
            {
                string folder_path = @"C:\ProgramData\Microsoft\Windows Defender\Scans\History\Service";
                File.SetAttributes(folder_path, FileAttributes.Normal);

                // Delete history folder
                DeleteDirectory(folder_path);

                rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared Windows Defender history.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        private void ClearWindowsIconCache()
        {
            try
            {
                string computerUsername = Environment.UserName; ;
                string path = @"C:\Users\" + computerUsername + @"\AppData\Local\Microsoft\Windows\Explorer";

                // Only delete files containing 'iconcache' in their filename
                string filesToDelete = @"*iconcache*.db";

                // Save all file paths from directory in 'fileNames' array
                string[] fileList = System.IO.Directory.GetFiles(path, filesToDelete);

                // Ensure we have high enough privileges to delete file(s)
                File.SetAttributes(path, FileAttributes.Normal);

                foreach (string file in fileList)
                {
                    System.IO.File.Delete(file);
                }
                rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared icon cache.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        private void ClearWindowsShortcuts()
        {
            try
            {
                DeleteDirectory(@"C:\ProgramData\Microsoft\Windows\Start Menu\Programs");
                rtxCleaningLogs.AppendText(Environment.NewLine + "Cleared all shortcuts.");

            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        // Browser Cleanup
        private void InternetExplorerCleanup()
        {
            try
            {
                string computerUsername = Environment.UserName; ;

                // Delete cookies
                DeleteDirectory(@"C:\Documents and Settings\" + computerUsername + @"\Cookies\");

                // Delete history
                DeleteDirectory(@"C:\Documents and Settings\" + computerUsername + @"\Local Settings\History\History.IE5\");
                DeleteDirectory(@"C:\Documents and Settings\" + computerUsername + @"\Local Settings\History\History.IE5\MSHist01YYYYMMDDYYYYMMDD\");

                // Delete Cache
                DeleteDirectory(@"C:\Documents and Settings\" + computerUsername + @"\Local Settings\Temporary Internet Files\Content.IE5\");

                rtxCleaningLogs.AppendText(Environment.NewLine + "Deleted cookies, history and cache from Internet Explorer.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + "User account does not have sufficient privileges to access Internet Explorer data.");
            }
        }

        private void GoogleChromeCleanup()
        {
            try
            {
                string computerUsername = Environment.UserName; ;
                File.SetAttributes(@"C:\Users\" + computerUsername + @"\AppData\Local\Google\Chrome\User Data", FileAttributes.Normal);

                // Delete user data folder
                DeleteDirectory(@"C:\Users\" + computerUsername + @"\AppData\Local\Google\Chrome\User Data");

                rtxCleaningLogs.AppendText(Environment.NewLine + "Deleted all Google Chrome data.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        private void MicrosoftEdgeCleanup()
        {
            try
            {
                string computerUsername = Environment.UserName; ;
                File.SetAttributes(@"C:\Users\" + computerUsername + @"\AppData\Local\Microsoft\Edge\User Data\Default", FileAttributes.Normal);

                // Delete user data folder
                DeleteDirectory(@"C:\Users\" + computerUsername + @"\AppData\Local\Microsoft\Edge\User Data\Default");

                rtxCleaningLogs.AppendText(Environment.NewLine + "Deleted all Microsoft Edge data.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        private void MozillaFirefoxCleanup()
        {
            try
            {
                string computerUsername = Environment.UserName; ;
                File.SetAttributes(@"C:\Users\" + computerUsername + @"\AppData\Roaming\Mozilla\Firefox\Profiles\", FileAttributes.Normal);

                // Delete folder containing all user profiles
                DeleteDirectory(@"C:\Users\" + computerUsername + @"\AppData\Roaming\Mozilla\Firefox\Profiles\");

                rtxCleaningLogs.AppendText(Environment.NewLine + "Deleted all Mozilla Firefox profiles/data.");
            }
            catch (Exception e)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + e.ToString());
            }
        }

        private void OperaCleanup()
        {
            try
            {
                int counter = 0;
                string path = @"C:\Users\" + computerUsername + @"conno\AppData\Roaming\Opera Software\Opera Stable";

                // Only delete files containing 'thumbcache' in their filenames
                string filesToDelete = @"*Cookies*";

                // Save all file paths from directory in 'fileList' array
                string[] fileList = System.IO.Directory.GetFiles(path, filesToDelete);

                // Ensure we have high enough privileges to delete file(s)
                File.SetAttributes(path, FileAttributes.Normal);

                foreach (string file in fileList)
                {
                    System.IO.File.Delete(file);
                    counter++;
                }
            }
            catch (Exception DirectoryNotFoundException)
            {
                rtxCleaningLogs.AppendText(Environment.NewLine + "Opera is not installed.");
            }
        }

        // App cleanup
        private void ClearWordpadHistory()
        {
            // Path to Windows menu order cache registry key folder
            string regPath = @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Wordpad";

            using (RegistryKey explorerKey =
                Registry.CurrentUser.OpenSubKey(regPath, writable: true))
            {
                if (explorerKey != null)
                {
                    explorerKey.DeleteSubKeyTree("Recent File List");
                }
            }
        }

        private void ClearNotepadHistory()
        {
            // Path to Windows menu order cache registry key folder
            string regPath = @"Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Wordpad";

            using (RegistryKey explorerKey =
                Registry.CurrentUser.OpenSubKey(regPath, writable: true))
            {
                if (explorerKey != null)
                {
                    explorerKey.DeleteSubKeyTree("Recent File List");
                }
            }
        }
    }
}
