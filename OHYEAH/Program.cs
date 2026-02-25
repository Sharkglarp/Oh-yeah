using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace CatWallpaperInstaller
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm());
        }
    }

    internal sealed class InstallerForm : Form
    {
        private const int SPI_SETDESKWALLPAPER = 0x0014;
        private const int SPIF_UPDATEINIFILE = 0x0001;
        private const int SPIF_SENDWININICHANGE = 0x0002;
        private const string StartupSoundAlias = "cat_startup_sound";

        private readonly PictureBox previewBox;
        private readonly Button installButton;
        private Image previewImage;
        private string startupSoundTempPath;

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SystemParametersInfo(
            int uiAction,
            int uiParam,
            string pvParam,
            int fWinIni
        );

        [DllImport("winmm.dll", CharSet = CharSet.Unicode)]
        private static extern int mciSendString(
            string command,
            StringBuilder buffer,
            int bufferSize,
            IntPtr hwndCallback
        );

        public InstallerForm()
        {
            Text = " ";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(760, 560);

            previewBox = new PictureBox();
            previewBox.BorderStyle = BorderStyle.FixedSingle;
            previewBox.SizeMode = PictureBoxSizeMode.Zoom;
            previewBox.Location = new Point(20, 20);
            previewBox.Size = new Size(720, 480);

            Button cancelButton = new Button();
            cancelButton.Text = "forgive me";
            cancelButton.Size = new Size(110, 34);
            cancelButton.Location = new Point(500, 515);
            cancelButton.Click += delegate
            {
                Close();
            };

            installButton = new Button();
            installButton.Text = "ascend me";
            installButton.Size = new Size(130, 34);
            installButton.Location = new Point(620, 515);
            installButton.Click += InstallButton_Click;

            Controls.Add(previewBox);
            Controls.Add(cancelButton);
            Controls.Add(installButton);

            AcceptButton = installButton;
            CancelButton = cancelButton;

            LoadPreviewImage();
            Shown += delegate { PlayStartupSound(); };
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (previewImage != null)
                {
                    previewImage.Dispose();
                    previewImage = null;
                }

                StopStartupSound();
            }

            base.Dispose(disposing);
        }

        private void LoadPreviewImage()
        {
            try
            {
                Image embeddedImage = LoadEmbeddedImage();
                if (embeddedImage == null)
                {
                    throw new InvalidOperationException("Embedded wallpaper image was not found.");
                }

                previewImage = embeddedImage;
                previewBox.Image = previewImage;
            }
            catch (Exception ex)
            {
                installButton.Enabled = false;
                MessageBox.Show(
                    "Could not load image:\n\n" + ex.Message,
                    "Load Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private static Image LoadEmbeddedImage()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream("CatWallpaperInstaller.wallpaper"))
            {
                if (stream == null)
                {
                    return null;
                }

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    memoryStream.Position = 0;
                    using (Image loaded = Image.FromStream(memoryStream))
                    {
                        return new Bitmap(loaded);
                    }
                }
            }
        }

        private void PlayStartupSound()
        {
            try
            {
                string soundPath = EnsureStartupSoundFile();
                if (string.IsNullOrEmpty(soundPath))
                {
                    return;
                }

                mciSendString("close " + StartupSoundAlias, null, 0, IntPtr.Zero);
                int openResult = mciSendString(
                    "open \"" + soundPath + "\" type mpegvideo alias " + StartupSoundAlias,
                    null,
                    0,
                    IntPtr.Zero
                );

                if (openResult == 0)
                {
                    mciSendString("play " + StartupSoundAlias, null, 0, IntPtr.Zero);
                }
            }
            catch
            {
                // Ignore startup sound errors so the app still opens normally.
            }
        }

        private void StopStartupSound()
        {
            try
            {
                mciSendString("stop " + StartupSoundAlias, null, 0, IntPtr.Zero);
                mciSendString("close " + StartupSoundAlias, null, 0, IntPtr.Zero);
            }
            catch
            {
                // Ignore cleanup errors.
            }

            try
            {
                if (!string.IsNullOrEmpty(startupSoundTempPath) && File.Exists(startupSoundTempPath))
                {
                    File.SetAttributes(startupSoundTempPath, FileAttributes.Normal);
                    File.Delete(startupSoundTempPath);
                }

                startupSoundTempPath = null;
            }
            catch
            {
                // Ignore cleanup errors.
            }
        }

        private string EnsureStartupSoundFile()
        {
            if (!string.IsNullOrEmpty(startupSoundTempPath) && File.Exists(startupSoundTempPath))
            {
                return startupSoundTempPath;
            }

            byte[] soundBytes = LoadEmbeddedBinary("CatWallpaperInstaller.startup_mp3");
            if (soundBytes == null || soundBytes.Length == 0)
            {
                return null;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "cat-startup-" + Guid.NewGuid().ToString("N") + ".mp3"
            );

            File.WriteAllBytes(tempPath, soundBytes);
            try
            {
                File.SetAttributes(tempPath, FileAttributes.Hidden | FileAttributes.Temporary);
            }
            catch
            {
                // Ignore attribute errors.
            }

            startupSoundTempPath = tempPath;
            return startupSoundTempPath;
        }

        private static byte[] LoadEmbeddedBinary(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }

                using (MemoryStream memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    return memoryStream.ToArray();
                }
            }
        }

        private void InstallButton_Click(object sender, EventArgs e)
        {
            if (previewImage == null)
            {
                MessageBox.Show(
                    "No image is loaded.",
                    "Image Missing",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            installButton.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                string installedWallpaperPath = SaveWallpaperAsBmp(previewImage);
                SetWallpaperStyleFill();

                bool success = SystemParametersInfo(
                    SPI_SETDESKWALLPAPER,
                    0,
                    installedWallpaperPath,
                    SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE
                );

                if (!success)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new InvalidOperationException(
                        "Windows rejected the wallpaper update. Error code: " + errorCode
                    );
                }

                MessageBox.Show(
                    "Wallpaper installed successfully.",
                    "Done",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Install failed:\n\n" + ex.Message,
                    "Install Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                Cursor = Cursors.Default;
                installButton.Enabled = true;
            }
        }

        private static string SaveWallpaperAsBmp(Image image)
        {
            string commonData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string installDirectory = Path.Combine(commonData, "CatWallpaperInstaller");
            Directory.CreateDirectory(installDirectory);

            string wallpaperPath = Path.Combine(installDirectory, "cat-wallpaper.bmp");
            using (Bitmap bmp = new Bitmap(image))
            {
                bmp.Save(wallpaperPath, ImageFormat.Bmp);
            }

            return wallpaperPath;
        }

        private static void SetWallpaperStyleFill()
        {
            RegistryKey desktopKey = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
            if (desktopKey == null)
            {
                return;
            }

            desktopKey.SetValue("WallpaperStyle", "10");
            desktopKey.SetValue("TileWallpaper", "0");
            desktopKey.Close();
        }
    }
}
