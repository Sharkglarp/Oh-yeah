using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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

        private readonly PictureBox previewBox;
        private readonly Label statusLabel;
        private readonly Button installButton;
        private Image previewImage;

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SystemParametersInfo(
            int uiAction,
            int uiParam,
            string pvParam,
            int fWinIni
        );

        public InstallerForm()
        {
            Text = "Cat Wallpaper Setup";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(760, 560);

            Panel headerPanel = new Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 96;
            headerPanel.BackColor = Color.FromArgb(238, 243, 248);

            Label titleLabel = new Label();
            titleLabel.Text = "Cat Wallpaper Installer";
            titleLabel.Font = new Font("Segoe UI", 16f, FontStyle.Bold);
            titleLabel.AutoSize = true;
            titleLabel.Location = new Point(20, 18);

            Label subtitleLabel = new Label();
            subtitleLabel.Text =
                "Preview the image, then click Install Wallpaper to set it as your home screen.";
            subtitleLabel.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
            subtitleLabel.AutoSize = true;
            subtitleLabel.Location = new Point(22, 56);

            headerPanel.Controls.Add(titleLabel);
            headerPanel.Controls.Add(subtitleLabel);

            previewBox = new PictureBox();
            previewBox.BorderStyle = BorderStyle.FixedSingle;
            previewBox.SizeMode = PictureBoxSizeMode.Zoom;
            previewBox.Location = new Point(20, 112);
            previewBox.Size = new Size(720, 360);

            statusLabel = new Label();
            statusLabel.Text = "Loading wallpaper image...";
            statusLabel.Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            statusLabel.AutoSize = false;
            statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusLabel.Location = new Point(20, 478);
            statusLabel.Size = new Size(720, 40);

            Button cancelButton = new Button();
            cancelButton.Text = "Cancel";
            cancelButton.Size = new Size(110, 34);
            cancelButton.Location = new Point(510, 520);
            cancelButton.Click += delegate
            {
                Close();
            };

            installButton = new Button();
            installButton.Text = "Install Wallpaper";
            installButton.Size = new Size(130, 34);
            installButton.Location = new Point(630, 520);
            installButton.Click += InstallButton_Click;

            Controls.Add(headerPanel);
            Controls.Add(previewBox);
            Controls.Add(statusLabel);
            Controls.Add(cancelButton);
            Controls.Add(installButton);

            AcceptButton = installButton;
            CancelButton = cancelButton;

            LoadPreviewImage();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && previewImage != null)
            {
                previewImage.Dispose();
                previewImage = null;
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
                statusLabel.Text = "Ready. Click Install Wallpaper to apply this picture.";
            }
            catch (Exception ex)
            {
                installButton.Enabled = false;
                statusLabel.Text = "Could not load image: " + ex.Message;
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

                statusLabel.Text = "Installed successfully. Your home screen wallpaper is now updated.";
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
