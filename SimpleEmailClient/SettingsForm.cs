using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SimpleEmailClient
{
    public partial class SettingsForm : Form
    {
        private CheckBox chkRegisterMailto;
        private CheckBox chkAutoLoadInbox;
        private CheckBox chkEnableNotifications;
        private Label lblNotificationStatus;
        private Button btnSave;
        private Button btnCancel;
        private TabControl tabControl;
        private TabPage tabGeneral, tabAbout;
        private Label lblAbout;

        public bool RegisterMailto { get; private set; }
        public bool AutoLoadInbox { get; private set; }
        public bool EnableNotifications { get; private set; }

        public SettingsForm(bool currentRegisterMailto, bool currentAutoLoadInbox, bool currentEnableNotifications)
        {
            InitializeComponent();
            LoadSettings(currentRegisterMailto, currentAutoLoadInbox, currentEnableNotifications);
        }

        private void InitializeComponent()
        {
            this.Text = "设置 - Simple Email Client";
            this.Size = new System.Drawing.Size(450, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            tabControl = new TabControl { Dock = DockStyle.Fill };
            tabGeneral = new TabPage("常规设置");
            tabAbout = new TabPage("关于");
            tabControl.TabPages.Add(tabGeneral);
            tabControl.TabPages.Add(tabAbout);

            // 常规设置页
            chkRegisterMailto = new CheckBox { Text = "注册 mailto 协议（设为默认邮件客户端）", AutoSize = true, Location = new System.Drawing.Point(20, 30), Width = 300 };
            chkAutoLoadInbox = new CheckBox { Text = "选中邮箱后自动获取收件箱", AutoSize = true, Location = new System.Drawing.Point(20, 70), Width = 250 };
            chkEnableNotifications = new CheckBox { Text = "允许系统通知（新邮件提醒）", AutoSize = true, Location = new System.Drawing.Point(20, 110), Width = 250 };
            lblNotificationStatus = new Label { Text = "", Location = new System.Drawing.Point(20, 140), AutoSize = true, ForeColor = System.Drawing.Color.Gray };

            btnSave = new Button { Text = "保存", Location = new System.Drawing.Point(260, 270), Size = new System.Drawing.Size(80, 30) };
            btnCancel = new Button { Text = "取消", Location = new System.Drawing.Point(350, 270), Size = new System.Drawing.Size(80, 30) };
            btnSave.Click += BtnSave_Click;
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            tabGeneral.Controls.Add(chkRegisterMailto);
            tabGeneral.Controls.Add(chkAutoLoadInbox);
            tabGeneral.Controls.Add(chkEnableNotifications);
            tabGeneral.Controls.Add(lblNotificationStatus);
            tabGeneral.Controls.Add(btnSave);
            tabGeneral.Controls.Add(btnCancel);

            // 关于页面
            lblAbout = new Label
            {
                Text = "Simple Email Client v1.0\nBy ASimpUsr\n\nhttps://github.com/ASimpUsr/SimpleEmailClient\n\nOpen with GPLv3",
                AutoSize = true,
                Location = new System.Drawing.Point(20, 20),
                Font = new System.Drawing.Font("微软雅黑", 9)
            };
            tabAbout.Controls.Add(lblAbout);

            this.Controls.Add(tabControl);
        }

        private void LoadSettings(bool registerMailto, bool autoLoadInbox, bool enableNotifications)
        {
            chkRegisterMailto.Checked = registerMailto;
            chkAutoLoadInbox.Checked = autoLoadInbox;
            chkEnableNotifications.Checked = enableNotifications;

            // 检测 Windows 版本
            if (!IsWindows10OrNewer())
            {
                chkEnableNotifications.Enabled = false;
                chkEnableNotifications.Checked = false;
                lblNotificationStatus.Text = "需要 Windows 10 或更高版本才能使用通知功能。";
            }
            else
            {
                chkEnableNotifications.Enabled = true;
                lblNotificationStatus.Text = "";
            }
        }

        private bool IsWindows10OrNewer()
        {
            var version = Environment.OSVersion.Version;
            return version.Major >= 10 && version.Minor >= 0;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            RegisterMailto = chkRegisterMailto.Checked;
            AutoLoadInbox = chkAutoLoadInbox.Checked;
            EnableNotifications = chkEnableNotifications.Checked && IsWindows10OrNewer();

            // 处理 mailto 协议注册/注销
            if (RegisterMailto)
                RegisterMailtoProtocol();
            else
                UnregisterMailtoProtocol();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void RegisterMailtoProtocol()
        {
            try
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                using (RegistryKey key = Registry.ClassesRoot.CreateSubKey("mailto"))
                {
                    key.SetValue("", "URL:MailTo Protocol");
                    key.SetValue("URL Protocol", "");
                    using (RegistryKey shellKey = key.CreateSubKey("shell"))
                    using (RegistryKey openKey = shellKey.CreateSubKey("open"))
                    using (RegistryKey commandKey = openKey.CreateSubKey("command"))
                    {
                        commandKey.SetValue("", $"\"{exePath}\" \"%1\"");
                    }
                }
                // 设置为默认程序
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice"))
                {
                    key.SetValue("ProgId", "mailto");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"注册 mailto 协议失败: {ex.Message}\n可能需要管理员权限。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UnregisterMailtoProtocol()
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKeyTree("mailto", false);
            }
            catch { }
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice", true))
                {
                    if (key != null) key.DeleteValue("ProgId", false);
                }
            }
            catch { }
        }
    }
}