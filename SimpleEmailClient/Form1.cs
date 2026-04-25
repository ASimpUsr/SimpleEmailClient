using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;

namespace SimpleEmailClient
{
    public partial class Form1 : Form
    {
        // Graph 常量
        private const string ClientId = "7fb58687-b9e5-49e4-835b-30205891e533";
        private const string Authority = "https://login.microsoftonline.com/common";
        private static readonly string[] GraphScopes = { "Mail.Send", "offline_access", "User.Read", "Mail.Read", "Mail.ReadWrite" };

        // UI 控件
        private SplitContainer mainSplit;                // 左：账户面板，右：主内容
        private Panel leftPanel;
        private ListBox accountListBox;
        private Button btnAddAccount, btnSelectAccount, btnRemoveAccount, btnSettings;
        private TabControl mainTabControl;
        private TabPage inboxPage, sendPage;

        // 收件箱内部布局：左中右再分割
        private SplitContainer inboxMainSplit;            // 水平分割：左(文件夹树) + 右(邮件列表+预览)
        private TreeView folderTree;
        private SplitContainer mailVerticalSplit;         // 垂直分割：上(邮件列表) + 下(预览)
        private DataGridView inboxGrid;
        private WebView2 webView;
        private Panel previewPanel;
        private Button btnClosePreview;
        private Panel mailToolbar;
        private Button btnRefreshInbox, btnDeleteSelected, btnMoveSelected;
        private ComboBox moveToFolderCombo;
        private Label lblInboxStatus;

        // 发送邮件页面控件
        private Label lblCurrentAccount;
        private TextBox txtTo, txtSubject;
        private RichTextBox txtBody;
        private Button btnSend;

        // 数据
        private List<EmailAccount> accounts = new List<EmailAccount>();
        private EmailAccount currentAccount = null;
        private string graphAccessToken;
        private string graphRefreshToken;
        private string userEmail;
        private bool useGraph = false;
        private readonly HttpClient httpClient = new HttpClient();

        // 文件夹与邮件缓存
        private string currentFolderId = "inbox";
        private List<dynamic> currentMessages = new List<dynamic>();
        private Dictionary<string, string> folderIdNameMap = new Dictionary<string, string>(); // Graph

        // 通知
        private NotifyIcon notifyIcon;
        private bool notificationsEnabled = false;

        // 用户设置
        private UserSettings userSettings;
        private string settingsPath;

        // WebView2 检测
        private static bool webView2Prompted = false;

        // 加密
        private static byte[] Protect(byte[] data) => ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
        private static byte[] Unprotect(byte[] data) => ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
        private string EncryptString(string plain) => Convert.ToBase64String(Protect(Encoding.UTF8.GetBytes(plain)));
        private string DecryptString(string cipher) => Encoding.UTF8.GetString(Unprotect(Convert.FromBase64String(cipher)));

        private class UserSettings
        {
            public bool RegisterMailto { get; set; } = false;
            public bool AutoLoadInbox { get; set; } = false;
            public bool EnableNotifications { get; set; } = true;
        }

        private class EmailAccount
        {
            public string Email { get; set; }
            public bool UseGraph { get; set; }
            public string SmtpServer { get; set; }
            public int SmtpPort { get; set; }
            public bool SmtpUseTls { get; set; }
            public string EncryptedPassword { get; set; }
            public string ImapServer { get; set; }
            public int ImapPort { get; set; }
            public bool ImapUseSsl { get; set; }
            public string GraphRefreshToken { get; set; }
        }

        public Form1()
        {
            this.Text = "Simple Email Client";
            this.Size = new Size(1200, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormClosing += Form1_FormClosing;

            LoadUserSettings();
            ApplyMailtoRegistration();
            BuildUI();
            LoadAccounts();
            UpdateAccountListUI();
            EnableInboxAndSend(false);
            lblInboxStatus.Text = "请从左侧选择一个账户";
            if (notificationsEnabled) InitNotifyIcon();
        }

        private void LoadUserSettings()
        {
            settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SimpleEmailClient", "settings.json");
            if (File.Exists(settingsPath))
            {
                try
                {
                    string json = File.ReadAllText(settingsPath);
                    userSettings = JsonSerializer.Deserialize<UserSettings>(json);
                }
                catch { userSettings = new UserSettings(); }
            }
            else
            {
                userSettings = new UserSettings();
                if (!Directory.Exists(Path.GetDirectoryName(settingsPath)))
                    Directory.CreateDirectory(Path.GetDirectoryName(settingsPath));
                SaveUserSettings();
            }
            notificationsEnabled = userSettings.EnableNotifications && IsWindows10OrNewer();
        }

        private void SaveUserSettings()
        {
            string json = JsonSerializer.Serialize(userSettings);
            File.WriteAllText(settingsPath, json);
        }

        private void ApplyMailtoRegistration()
        {
            if (userSettings.RegisterMailto)
                RegisterMailtoProtocol();
            else
                UnregisterMailtoProtocol();
        }

        private void RegisterMailtoProtocol()
        {
            try
            {
                string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                using (var key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey("mailto"))
                {
                    key.SetValue("", "URL:MailTo Protocol");
                    key.SetValue("URL Protocol", "");
                    using (var shellKey = key.CreateSubKey("shell"))
                    using (var openKey = shellKey.CreateSubKey("open"))
                    using (var commandKey = openKey.CreateSubKey("command"))
                    {
                        commandKey.SetValue("", $"\"{exePath}\" \"%1\"");
                    }
                }
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice"))
                {
                    key.SetValue("ProgId", "mailto");
                }
            }
            catch { }
        }

        private void UnregisterMailtoProtocol()
        {
            try { Microsoft.Win32.Registry.ClassesRoot.DeleteSubKeyTree("mailto", false); } catch { }
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice", true))
                {
                    if (key != null) key.DeleteValue("ProgId", false);
                }
            }
            catch { }
        }

        private bool IsWindows10OrNewer()
        {
            var version = Environment.OSVersion.Version;
            return version.Major >= 10 && version.Minor >= 0;
        }

        private void InitNotifyIcon()
        {
            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = SystemIcons.Information;
            notifyIcon.Visible = true;
            notifyIcon.BalloonTipTitle = "Simple Email Client";
            notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
        }

        private void ShowNotification(string title, string text)
        {
            if (notificationsEnabled && notifyIcon != null)
                notifyIcon.ShowBalloonTip(3000, title, text, ToolTipIcon.Info);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            notifyIcon?.Dispose();
            httpClient.Dispose();
        }

        public void SetMailto(string mailtoLink)
        {
            try
            {
                var uri = new Uri(mailtoLink);
                string to = uri.LocalPath;
                var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
                string subject = query["subject"];
                string body = query["body"];
                if (!string.IsNullOrEmpty(to))
                {
                    txtTo.Text = to;
                    if (!string.IsNullOrEmpty(subject)) txtSubject.Text = subject;
                    if (!string.IsNullOrEmpty(body)) txtBody.Text = body;
                    mainTabControl.SelectedTab = sendPage;
                }
            }
            catch { }
        }

        // ---------- UI 构建 ----------
        private void BuildUI()
        {
            mainSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 180, FixedPanel = FixedPanel.Panel1 };
            this.Controls.Add(mainSplit);

            // 左侧账户面板
            leftPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };
            var leftLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4, ColumnCount = 1 };
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            leftLayout.Controls.Add(new Label { Text = "邮箱账户", Font = new Font("微软雅黑", 10, FontStyle.Bold), Dock = DockStyle.Fill }, 0, 0);
            accountListBox = new ListBox { Dock = DockStyle.Fill, Font = new Font("微软雅黑", 9) };
            leftLayout.Controls.Add(accountListBox, 0, 1);

            var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 0) };
            btnAddAccount = new Button { Text = "添加", Width = 100, Height = 32 };
            btnSelectAccount = new Button { Text = "选择", Width = 100, Height = 32 };
            btnRemoveAccount = new Button { Text = "移除", Width = 100, Height = 32 };
            btnPanel.Controls.Add(btnAddAccount);
            btnPanel.Controls.Add(btnSelectAccount);
            btnPanel.Controls.Add(btnRemoveAccount);
            leftLayout.Controls.Add(btnPanel, 0, 2);

            btnSettings = new Button { Text = "⚙ 设置", Width = 100, Height = 32, Anchor = AnchorStyles.Top | AnchorStyles.Right };
            leftLayout.Controls.Add(btnSettings, 0, 3);
            leftPanel.Controls.Add(leftLayout);
            mainSplit.Panel1.Controls.Add(leftPanel);

            // 右侧标签页
            mainTabControl = new TabControl { Dock = DockStyle.Fill };
            mainSplit.Panel2.Controls.Add(mainTabControl);
            inboxPage = new TabPage("收件箱");
            sendPage = new TabPage("发送邮件");
            mainTabControl.TabPages.Add(inboxPage);
            mainTabControl.TabPages.Add(sendPage);

            // ---------- 收件箱页面（文件夹树 + 邮件列表 + 预览）----------
            // 水平分割：左边文件夹树，右边邮件区域
            inboxMainSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Vertical };
            inboxMainSplit.SplitterDistance = 220;
            inboxMainSplit.FixedPanel = FixedPanel.Panel1;
            inboxPage.Controls.Add(inboxMainSplit);

            // 左边：文件夹树
            folderTree = new TreeView { Dock = DockStyle.Fill, Font = new Font("微软雅黑", 9) };
            folderTree.AfterSelect += FolderTree_AfterSelect;
            inboxMainSplit.Panel1.Controls.Add(folderTree);

            // 右边：邮件列表 + 预览（垂直分割）
            mailVerticalSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal };
            mailVerticalSplit.SplitterDistance = 380;
            inboxMainSplit.Panel2.Controls.Add(mailVerticalSplit);

            // 上方：邮件列表 + 工具栏
            var mailListPanel = new Panel { Dock = DockStyle.Fill };
            inboxGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };
            inboxGrid.Columns.Add("From", "发件人");
            inboxGrid.Columns.Add("Subject", "主题");
            inboxGrid.Columns.Add("Received", "接收时间");
            inboxGrid.Columns[0].Width = 200;
            inboxGrid.Columns[1].Width = 400;
            inboxGrid.CellDoubleClick += InboxGrid_CellDoubleClick;
            mailListPanel.Controls.Add(inboxGrid);

            mailToolbar = new Panel { Dock = DockStyle.Top, Height = 40 };
            btnRefreshInbox = new Button { Text = "刷新", Location = new Point(10, 5), Size = new Size(80, 28) };
            btnDeleteSelected = new Button { Text = "删除", Location = new Point(100, 5), Size = new Size(60, 28) };
            btnMoveSelected = new Button { Text = "移动到", Location = new Point(170, 5), Size = new Size(60, 28) };
            moveToFolderCombo = new ComboBox { Location = new Point(240, 5), Size = new Size(100, 28), DropDownStyle = ComboBoxStyle.DropDownList };
            moveToFolderCombo.Items.AddRange(new[] { "垃圾邮件", "已删除邮件" });
            moveToFolderCombo.SelectedIndex = 0;
            lblInboxStatus = new Label { Text = "", Location = new Point(360, 8), AutoSize = true };
            mailToolbar.Controls.Add(btnRefreshInbox);
            mailToolbar.Controls.Add(btnDeleteSelected);
            mailToolbar.Controls.Add(btnMoveSelected);
            mailToolbar.Controls.Add(moveToFolderCombo);
            mailToolbar.Controls.Add(lblInboxStatus);
            mailListPanel.Controls.Add(mailToolbar);
            mailVerticalSplit.Panel1.Controls.Add(mailListPanel);

            // 下方：邮件预览区域（可关闭）
            previewPanel = new Panel { Dock = DockStyle.Fill };
            btnClosePreview = new Button { Text = "关闭预览", Dock = DockStyle.Top, Height = 30, Visible = false };
            btnClosePreview.Click += (s, e) => { previewPanel.Visible = false; btnClosePreview.Visible = false; };
            webView = new WebView2 { Dock = DockStyle.Fill };
            previewPanel.Controls.Add(webView);
            previewPanel.Controls.Add(btnClosePreview);
            mailVerticalSplit.Panel2.Controls.Add(previewPanel);
            previewPanel.Visible = false;

            // ---------- 发送邮件页面 ----------
            var sendLayout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(15), AutoSize = true };
            sendLayout.RowCount = 6;
            sendLayout.ColumnCount = 2;
            for (int i = 0; i < 6; i++) sendLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            sendLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            sendLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            lblCurrentAccount = new Label { Text = "未选择账户", ForeColor = Color.Blue, Font = new Font("微软雅黑", 9, FontStyle.Bold), AutoSize = true };
            sendLayout.Controls.Add(lblCurrentAccount, 0, 0);
            sendLayout.SetColumnSpan(lblCurrentAccount, 2);

            sendLayout.Controls.Add(new Label { Text = "收件人:", AutoSize = true }, 0, 1);
            txtTo = new TextBox { Width = 400 };
            sendLayout.Controls.Add(txtTo, 1, 1);

            sendLayout.Controls.Add(new Label { Text = "主题:", AutoSize = true }, 0, 2);
            txtSubject = new TextBox { Width = 400 };
            sendLayout.Controls.Add(txtSubject, 1, 2);

            sendLayout.Controls.Add(new Label { Text = "内容:", AutoSize = true }, 0, 3);
            txtBody = new RichTextBox { Height = 200 };
            sendLayout.Controls.Add(txtBody, 1, 3);

            btnSend = new Button { Text = "发送", Width = 100, Height = 35 };
            var btnPanelSend = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight };
            btnPanelSend.Controls.Add(btnSend);
            sendLayout.Controls.Add(btnPanelSend, 1, 4);

            sendPage.Controls.Add(sendLayout);

            // 事件绑定
            btnAddAccount.Click += BtnAddAccount_Click;
            btnSelectAccount.Click += BtnSelectAccount_Click;
            btnRemoveAccount.Click += BtnRemoveAccount_Click;
            btnSettings.Click += BtnSettings_Click;
            btnSend.Click += BtnSend_Click;
            btnRefreshInbox.Click += BtnRefreshInbox_Click;
            btnDeleteSelected.Click += BtnDeleteSelected_Click;
            btnMoveSelected.Click += BtnMoveSelected_Click;
        }

        private void EnableInboxAndSend(bool enabled)
        {
            folderTree.Enabled = enabled;
            inboxGrid.Enabled = enabled;
            btnRefreshInbox.Enabled = enabled;
            btnDeleteSelected.Enabled = enabled;
            btnMoveSelected.Enabled = enabled;
            txtTo.Enabled = enabled;
            txtSubject.Enabled = enabled;
            txtBody.Enabled = enabled;
            btnSend.Enabled = enabled;
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            var settingsForm = new SettingsForm(userSettings.RegisterMailto, userSettings.AutoLoadInbox, userSettings.EnableNotifications);
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                bool oldNotifications = userSettings.EnableNotifications;
                userSettings.RegisterMailto = settingsForm.RegisterMailto;
                userSettings.AutoLoadInbox = settingsForm.AutoLoadInbox;
                userSettings.EnableNotifications = settingsForm.EnableNotifications;
                SaveUserSettings();
                ApplyMailtoRegistration();
                if (oldNotifications != userSettings.EnableNotifications)
                {
                    notificationsEnabled = userSettings.EnableNotifications && IsWindows10OrNewer();
                    if (notificationsEnabled && notifyIcon == null) InitNotifyIcon();
                    else if (!notificationsEnabled && notifyIcon != null) { notifyIcon.Dispose(); notifyIcon = null; }
                }
            }
        }

        // ---------- 账户管理 ----------
        private void LoadAccounts()
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".seser_email");
            if (!File.Exists(path)) return;
            try
            {
                byte[] encrypted = File.ReadAllBytes(path);
                byte[] jsonBytes = Unprotect(encrypted);
                string json = Encoding.UTF8.GetString(jsonBytes);
                accounts = JsonSerializer.Deserialize<List<EmailAccount>>(json);
            }
            catch { accounts = new List<EmailAccount>(); }
        }

        private void SaveAccounts()
        {
            string json = JsonSerializer.Serialize(accounts);
            byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
            byte[] encrypted = Protect(jsonBytes);
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".seser_email");
            File.WriteAllBytes(path, encrypted);
        }

        private void UpdateAccountListUI()
        {
            accountListBox.Items.Clear();
            foreach (var acc in accounts)
                accountListBox.Items.Add(acc.Email);
        }

        private void BtnAddAccount_Click(object sender, EventArgs e)
        {
            var dialog = new Form
            {
                Text = "添加邮箱账户",
                Size = new Size(500, 500),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent
            };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            layout.RowCount = 8;
            layout.ColumnCount = 2;

            var loginTypeGroup = new GroupBox { Text = "登录方式", Height = 80, Dock = DockStyle.Fill };
            var radioGraph = new RadioButton { Text = "Graph API (Outlook/Hotmail)", Checked = true, Location = new Point(10, 20), AutoSize = true };
            var radioSmtp = new RadioButton { Text = "自行配置 SMTP/IMAP", Location = new Point(10, 50), AutoSize = true };
            loginTypeGroup.Controls.Add(radioGraph);
            loginTypeGroup.Controls.Add(radioSmtp);
            layout.Controls.Add(loginTypeGroup, 0, 0);
            layout.SetColumnSpan(loginTypeGroup, 2);

            layout.Controls.Add(new Label { Text = "邮箱地址:", AutoSize = true }, 0, 1);
            var txtEmailAdd = new TextBox();
            layout.Controls.Add(txtEmailAdd, 1, 1);

            var groupSmtp = new GroupBox { Text = "SMTP 发送设置", Height = 130, Dock = DockStyle.Fill, Visible = false };
            var smtpLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 4 };
            var txtSmtpServer = new TextBox();
            var txtSmtpPort = new TextBox { Text = "587" };
            var chkSmtpTls = new CheckBox { Text = "启用 TLS/SSL", Checked = true };
            var txtPwd = new TextBox { PasswordChar = '*' };
            smtpLayout.Controls.Add(new Label { Text = "服务器:" }, 0, 0);
            smtpLayout.Controls.Add(txtSmtpServer, 1, 0);
            smtpLayout.Controls.Add(new Label { Text = "端口:" }, 0, 1);
            smtpLayout.Controls.Add(txtSmtpPort, 1, 1);
            smtpLayout.Controls.Add(chkSmtpTls, 1, 2);
            smtpLayout.Controls.Add(new Label { Text = "密码/授权码:" }, 0, 3);
            smtpLayout.Controls.Add(txtPwd, 1, 3);
            groupSmtp.Controls.Add(smtpLayout);
            layout.Controls.Add(groupSmtp, 0, 2);
            layout.SetColumnSpan(groupSmtp, 2);

            var groupImap = new GroupBox { Text = "IMAP 接收设置", Height = 100, Dock = DockStyle.Fill, Visible = false };
            var imapLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3 };
            var txtImapServer = new TextBox();
            var txtImapPort = new TextBox { Text = "993" };
            var chkImapSsl = new CheckBox { Text = "使用 SSL", Checked = true };
            imapLayout.Controls.Add(new Label { Text = "服务器:" }, 0, 0);
            imapLayout.Controls.Add(txtImapServer, 1, 0);
            imapLayout.Controls.Add(new Label { Text = "端口:" }, 0, 1);
            imapLayout.Controls.Add(txtImapPort, 1, 1);
            imapLayout.Controls.Add(chkImapSsl, 1, 2);
            groupImap.Controls.Add(imapLayout);
            layout.Controls.Add(groupImap, 0, 3);
            layout.SetColumnSpan(groupImap, 2);

            var btnGraphLogin = new Button { Text = "使用 Graph 登录", Width = 150, Visible = true };
            layout.Controls.Add(btnGraphLogin, 1, 4);

            var btnOk = new Button { Text = "保存", Width = 80 };
            var btnCancel = new Button { Text = "取消", Width = 80 };
            var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft };
            btnPanel.Controls.Add(btnOk);
            btnPanel.Controls.Add(btnCancel);
            layout.Controls.Add(btnPanel, 1, 6);
            dialog.Controls.Add(layout);

            radioGraph.CheckedChanged += (s, ev) =>
            {
                groupSmtp.Visible = false;
                groupImap.Visible = false;
                btnGraphLogin.Visible = true;
                txtPwd.Enabled = false;
            };
            radioSmtp.CheckedChanged += (s, ev) =>
            {
                groupSmtp.Visible = true;
                groupImap.Visible = true;
                btnGraphLogin.Visible = false;
                txtPwd.Enabled = true;
            };

            btnGraphLogin.Click += async (sender2, e2) =>
            {
                var result = await PerformGraphLogin();
                if (result.success)
                {
                    txtEmailAdd.Text = result.email;
                    txtPwd.Text = "";
                    MessageBox.Show("Graph 登录成功！邮箱已自动填充。");
                }
                else
                {
                    MessageBox.Show("Graph 登录失败！");
                }
            };

            btnOk.Click += async (sender2, e2) =>
            {
                if (string.IsNullOrWhiteSpace(txtEmailAdd.Text))
                {
                    MessageBox.Show("请输入邮箱！");
                    return;
                }
                bool useGraphLogin = radioGraph.Checked;
                var acc = new EmailAccount { Email = txtEmailAdd.Text, UseGraph = useGraphLogin };

                if (useGraphLogin)
                {
                    if (string.IsNullOrEmpty(graphRefreshToken))
                    {
                        MessageBox.Show("请先点击“使用 Graph 登录”按钮完成登录");
                        return;
                    }
                    acc.GraphRefreshToken = graphRefreshToken;
                    acc.Email = userEmail ?? txtEmailAdd.Text;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(txtSmtpServer.Text) || string.IsNullOrWhiteSpace(txtPwd.Text))
                    {
                        MessageBox.Show("请填写 SMTP 服务器和密码/授权码");
                        return;
                    }
                    acc.SmtpServer = txtSmtpServer.Text;
                    acc.SmtpPort = int.Parse(txtSmtpPort.Text);
                    acc.SmtpUseTls = chkSmtpTls.Checked;
                    acc.EncryptedPassword = EncryptString(txtPwd.Text);
                    if (string.IsNullOrWhiteSpace(txtImapServer.Text))
                    {
                        MessageBox.Show("请填写 IMAP 服务器地址（用于接收邮件）");
                        return;
                    }
                    acc.ImapServer = txtImapServer.Text;
                    acc.ImapPort = int.Parse(txtImapPort.Text);
                    acc.ImapUseSsl = chkImapSsl.Checked;
                }
                accounts.Add(acc);
                SaveAccounts();
                UpdateAccountListUI();
                dialog.Close();
            };
            btnCancel.Click += (sender2, e2) => dialog.Close();
            dialog.ShowDialog(this);
        }

        private async void BtnSelectAccount_Click(object sender, EventArgs e)
        {
            if (accountListBox.SelectedItem == null) return;
            string email = accountListBox.SelectedItem.ToString();
            currentAccount = accounts.Find(a => a.Email == email);
            if (currentAccount == null) return;

            useGraph = currentAccount.UseGraph;
            if (useGraph)
            {
                graphRefreshToken = currentAccount.GraphRefreshToken;
                bool refreshed = await RefreshGraphTokenIfNeededAsync();
                if (!refreshed)
                {
                    MessageBox.Show("Graph token 刷新失败，请重新添加账户。");
                    return;
                }
                userEmail = currentAccount.Email;
                await LoadFolderTree();
            }
            else
            {
                await LoadImapFolderTree();
            }

            EnableInboxAndSend(true);
            lblCurrentAccount.Text = $"当前账户：{currentAccount.Email}";

            if (userSettings.AutoLoadInbox)
            {
                TreeNode inboxNode = FindFolderNode("inbox", true);
                if (inboxNode != null)
                {
                    folderTree.SelectedNode = inboxNode;
                    currentFolderId = inboxNode.Tag.ToString();
                    await LoadMessagesForFolder(currentFolderId);
                }
                else
                {
                    await LoadMessagesForFolder(currentFolderId);
                }
            }
            else
            {
                folderTree.Nodes.Clear();
                lblInboxStatus.Text = $"已选择账户 {currentAccount.Email}，请从左侧文件夹树选择文件夹加载邮件。";
                inboxGrid.Rows.Clear();
                currentMessages.Clear();
                previewPanel.Visible = false;
                btnClosePreview.Visible = false;
            }
        }

        private void BtnRemoveAccount_Click(object sender, EventArgs e)
        {
            if (accountListBox.SelectedItem == null) return;
            string email = accountListBox.SelectedItem.ToString();
            if (MessageBox.Show($"确定要移除账户 {email} 吗？", "确认", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                accounts.RemoveAll(a => a.Email == email);
                SaveAccounts();
                UpdateAccountListUI();
                if (currentAccount != null && currentAccount.Email == email)
                {
                    currentAccount = null;
                    useGraph = false;
                    graphAccessToken = null;
                    graphRefreshToken = null;
                    EnableInboxAndSend(false);
                    folderTree.Nodes.Clear();
                    inboxGrid.Rows.Clear();
                    currentMessages.Clear();
                    previewPanel.Visible = false;
                    btnClosePreview.Visible = false;
                    lblInboxStatus.Text = "请从左侧选择一个账户";
                    lblCurrentAccount.Text = "未选择账户";
                }
            }
        }

        private TreeNode FindFolderNode(string folderId, bool isGraph)
        {
            foreach (TreeNode node in folderTree.Nodes)
            {
                if (node.Tag.ToString() == folderId) return node;
                var child = FindChildNode(node, folderId);
                if (child != null) return child;
            }
            return null;
        }

        private TreeNode FindChildNode(TreeNode parent, string folderId)
        {
            foreach (TreeNode child in parent.Nodes)
            {
                if (child.Tag.ToString() == folderId) return child;
                var deep = FindChildNode(child, folderId);
                if (deep != null) return deep;
            }
            return null;
        }

        // ---------- 文件夹树加载 ----------
        private async Task LoadFolderTree()
        {
            folderTree.Nodes.Clear();
            if (!useGraph) return;
            if (!await RefreshGraphTokenIfNeededAsync()) return;

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
            var response = await client.GetAsync("https://graph.microsoft.com/v1.0/me/mailfolders?$top=100");
            if (!response.IsSuccessStatusCode) return;
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var roots = doc.RootElement.GetProperty("value");
            folderIdNameMap.Clear();
            var rootNodes = new List<TreeNode>();
            foreach (var folder in roots.EnumerateArray())
            {
                string id = folder.GetProperty("id").GetString();
                string name = folder.GetProperty("displayName").GetString();
                folderIdNameMap[id] = name;
                TreeNode node = new TreeNode(name) { Tag = id };
                rootNodes.Add(node);
            }
            folderTree.Nodes.AddRange(rootNodes.ToArray());
        }

        private async Task LoadImapFolderTree()
        {
            folderTree.Nodes.Clear();
            try
            {
                using var client = new ImapClient();
                await client.ConnectAsync(currentAccount.ImapServer, currentAccount.ImapPort, currentAccount.ImapUseSsl);
                await client.AuthenticateAsync(currentAccount.Email, DecryptString(currentAccount.EncryptedPassword));
                var personal = client.GetFolder(client.PersonalNamespaces[0]);
                var folders = personal.GetSubfolders(false);
                foreach (var f in folders)
                {
                    TreeNode node = new TreeNode(f.Name) { Tag = f.FullName };
                    folderTree.Nodes.Add(node);
                }
                await client.DisconnectAsync(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载 IMAP 文件夹失败: {ex.Message}");
            }
        }

        // ---------- 加载邮件 ----------
        private async void FolderTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (e.Node == null) return;
            currentFolderId = e.Node.Tag.ToString();
            await LoadMessagesForFolder(currentFolderId);
        }

        private async Task LoadMessagesForFolder(string folderId)
        {
            if (currentAccount == null) return;
            btnRefreshInbox.Enabled = false;
            lblInboxStatus.Text = "正在加载邮件...";
            inboxGrid.Rows.Clear();
            currentMessages.Clear();
            try
            {
                if (useGraph)
                {
                    await LoadGraphMessages(folderId);
                }
                else
                {
                    await LoadImapMessages(folderId);
                }
            }
            catch (Exception ex)
            {
                lblInboxStatus.Text = $"加载失败: {ex.Message}";
                MessageBox.Show($"加载邮件失败: {ex.Message}");
            }
            finally
            {
                btnRefreshInbox.Enabled = true;
            }
        }

        private async Task LoadGraphMessages(string folderId)
        {
            if (!await RefreshGraphTokenIfNeededAsync()) throw new Exception("刷新 token 失败");
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
            string url = $"https://graph.microsoft.com/v1.0/me/mailfolders/{folderId}/messages?$select=from,subject,receivedDateTime,id&$orderby=receivedDateTime desc&$top=50";
            var response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var messages = doc.RootElement.GetProperty("value");
            foreach (var msg in messages.EnumerateArray())
            {
                var fromObj = msg.GetProperty("from").GetProperty("emailAddress");
                string fromName = fromObj.TryGetProperty("name", out var name) ? name.GetString() : "";
                string fromEmail = fromObj.GetProperty("address").GetString();
                string fromDisplay = string.IsNullOrEmpty(fromName) ? fromEmail : $"{fromName} <{fromEmail}>";
                string subject = msg.GetProperty("subject").GetString() ?? "(无主题)";
                string received = msg.GetProperty("receivedDateTime").GetDateTime().ToString("yyyy-MM-dd HH:mm");
                string id = msg.GetProperty("id").GetString();
                inboxGrid.Rows.Add(fromDisplay, subject, received);
                currentMessages.Add(new { Id = id, From = fromDisplay, Subject = subject });
            }
            string folderName = folderIdNameMap.ContainsKey(folderId) ? folderIdNameMap[folderId] : folderId;
            lblInboxStatus.Text = $"文件夹 {folderName} 共 {messages.GetArrayLength()} 封邮件";
            if (notificationsEnabled && messages.GetArrayLength() > 0)
                ShowNotification("新邮件提醒", $"您有 {messages.GetArrayLength()} 封新邮件");
        }

        private async Task LoadImapMessages(string folderName)
        {
            using var client = new ImapClient();
            await client.ConnectAsync(currentAccount.ImapServer, currentAccount.ImapPort, currentAccount.ImapUseSsl);
            await client.AuthenticateAsync(currentAccount.Email, DecryptString(currentAccount.EncryptedPassword));
            var folder = client.GetFolder(folderName);
            await folder.OpenAsync(FolderAccess.ReadOnly);
            var uids = await folder.SearchAsync(SearchQuery.All);
            int count = Math.Min(50, uids.Count);
            for (int i = uids.Count - count; i < uids.Count; i++)
            {
                var uid = uids[i];
                var message = await folder.GetMessageAsync(uid);
                string fromDisplay = message.From.ToString();
                string subject = message.Subject ?? "(无主题)";
                string received = message.Date.LocalDateTime.ToString("yyyy-MM-dd HH:mm");
                string id = uid.ToString();
                inboxGrid.Rows.Add(fromDisplay, subject, received);
                currentMessages.Add(new { Id = id, From = fromDisplay, Subject = subject, RawMessage = message });
            }
            lblInboxStatus.Text = $"文件夹 {folderName} 共 {currentMessages.Count} 封邮件 (IMAP)";
            if (notificationsEnabled && currentMessages.Count > 0)
                ShowNotification("新邮件提醒", $"您有 {currentMessages.Count} 封新邮件");
            await client.DisconnectAsync(true);
        }

        // ---------- 双击显示邮件内容 ----------
        private async void InboxGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= currentMessages.Count) return;
            dynamic msg = currentMessages[e.RowIndex];
            string htmlContent = "";

            try
            {
                if (useGraph)
                {
                    string msgId = msg.Id;
                    if (!await RefreshGraphTokenIfNeededAsync()) throw new Exception("刷新 token 失败");
                    using var client = new HttpClient();
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
                    var response = await client.GetAsync($"https://graph.microsoft.com/v1.0/me/messages/{msgId}?$select=body,bodyPreview");
                    response.EnsureSuccessStatusCode();
                    var json = await response.Content.ReadAsStringAsync();
                    using var doc = JsonDocument.Parse(json);
                    string body = doc.RootElement.GetProperty("body").GetProperty("content").GetString();
                    htmlContent = $"<html><body style='font-family:Segoe UI; padding:10px;'>{body}</body></html>";
                }
                else
                {
                    var mimeMsg = msg.RawMessage as MimeMessage;
                    if (mimeMsg != null)
                    {
                        var body = mimeMsg.HtmlBody;
                        if (string.IsNullOrEmpty(body))
                            body = mimeMsg.TextBody;
                        if (string.IsNullOrEmpty(body))
                            body = "无法显示邮件正文";
                        htmlContent = $"<html><body style='font-family:Segoe UI; padding:10px;'>{body}</body></html>";
                    }
                }

                previewPanel.Visible = true;
                btnClosePreview.Visible = true;
                await EnsureWebView2();
                webView.CoreWebView2.NavigateToString(htmlContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载邮件内容失败: {ex.Message}");
                previewPanel.Visible = false;
                btnClosePreview.Visible = false;
            }
        }

        private async Task EnsureWebView2()
        {
            if (webView.CoreWebView2 != null) return;
            try
            {
                await webView.EnsureCoreWebView2Async();
            }
            catch (Exception ex)
            {
                if (!webView2Prompted)
                {
                    var result = MessageBox.Show($"未检测到 Microsoft Edge WebView2 运行时。\n是否现在下载安装？\n\n{ex.Message}", "缺少组件", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("https://go.microsoft.com/fwlink/p/?LinkId=2124703") { UseShellExecute = true });
                    webView2Prompted = true;
                }
                webView.CoreWebView2?.NavigateToString("<html><body>无法加载邮件内容，请安装 WebView2 后重试。</body></html>");
            }
        }

        // ---------- 刷新、删除、移动 ----------
        private async void BtnRefreshInbox_Click(object sender, EventArgs e)
        {
            if (currentAccount == null) { MessageBox.Show("请先选择账户"); return; }
            await LoadMessagesForFolder(currentFolderId);
        }

        private async void BtnDeleteSelected_Click(object sender, EventArgs e)
        {
            if (!useGraph) { MessageBox.Show("当前账户不支持删除操作（仅 Graph 账户）"); return; }
            if (inboxGrid.SelectedRows.Count == 0) return;
            int index = inboxGrid.SelectedRows[0].Index;
            string msgId = (currentMessages[index] as dynamic).Id;
            if (MessageBox.Show("确定删除这封邮件？", "确认", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                try
                {
                    if (!await RefreshGraphTokenIfNeededAsync()) throw new Exception("刷新 token 失败");
                    using var client = new HttpClient();
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
                    var response = await client.DeleteAsync($"https://graph.microsoft.com/v1.0/me/messages/{msgId}");
                    response.EnsureSuccessStatusCode();
                    await LoadMessagesForFolder(currentFolderId);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"删除失败: {ex.Message}");
                }
            }
        }

        private async void BtnMoveSelected_Click(object sender, EventArgs e)
        {
            if (!useGraph) { MessageBox.Show("当前账户不支持移动操作（仅 Graph 账户）"); return; }
            if (inboxGrid.SelectedRows.Count == 0) return;
            int index = inboxGrid.SelectedRows[0].Index;
            string msgId = (currentMessages[index] as dynamic).Id;
            string destFolder = moveToFolderCombo.SelectedItem.ToString();
            string folderId = destFolder == "垃圾邮件" ? "junkemail" : "deleteditems";
            try
            {
                if (!await RefreshGraphTokenIfNeededAsync()) throw new Exception("刷新 token 失败");
                using var client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
                var content = new StringContent($"{{\"destinationId\":\"{folderId}\"}}", Encoding.UTF8, "application/json");
                var response = await client.PostAsync($"https://graph.microsoft.com/v1.0/me/messages/{msgId}/move", content);
                response.EnsureSuccessStatusCode();
                await LoadMessagesForFolder(currentFolderId);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"移动失败: {ex.Message}");
            }
        }

        // ---------- 发送邮件 ----------
        private async void BtnSend_Click(object sender, EventArgs e)
        {
            if (currentAccount == null) { MessageBox.Show("请先选择账户"); return; }
            if (string.IsNullOrWhiteSpace(txtTo.Text) || !txtTo.Text.Contains("@")) { MessageBox.Show("收件人无效"); return; }
            if (string.IsNullOrWhiteSpace(txtSubject.Text)) { MessageBox.Show("请填写主题"); return; }

            btnSend.Enabled = false;
            try
            {
                if (useGraph)
                {
                    if (!await RefreshGraphTokenIfNeededAsync()) throw new Exception("Token 刷新失败，请重新登录");
                    await SendViaGraph(txtTo.Text.Trim(), txtSubject.Text.Trim(), txtBody.Text.Trim());
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(currentAccount.SmtpServer) || string.IsNullOrWhiteSpace(currentAccount.EncryptedPassword))
                        throw new Exception("账户 SMTP 配置不完整，请重新添加账户");
                    SendViaSmtp(currentAccount, txtTo.Text.Trim(), txtSubject.Text.Trim(), txtBody.Text.Trim());
                }
                MessageBox.Show("发送成功");
                txtTo.Clear();
                txtSubject.Clear();
                txtBody.Clear();
            }
            catch (Exception ex) { MessageBox.Show($"发送失败: {ex.Message}"); }
            finally { btnSend.Enabled = true; }
        }

        private async Task SendViaGraph(string to, string subject, string body)
        {
            var message = new
            {
                message = new
                {
                    subject = subject,
                    body = new { contentType = "Text", content = body },
                    toRecipients = new[] { new { emailAddress = new { address = to } } }
                },
                saveToSentItems = "true"
            };
            var json = JsonSerializer.Serialize(message);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", graphAccessToken);
            var response = await httpClient.PostAsync("https://graph.microsoft.com/v1.0/me/sendMail", content);
            if (!response.IsSuccessStatusCode)
                throw new Exception(await response.Content.ReadAsStringAsync());
        }

        private void SendViaSmtp(EmailAccount account, string to, string subject, string body)
        {
            using var smtp = new SmtpClient(account.SmtpServer, account.SmtpPort);
            smtp.EnableSsl = account.SmtpUseTls;
            smtp.Credentials = new NetworkCredential(account.Email, DecryptString(account.EncryptedPassword));
            using var mail = new MailMessage(account.Email, to, subject, body);
            smtp.Send(mail);
        }

        // ---------- Graph 辅助 ----------
        private async Task<bool> RefreshGraphTokenIfNeededAsync()
        {
            if (string.IsNullOrEmpty(graphRefreshToken)) return false;
            using var client = new HttpClient();
            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string,string>("client_id", ClientId),
                new KeyValuePair<string,string>("grant_type", "refresh_token"),
                new KeyValuePair<string,string>("refresh_token", graphRefreshToken),
                new KeyValuePair<string,string>("scope", string.Join(" ", GraphScopes))
            });
            var response = await client.PostAsync($"{Authority}/oauth2/v2.0/token", content);
            if (!response.IsSuccessStatusCode) return false;
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            graphAccessToken = root.GetProperty("access_token").GetString();
            graphRefreshToken = root.GetProperty("refresh_token").GetString();
            if (currentAccount != null && currentAccount.UseGraph)
            {
                currentAccount.GraphRefreshToken = graphRefreshToken;
                SaveAccounts();
            }
            return true;
        }

        private async Task<(bool success, string refreshToken, string email)> PerformGraphLogin()
        {
            try
            {
                var (deviceCode, userCode, verificationUri, interval) = await RequestDeviceCode();
                ShowDeviceCodeDialog(userCode, verificationUri);
                var token = await PollForToken(deviceCode, interval);
                string email = await FetchUserEmail(token.AccessToken);
                graphRefreshToken = token.RefreshToken;
                userEmail = email;
                return (true, token.RefreshToken, email);
            }
            catch { return (false, null, null); }
        }

        private async Task<(string deviceCode, string userCode, string verificationUri, int interval)> RequestDeviceCode()
        {
            using var client = new HttpClient();
            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string,string>("client_id", ClientId),
                new KeyValuePair<string,string>("scope", string.Join(" ", GraphScopes))
            });
            var response = await client.PostAsync($"{Authority}/oauth2/v2.0/devicecode", content);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            return (root.GetProperty("device_code").GetString(),
                    root.GetProperty("user_code").GetString(),
                    root.GetProperty("verification_uri").GetString(),
                    root.GetProperty("interval").GetInt32());
        }

        private void ShowDeviceCodeDialog(string userCode, string verificationUri)
        {
            var dialog = new Form
            {
                Text = "Outlook 授权",
                Size = new Size(450, 200),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            layout.RowCount = 3;
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var msgLabel = new Label
            {
                Text = $"请使用浏览器访问:\n\n{verificationUri}\n\n并输入以下设备码:\n\n{userCode}",
                Font = new Font("Consolas", 10),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            layout.Controls.Add(msgLabel, 0, 0);

            var openBtn = new Button { Text = "打开浏览器", Width = 120 };
            var copyBtn = new Button { Text = "复制设备码", Width = 120 };
            var doneBtn = new Button { Text = "我已完成授权", Width = 120 };
            var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight };
            btnPanel.Controls.Add(openBtn);
            btnPanel.Controls.Add(copyBtn);
            btnPanel.Controls.Add(doneBtn);
            layout.Controls.Add(btnPanel, 0, 1);

            openBtn.Click += (s, e) => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(verificationUri) { UseShellExecute = true });
            copyBtn.Click += (s, e) => { Clipboard.SetText(userCode); MessageBox.Show("设备码已复制"); };
            doneBtn.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(layout);
            dialog.ShowDialog();
        }

        private async Task<(string AccessToken, string RefreshToken)> PollForToken(string deviceCode, int interval)
        {
            using var client = new HttpClient();
            var content = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string,string>("client_id", ClientId),
                new KeyValuePair<string,string>("grant_type", "urn:ietf:params:oauth:grant-type:device_code"),
                new KeyValuePair<string,string>("device_code", deviceCode)
            });
            while (true)
            {
                await Task.Delay(interval * 1000);
                var response = await client.PostAsync($"{Authority}/oauth2/v2.0/token", content);
                var json = await response.Content.ReadAsStringAsync();
                if (response.IsSuccessStatusCode)
                {
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;
                    return (root.GetProperty("access_token").GetString(), root.GetProperty("refresh_token").GetString());
                }
                else
                {
                    using var doc = JsonDocument.Parse(json);
                    var error = doc.RootElement.GetProperty("error").GetString();
                    if (error == "authorization_pending") continue;
                    throw new Exception($"授权失败: {error}");
                }
            }
        }

        private async Task<string> FetchUserEmail(string accessToken)
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            var response = await client.GetAsync("https://graph.microsoft.com/v1.0/me");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;
            return root.TryGetProperty("mail", out var mail) ? mail.GetString() : root.GetProperty("userPrincipalName").GetString();
        }
    }
}