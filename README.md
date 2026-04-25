# Simple Email Client

一个轻量级的 Windows 邮件客户端，支持 Outlook (Graph API)、QQ邮箱、163邮箱、Gmail 等，支持 SMTP 发送和 IMAP 接收，并允许同时管理多个邮箱账户。

## 功能特性

- **多账户支持**：可添加多个邮箱账户（Graph API 或自定义 SMTP/IMAP）。
- **收件箱**：查看邮件列表、阅读邮件内容（内置 WebView2 渲染 HTML）。
- **发送邮件**：支持富文本编辑，发送后自动保存到已发送。
- **文件夹管理**：支持标准邮件文件夹（收件箱、已发送、垃圾邮件、已删除等）。
- **系统通知**：新邮件到达时弹出系统通知（Windows 10+）。
- **mailto 协议**：可注册为系统默认邮件客户端，点击网页上的邮件链接自动打开。
- **账户加密存储**：密码和 Token 使用 Windows DPAPI 加密，仅当前用户可解密。
- **轻量部署**：单个 exe 文件（需 .NET 5 运行时）或自包含发布（约 60MB）。

## 截图

<img width="1365" height="723" alt="image" src="https://github.com/user-attachments/assets/f966b24b-d84c-40de-bdaa-db309f8a9104" />

## 运行需求

- 需要安装.NET 5.0运行时：https://dotnet.microsoft.com/zh-cn/download/dotnet/thank-you/runtime-aspnetcore-5.0.17-windows-x64-installer
- 以及 Microsoft Edge Webview2: https://go.microsoft.com/fwlink/p/?LinkId=2124703

## 快速开始

1. 启动程序。
2. 点击 **“添加”** 账户：
   - 若使用 Outlook/Hotmail，选择 **Graph API**，点击 **“使用 Graph 登录”** 完成授权。
   - 若使用其他邮箱（如 QQ、163、Gmail），选择 **自行配置 SMTP/IMAP**，填写服务器地址、端口、授权码。
3. 选择添加的账户，点击 **“选择”**。
4. 在 **“收件箱”** 标签页，点击 **“刷新”** 加载邮件。
5. 在 **“发送邮件”** 标签页，填写收件人、主题、内容，点击 **“发送”**。

## 配置说明

- **设置**：支持注册 mailto 协议、自动加载收件箱、开启系统通知。
- **多文件夹**：在收件箱左侧树形视图可切换不同邮件文件夹。
- **删除/移动**：仅 Graph API 账户支持删除和移动邮件（IMAP 账户暂不支持）。

## 构建环境要求

- Visual Studio 2019 或更高版本（需支持 .NET 5）
- .NET 5 SDK
- NuGet 包：
  - `Microsoft.Web.WebView2`
  - `MailKit`
