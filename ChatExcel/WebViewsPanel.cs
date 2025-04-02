using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChatExcel
{
    public partial class WebViewsPanel : UserControl
    {
        private Panel webViewPanel;
        private Panel debugPanel;

        public WebViewsPanel()
        {
            InitializeComponent(); 
            
            // 创建主面板布局
            webViewPanel = new Panel
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(webViewPanel);

            debugPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 170  // 调试区域总高度
            };
            Controls.Add(debugPanel);

            InitializeWebView2();
            InitializeVbaControls();
        }

        private async void InitializeWebView2()
        {
            var webView = new Microsoft.Web.WebView2.WinForms.WebView2();
            webView.Dock = System.Windows.Forms.DockStyle.Fill;

            // 指定自定义缓存路径（例如用户 AppData 文件夹）
            string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ChatExcel");
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await webView.EnsureCoreWebView2Async(env);
            webView.Source = new Uri("https://udify.app/chatbot/AVX31tbxs79E7br4");
            webViewPanel.Controls.Add(webView);
        }

        private TextBox txtVbaCode;
        private Button btnExecute;
        private Label lblDebugArea;

        // 添加输入框和执行按钮控件
        private void InitializeVbaControls()
        {
            // 标识头部区域 - "调试区域"
            lblDebugArea = new Label
            {
                Dock = DockStyle.Top,
                Height = 30,
                Text = "调试区域",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                Padding = new Padding(10, 0, 0, 0),
                Margin = new Padding(10, 15, 10, 10)
            };
            debugPanel.Controls.Add(lblDebugArea);

            // 执行按钮
            btnExecute = new Button
            {
                Dock = DockStyle.Bottom,
                Text = "执行 VBA 代码",
                Height = 40
            };
            btnExecute.Click += BtnExecute_Click;
            debugPanel.Controls.Add(btnExecute);

            // 输入框 (VBA 代码)
            txtVbaCode = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
            };
            debugPanel.Controls.Add(txtVbaCode);
        }

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            string vbaCode = txtVbaCode.Text;

            if (string.IsNullOrWhiteSpace(vbaCode))
            {
                MessageBox.Show("请输入 VBA 代码！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 调用 ThisAddIn 中的 RunVba 方法执行代码
            var addIn = Globals.ThisAddIn; // 获取当前的 VSTO Add-in 实例
            addIn.RunVba(vbaCode); // 执行 VBA 代码
        }
    }
}
