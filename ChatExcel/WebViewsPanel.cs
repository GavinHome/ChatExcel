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
        private bool debugPanelVisible = false; // 保留变量跟踪调试面板状态

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
                Height = 200,  // 调试区域总高度
                Visible = false // 默认隐藏调试区域
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
            string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), SystemConfig.WebViewApp);
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await webView.EnsureCoreWebView2Async(env);
            webView.Source = new Uri(SystemConfig.WebSiteUrl);
            webViewPanel.Controls.Add(webView);
        }

        private TextBox txtVbaCode;
        private Button btnExecute;
        private Label lblDebugArea;
        private Button btnCloseDebug; // 添加关闭按钮

        // 添加输入框和执行按钮控件
        private void InitializeVbaControls()
        {
            // 创建一个面板来容纳标题和关闭按钮
            Panel titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
            };
            
            // 标识头部区域 - "调试区域"
            lblDebugArea = new Label
            {
                Dock = DockStyle.Left,
                Width = this.Width - 40, // 预留右侧空间给关闭按钮
                Text = "调试区域",
                TextAlign = ContentAlignment.MiddleLeft,
                Font = SystemConfig.SystemFont,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                Margin = new Padding(10, 0, 0, 0)
            };
            
            // 添加关闭按钮
            btnCloseDebug = new Button
            {
                Dock = DockStyle.Right,
                Text = "×",
                Width = 30,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(SystemConfig.SystemFont.FontFamily, 12, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnCloseDebug.FlatAppearance.BorderSize = 0;
            btnCloseDebug.Click += BtnCloseDebug_Click;
            
            // 将标题和关闭按钮添加到标题面板
            titlePanel.Controls.Add(btnCloseDebug);
            titlePanel.Controls.Add(lblDebugArea);
            
            // 执行按钮
            btnExecute = new Button
            {
                Dock = DockStyle.Bottom,
                Text = "执行 VBA 代码",
                Height = 40
            };
            btnExecute.Click += BtnExecute_Click;
            
            // 输入框 (VBA 代码)
            txtVbaCode = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
            };
            
            // 按照正确的顺序添加控件，确保布局正确
            debugPanel.Controls.Add(txtVbaCode);  // 先添加输入框
            debugPanel.Controls.Add(btnExecute);  // 再添加按钮
            debugPanel.Controls.Add(titlePanel);  // 添加标题面板，替代原来的标题标签
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

        // 关闭按钮点击事件处理
        private void BtnCloseDebug_Click(object sender, EventArgs e)
        {
            SetDebugPanelVisibility(false);
        }

        // 保留切换调试面板显示/隐藏的方法，以便将来可能需要
        private void ToggleDebugPanel()
        {
            debugPanelVisible = !debugPanelVisible;
            debugPanel.Visible = debugPanelVisible;
        }
        
        // 添加公共方法，允许外部代码切换调试面板
        public void ToggleDebugPanelVisibility()
        {
            ToggleDebugPanel();
        }
        
        // 添加公共方法，允许外部代码设置调试面板的可见性
        public void SetDebugPanelVisibility(bool visible)
        {
            try
            {
                // 确保在UI线程上执行
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => SetDebugPanelVisibility(visible)));
                    return;
                }
                
                // 设置可见性
                debugPanelVisible = visible;
                debugPanel.Visible = visible;
                
                // 调整WebView面板的大小
                AdjustWebViewSize();
                
                // 确保面板在前面并刷新
                debugPanel.BringToFront();
                this.Refresh();
                
                // 如果是显示面板，设置焦点
                if (visible)
                {
                    txtVbaCode.Focus();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SetDebugPanelVisibility 发生异常: {ex.Message}");
            }
        }
        
        // 添加方法来调整WebView的大小
        private void AdjustWebViewSize()
        {
            // 如果调试面板可见，调整WebView面板的大小
            if (debugPanelVisible)
            {
                // 设置WebView面板的Dock属性为Fill，但不包括底部的调试面板区域
                webViewPanel.Dock = DockStyle.None;
                webViewPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
                webViewPanel.Top = 0;
                webViewPanel.Left = 0;
                webViewPanel.Width = this.ClientSize.Width;
                webViewPanel.Height = this.ClientSize.Height - debugPanel.Height;
            }
            else
            {
                // 如果调试面板不可见，WebView面板占据整个区域
                webViewPanel.Dock = DockStyle.Fill;
            }
        }

        // 重写OnResize方法，确保在控件大小改变时调整WebView的大小
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            AdjustWebViewSize();
        }
        
        // 添加公共属性，允许外部代码获取调试面板的当前可见状态
        public bool IsDebugPanelVisible
        {
            get { return debugPanelVisible; }
        }
    }
}
