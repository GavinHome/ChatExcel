using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Vbe.Interop;
using System.Windows.Forms;
using Microsoft.Win32;
using Serilog;
using System.IO;

namespace ChatExcel
{
    public partial class ThisAddIn
    {
        private static WebSocketClient _webSocketClient = new WebSocketClient("wss://ws-server.gavinhome.partykit.dev/party/chat-excel");

        private Microsoft.Office.Tools.CustomTaskPane customTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LoadConfig();

            // 自动打开工作簿
            CreateAndOpenFile();

            // 加载面板
            CreateCustomTaskPane();

            WebSocketInit();
        }

        private void LoadConfig()
        {
            Log.Information("开始加载配置");
            try
            {
                // 配置 Serilog
                var logFilePath = Path.Combine(SystemConfig.LogDirectory, "Serilog.log");
                Log.Information("日志文件路径: {LogFilePath}", logFilePath);

                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()           // 设置最低日志级别
                    .Enrich.FromLogContext()     // 启用日志上下文
                    .WriteTo.File(logFilePath,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [{SourceContext}] [{EventId}]  {Message:lj}{NewLine}{Exception}",
                        rollingInterval: RollingInterval.Day,
                        rollOnFileSizeLimit: true,
                        fileSizeLimitBytes: 10000000)
                    .CreateLogger();

                // 记录程序启动日志
                Log.Information("CADAgent应用程序启动，ID: {CADAgentID}", SystemConfig.AppID);
                Log.Information("配置初始化完成");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载配置时发生异常: {ex.Message}");
            }
        }

        private void CreateAndOpenFile()
        {
            try
            {
                string workbookPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "工作簿1.xlsm");
                if (System.IO.File.Exists(workbookPath))
                {
                    Application.Workbooks.Open(workbookPath);
                }
                else
                {
                    // 如果文件不存在，创建一个新的工作簿并保存
                    var newWorkbook = Application.Workbooks.Add();
                    newWorkbook.SaveAs(workbookPath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开工作簿时出错: " + ex.Message);
            }
        }

        private void WebSocketInit()
        {
            Log.Information("开始初始化WebSocket");
            try
            {
                // 注册消息处理程序
                _webSocketClient.OnCommandRequestReceived += HandleCommandRequest;
                _webSocketClient.OnMessageReceived += HandleMessageReceived;
                _webSocketClient.OnConnectionStateChanged += HandleConnectionStateChanged;
                _webSocketClient.OnReconnectAttempt += HandleReconnectAttempt;
                _webSocketClient.OnReconnectFailed += HandleReconnectFailed;
                Log.Information("WebSocket事件处理程序注册完成");

                // 连接
                _webSocketClient.Connect();
                Log.Debug("WebSocket连接请求已发送");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "WebSocket客户端初始化失败: {ErrorMessage}, 请检查网络连接或服务器状态", ex.Message);
            }

            Log.Information("WebSocket初始化完成");
        }

        #region CustomTaskPane 

        private void CreateCustomTaskPane()
        {
            var webViewPanel = new WebViewsPanel();
            customTaskPane = this.CustomTaskPanes.Add(webViewPanel, "ChatExcel");
            customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

            // 设置初始宽度为Excel窗口宽度的40%
            SetTaskPaneWidth();

            // 监听Excel窗口大小变化
            Application.WindowResize += Application_WindowResize;

            customTaskPane.Visible = true;
        }

        private void Application_WindowResize(Excel.Workbook Wb, Excel.Window Wn)
        {
            SetTaskPaneWidth();
        }

        private void SetTaskPaneWidth()
        {
            if (Application.ActiveWindow != null)
            {
                // 获取Excel窗口宽度并计算40%的宽度
                double windowWidth = Application.ActiveWindow.Width;
                int calculatedWidth = (int)(windowWidth * 0.4);
                customTaskPane.Width = Math.Max(calculatedWidth, 400);
            }
            else
            {
                customTaskPane.Width = 400;
            }
        }

        #endregion

        #region VBA

        /// <summary>
        /// 格式化VBA代码，判断是否需要添加Sub和End Sub
        /// </summary>
        /// <param name="vbaCode">原始VBA代码</param>
        /// <returns>格式化后的VBA代码</returns>
        private string FormatVbaCode(string vbaCode)
        {
            // 移除代码块标记
            vbaCode = vbaCode.Replace("```vba", "").Replace("```", "").Trim();

            // 检查代码是否已经包含了Public Sub GeneratedMacro()和End Sub
            if (vbaCode.Contains("Public Sub GeneratedMacro()") && vbaCode.Contains("End Sub"))
            {
                // 如果已经包含了完整的结构，直接返回
                return vbaCode;
            }
            else
            {
                // 否则，添加必要的结构
                return $"Public Sub GeneratedMacro()\n{vbaCode}\nEnd Sub";
            }
        }

        public void RunVba(string vbaCode)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;

            if (wb == null)
            {
                MessageBox.Show("请先打开一个 Excel 工作簿");
                return;
            }

            try
            {
                // 检查文件格式
                string fileExtension = System.IO.Path.GetExtension(wb.FullName).ToLower();
                if (fileExtension != ".xlsm" && fileExtension != ".xls")
                {
                    DialogResult result = MessageBox.Show(
                        "当前工作簿不支持宏。需要将工作簿保存为启用宏的格式(.xlsm)。是否继续？\n\n" +
                        "注意：保存为.xlsm格式后，请重新打开工作簿。",
                        "需要启用宏支持",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            string newPath = System.IO.Path.ChangeExtension(wb.FullName, ".xlsm");
                            wb.SaveAs(newPath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                            MessageBox.Show("已保存为启用宏的格式(.xlsm)。请关闭并重新打开工作簿，然后再次尝试操作。");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("保存工作簿时出错: " + ex.Message);
                        }
                    }
                    return;
                }

                // 检查VBA访问权限
                try
                {
                    var test = wb.VBProject;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    EnableVBAAccess();
                    MessageBox.Show(
                        "请按照以下步骤操作：\n\n" +
                        "1. 关闭所有Excel窗口\n" +
                        "2. 打开Excel\n" +
                        "3. 点击'文件' -> '选项' -> '信任中心' -> '信任中心设置' -> '宏设置'\n" +
                        "4. 勾选'信任对VBA项目对象模型的访问'\n" +
                        "5. 点击确定并重启Excel\n\n" +
                        "完成以上步骤后，请重新运行此操作。",
                        "需要启用VBA访问权限",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                // 插入并运行VBA代码
                try
                {
                    var vbaModule = wb.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                    // 使用FormatVbaCode方法格式化VBA代码
                    string formattedVbaCode = FormatVbaCode(vbaCode);

                    // 先尝试删除可能存在的同名模块
                    try
                    {
                        foreach (VBComponent comp in wb.VBProject.VBComponents)
                        {
                            if (comp.Name == "Module1")
                            {
                                wb.VBProject.VBComponents.Remove(comp);
                                break;
                            }
                        }
                    }
                    catch { }

                    // 添加代码到模块
                    vbaModule.CodeModule.AddFromString(formattedVbaCode);

                    // 保存工作簿以确保VBA代码被保存
                    wb.Save();

                    // 运行宏
                    try
                    {
                        app.Run("GeneratedMacro");
                    }
                    catch (Exception runEx)
                    {
                        MessageBox.Show($"运行宏时出错: {runEx.Message}\n\n请检查Excel的宏设置是否已启用。");
                        return;
                    }

                    // 清理：删除临时模块
                    wb.VBProject.VBComponents.Remove(vbaModule);
                    wb.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"执行 VBA 代码失败: {ex.Message}\n\n详细信息：\n{ex.StackTrace}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("执行操作时出错: " + ex.Message);
            }
        }

        public void EnableVBAAccess()
        {
            try
            {
                // 获取 Excel 版本号
                Excel.Application excelApp = new Excel.Application();
                string version = excelApp.Version;
                excelApp.Quit();

                string regPath = $@"Software\Microsoft\Office\{version}\Excel\Security";

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(regPath, true))
                {
                    if (key != null)
                    {
                        key.SetValue("AccessVBOM", 1, RegistryValueKind.DWord); // 启用 VBA 访问
                        key.SetValue("VBAWarnings", 1, RegistryValueKind.DWord); // 启用宏
                    }
                    else
                    {
                        MessageBox.Show("无法找到 Excel 安装的注册表路径，可能是未正确安装 Office。");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("修改 VBA 访问权限失败，请手动启用。\n" + ex.Message);
            }
        }

        #endregion

        #region WebSocket

        private void HandleMessageReceived(string obj)
        {
            string vbaCode = obj;

            if (string.IsNullOrWhiteSpace(vbaCode))
            {
                MessageBox.Show("请输入 VBA 代码！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (vbaCode.Contains("hello"))
            {
                MessageBox.Show($"{vbaCode}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            RunVba(vbaCode); // 执行 VBA 代码
        }

        // 处理从WebSocket接收到的命令请求
        private void HandleCommandRequest(CommandRequest request)
        {
            Log.Information("HandleCommandRequest");
        }

        // 处理WebSocket连接状态变化
        private void HandleConnectionStateChanged(bool isConnected)
        {
            if (isConnected)
            {
                Log.Information("WebSocket连接成功");
            }
            else
            {
                Log.Warning("WebSocket连接断开，请检查网络连接或服务器状态");
            }
        }

        // 处理WebSocket重连尝试
        private void HandleReconnectAttempt(int attempt)
        {
            Log.Information("WebSocket正在进行第{Attempt}次重连尝试", attempt);
        }

        // 处理WebSocket重连失败
        private void HandleReconnectFailed()
        {
            Log.Warning("WebSocket自动重连失败，等待用户手动重连");
            try
            {
                MessageBox.Show(
                    "WebSocket连接已断开，自动重连失败。\n" +
                    "请检查网络连接和服务器状态，\n" +
                    "如果服务器已恢复正常，可以点击确定进行手动重连。",
                    "连接断开",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                Log.Debug("已显示WebSocket重连失败对话框");

                try
                {
                    Log.Information("尝试手动重连WebSocket");
                    _webSocketClient.ManualReconnect();
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "手动重连失败: {ErrorMessage}", ex.Message);
                    MessageBox.Show(
                        "手动重连失败，请检查网络连接和服务器状态。",
                        "重连失败",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    Log.Debug("已显示手动重连失败对话框");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "显示重连对话框或执行手动重连过程中发生错误: {ErrorMessage}", ex.Message);
            }
        }

        #endregion


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                Log.Information("开始执行CADCommand.Terminate");
                _webSocketClient.Disconnect();
                _webSocketClient = null;

                Log.Information("CADAgent应用程序终止");
                Log.CloseAndFlush();
            }
            catch (Exception ex)
            {
                // 在终止方法中异常可能无法正常记录
                System.Diagnostics.Debug.WriteLine($"终止程序时发生异常: {ex.Message}");
                Log.Error(ex, "CADCommand.Terminate执行失败: {ErrorMessage}", ex.Message);
                Log.CloseAndFlush();
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
