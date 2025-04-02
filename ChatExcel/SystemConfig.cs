using System.Drawing;
using System.IO;
using System.Reflection;

namespace ChatExcel
{
    internal class SystemConfig
    {
        /// <summary>
        /// WebSocketUrl
        /// </summary>
        public static string WebSocketUrl { get; set; } = "wss://ws-server.gavinhome.partykit.dev/party/chat-excel";

        /// <summary>
        /// WebSiteUrl
        /// </summary>
        public static string WebSiteUrl { get; set; } = "https://udify.app/chatbot/AVX31tbxs79E7br4";
       
        /// <summary>
        /// 标识
        /// </summary>
        public static string AppID { get; internal set; } = "ChatExcel";

        /// <summary>
        /// 名称
        /// </summary>
        public static string AppName { get; internal set; } = "智算大师";

        /// <summary>
        /// WebView缓存目录
        /// </summary>
        public static string WebViewApp { get; private set; } = "ChatExcel";

        /// <summary>
        /// 程序目录
        /// </summary>
        static string workDirectory;

        /// <summary>
        /// 获取程序目录
        /// </summary>
        public static string WorkDirectory
        {
            get
            {
                if (string.IsNullOrEmpty(workDirectory))
                {
                    workDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8));
                    if (workDirectory.ToLower().EndsWith("bin"))
                        workDirectory = Directory.GetParent(workDirectory).FullName;
                }
                return workDirectory;
            }
        }

        /// <summary>
        /// 获取日志目录全路径
        /// </summary>
        public static string LogDirectory
        {
            get
            {
                string path = Path.Combine(WorkDirectory, "Log");
                if (Directory.Exists(path) == false) 
                    Directory.CreateDirectory(path); 
                return path;
            }
        }

        /// <summary>
        /// 获取系统字体
        /// </summary>
        public static Font SystemFont { get; private set; } = new Font("宋体", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(134)));
        /// <summary>
        /// 获取系统前景色
        /// </summary>
        public static Color SystemColor { get; private set; } = SystemColors.ControlText;
        /// <summary>
        /// 获取系统背景色
        /// </summary>
        public static Color SystemBackColor { get; private set; } = SystemColors.Control;
    }
}
