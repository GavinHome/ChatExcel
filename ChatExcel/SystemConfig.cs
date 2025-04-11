using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Serilog;

namespace ChatExcel
{
    internal class SystemConfig
    {
        /// <summary>
        /// WebSocketUrl
        /// </summary>
        public static string WebSocketUrl { get; set; } = GetConfigValue("WebSocketUrl", "");

        /// <summary>
        /// WebSiteUrl
        /// </summary>
        public static string WebSiteUrl { get; set; } = GetConfigValue("WebSiteUrl", "");

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

        /// <summary>
        /// 从配置文件中读取值，如果不存在则使用默认值
        /// </summary>
        /// <param name="key">配置键名</param>
        /// <param name="defaultValue">默认值</param>
        /// <returns>配置值或默认值</returns>
        private static string GetConfigValue(string key, string defaultValue)
        {
            try
            {
                string configPath = Path.Combine(WorkDirectory, "appsettings.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    JObject config = JObject.Parse(json);
                    JToken value = config[key];
                    if (value != null)
                    {
                        string configValue = value.ToString();

                        Log.Debug("从配置文件读取 {Key} = {Value}", key, configValue);
                        return configValue;
                    }
                }
            }
            catch (Exception ex)
            {
                // 读取配置失败时记录错误并使用默认值
                System.Diagnostics.Debug.WriteLine($"读取配置文件时出错: {ex.Message}");
                Log.Error(ex, "读取配置文件 {Key} 时出错: {ErrorMessage}", key, ex.Message);
            }

            return defaultValue;
        }
    }
}
