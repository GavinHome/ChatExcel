using Newtonsoft.Json;
using Serilog;
using System;
using System.Timers;
using WebSocketSharp;
using Timer = System.Timers.Timer;

namespace ChatExcel
{
    public class WebSocketClient
    {
        private WebSocket _webSocket;
        private string _serverUrl;
        private bool _isConnected;
        private Timer _reconnectTimer;
        private int _reconnectAttempts;
        private const int MaxReconnectAttempts = 5;
        private const int ReconnectInterval = 5000; // 5秒

        // 消息接收事件
        public event Action<string> OnMessageReceived;
        // 对象消息接收事件
        public event Action<CommandRequest> OnCommandRequestReceived;
        // 连接状态变化事件
        public event Action<bool> OnConnectionStateChanged;
        // 重连状态事件
        public event Action<int> OnReconnectAttempt;
        // 重连失败事件
        public event Action OnReconnectFailed;

        // 构造函数
        public WebSocketClient(string serverUrl)
        {
            Log.Information("开始创建WebSocket客户端");
            
            _serverUrl = serverUrl;
            _isConnected = false;
            _reconnectAttempts = 0;
            
            // 初始化重连定时器
            Log.Debug("初始化WebSocket重连定时器，间隔: {ReconnectInterval}毫秒，最大尝试次数: {MaxAttempts}", 
                ReconnectInterval, MaxReconnectAttempts);
                
            _reconnectTimer = new Timer(ReconnectInterval);
            _reconnectTimer.Elapsed += OnReconnectTimerElapsed;
            _reconnectTimer.AutoReset = false;
            _reconnectTimer.Enabled = false;
            
            Log.Information("WebSocket客户端已创建，服务器地址: {ServerUrl}", serverUrl);
        }

        private void OnReconnectTimerElapsed(object sender, ElapsedEventArgs e)
        {
            Log.Debug("重连定时器触发，当前尝试次数: {CurrentAttempt}/{MaxAttempts}", 
                _reconnectAttempts, MaxReconnectAttempts);
                
            if (_reconnectAttempts < MaxReconnectAttempts)
            {
                _reconnectAttempts++;
                Log.Information("尝试重新连接WebSocket服务器 (第{Attempt}/{MaxAttempts}次)", 
                    _reconnectAttempts, MaxReconnectAttempts);
                OnReconnectAttempt?.Invoke(_reconnectAttempts);
                Connect();
            }
            else
            {
                Log.Error("WebSocket重连失败，已达到最大重试次数 ({MaxAttempts})", MaxReconnectAttempts);
                _reconnectTimer.Enabled = false;
                OnReconnectFailed?.Invoke();
            }
        }

        // 手动重连
        public void ManualReconnect()
        {
            Log.Information("用户触发手动重连");
            _reconnectAttempts = 0;
            _reconnectTimer.Enabled = false;
            Log.Debug("重置重连尝试次数和定时器");
            Connect();
        }

        // 连接到WebSocket服务器
        public void Connect()
        {
            if (_isConnected)
            {
                Log.Warning("WebSocket已经处于连接状态，跳过连接操作");
                return;
            }

            try
            {
                Log.Information("开始连接到WebSocket服务器: {ServerUrl}", _serverUrl);
                _webSocket = new WebSocket(_serverUrl);
                Log.Debug("WebSocket实例已创建");

                // 注册WebSocket事件处理程序
                Log.Debug("注册WebSocket事件处理程序");
                
                _webSocket.OnOpen += (sender, e) =>
                {
                    _isConnected = true;
                    _reconnectAttempts = 0;
                    _reconnectTimer.Enabled = false;
                    OnConnectionStateChanged?.Invoke(true);
                    Log.Information("成功连接到WebSocket服务器: {ServerUrl}", _serverUrl);
                };

                _webSocket.OnClose += (sender, e) =>
                {
                    _isConnected = false;
                    OnConnectionStateChanged?.Invoke(false);
                    Log.Warning("与WebSocket服务器的连接已关闭，代码: {Code}, 原因: {Reason}", e.Code, e.Reason);

                    // 启动重连定时器
                    if (_reconnectAttempts < MaxReconnectAttempts)
                    {
                        Log.Debug("开始自动重连，当前尝试次数: {CurrentAttempt}/{MaxAttempts}", 
                            _reconnectAttempts, MaxReconnectAttempts);
                        _reconnectTimer.Enabled = true;
                    }
                    else
                    {
                        Log.Warning("已达到最大重连尝试次数，停止自动重连");
                        _reconnectTimer.Enabled = false;
                        OnReconnectFailed?.Invoke();
                    }
                };

                _webSocket.OnError += (sender, e) =>
                {
                    Log.Error(e.Exception, "WebSocket发生错误: {ErrorMessage}", e.Message);
                };

                _webSocket.OnMessage += (sender, e) =>
                {
                    try
                    {
                        // 处理接收到的消息
                        string message = e.Data;
                        // 如果消息过长，只记录前100个字符
                        string truncatedMessage = message.Length > 100 
                            ? message.Substring(0, 100) + "..." 
                            : message;
                        Log.Information("收到WebSocket消息: {TruncatedMessage}, 长度: {Length}字符", 
                            truncatedMessage, message.Length);
                        
                        try
                        {
                            // 尝试将消息解析为 CommandRequest 对象
                            var commandRequest = JsonConvert.DeserializeObject<CommandRequest>(message);
                            if (commandRequest != null)
                            {
                                // 触发对象消息接收事件
                                //Log.Information("解析为CommandRequest: EventName={EventName}, 参数长度={ParamLength}字符", 
                                //    commandRequest.EventName, commandRequest.EventParams?.Length ?? 0);
                                OnCommandRequestReceived?.Invoke(commandRequest);
                            }
                            else
                            {
                                // 如果不是 CommandRequest 对象，则触发普通消息事件
                                Log.Information("解析为空CommandRequest，触发普通消息事件");
                                OnMessageReceived?.Invoke(message);
                            }
                        }
                        catch (JsonException ex)
                        {
                            // 如果解析失败，则作为普通字符串消息处理
                            Log.Error(ex, "JSON解析失败，将作为普通消息处理: {TruncatedMessage}", truncatedMessage);
                            OnMessageReceived?.Invoke(message);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "处理WebSocket消息时发生错误");
                    }
                };

                // 连接到服务器
                Log.Debug("开始连接到WebSocket服务器");
                _webSocket.Connect();
                Log.Debug("WebSocket连接请求已发送");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "连接到WebSocket服务器时发生错误: {ErrorMessage}", ex.Message);
                _isConnected = false;
                
                // 启动重连定时器
                if (_reconnectAttempts < MaxReconnectAttempts)
                {
                    Log.Debug("连接失败，启动自动重连定时器");
                    _reconnectTimer.Enabled = true;
                }
            }
        }

        // 断开与WebSocket服务器的连接
        public void Disconnect()
        {
            if (!_isConnected || _webSocket == null)
            {
                Log.Warning("WebSocket未连接或实例为空，跳过断开连接操作");
                return;
            }

            try
            {
                Log.Information("开始断开WebSocket连接");
                _reconnectTimer.Enabled = false;
                _reconnectAttempts = 0;
                
                // 正常关闭WebSocket连接
                Log.Debug("发送WebSocket关闭请求，代码: 1000 (正常关闭)");
                _webSocket.Close(1000, "客户端主动断开连接");
                _isConnected = false;
                Log.Information("成功断开WebSocket连接");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "断开WebSocket连接时发生错误: {ErrorMessage}", ex.Message);
                // 确保连接状态更新，即使出现异常
                _isConnected = false;
            }
        }

        // 发送消息到WebSocket服务器
        public void SendMessage(string message)
        {
            if (!_isConnected || _webSocket == null)
            {
                Log.Error("无法发送消息：WebSocket未连接");
                return;
            }

            try
            {
                // 如果消息过长，只记录前100个字符
                string truncatedMessage = message.Length > 100 
                    ? message.Substring(0, 100) + "..." 
                    : message;
                Log.Information("准备发送消息: {TruncatedMessage}, 长度: {Length}字符", 
                    truncatedMessage, message.Length);
                    
                _webSocket.Send(message);
                Log.Debug("消息发送成功，长度: {Length}字符", message.Length);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "发送消息时发生错误: {ErrorMessage}, 消息长度: {Length}字符", 
                    ex.Message, message?.Length ?? 0);
            }
        }

        // 发送对象到WebSocket服务器（自动序列化为JSON）
        public void SendMessage<T>(T obj)
        {
            if (!_isConnected || _webSocket == null)
            {
                Log.Error("无法发送对象：WebSocket未连接");
                return;
            }

            try
            {
                Log.Debug("序列化对象为JSON，类型: {ObjectType}", typeof(T).Name);
                string jsonMessage = JsonConvert.SerializeObject(obj, Formatting.Indented, new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });
                Log.Information("准备发送对象: {@jsonMessage}", jsonMessage);
                _webSocket.Send(jsonMessage);
                Log.Debug("JSON对象发送成功，长度: {Length}字符", jsonMessage.Length);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "发送对象时发生错误: {ErrorMessage}, 对象类型: {ObjectType}", 
                    ex.Message, typeof(T).Name);
            }
        }

        // 检查是否已连接到WebSocket服务器
        public bool IsConnected
        {
            get 
            { 
                bool connected = _isConnected && _webSocket != null;
                Log.Debug("检查WebSocket连接状态: {IsConnected}", connected);
                return connected; 
            }
        }
    }

    public class CommandRequest
    {
        public string Message { get; set; }
    }
}