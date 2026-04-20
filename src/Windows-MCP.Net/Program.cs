using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using WindowsMCP.Net.Services;
using WindowsMCP.Net.Tools;
using WindowsMCP.Net.Tools.OCR;
using Interface;

// 配置全局日志记录器
Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .MinimumLevel.Override("Microsoft", LogEventLevel.Information)
    .MinimumLevel.Override("Tools.Desktop", LogEventLevel.Debug) // 为Tools.Desktop命名空间启用Debug级别
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File(
        "logs/winmcplog-.txt",
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 31,
        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"
    )
    .CreateLogger();

try
{
    // 检查命令行参数以确定运行模式
    var useRemote = args.Any(arg => arg.Equals("--remote", StringComparison.OrdinalIgnoreCase) ||
                                    arg.Equals("-r", StringComparison.OrdinalIgnoreCase) ||
                                    arg.Equals("--mode=REMOTE", StringComparison.OrdinalIgnoreCase));
    
    var useTestMode = args.Any(arg => arg.Equals("--test", StringComparison.OrdinalIgnoreCase) ||
                                      arg.Equals("--test-ocr", StringComparison.OrdinalIgnoreCase));

    if (useTestMode)
    {
        // 测试模式 - 运行OCR服务功能测试
        Log.Information("Starting Windows MCP Server in TEST mode (OCR service validation)");
        await RunOcrTest();
        Log.Information("Test completed successfully");
        return;
    }
    else if (useRemote)
    {
        // 远程模式 - 使用 Kestrel 服务器
        Log.Information("Starting Windows MCP Server in REMOTE mode (HTTP transport)");

        var builder = WebApplication.CreateBuilder(args);

        // 解析命令行参数
        var port = 8888;
        var host = "localhost";

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].Equals("--port", StringComparison.OrdinalIgnoreCase) || args[i].Equals("-p", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 < args.Length && int.TryParse(args[i + 1], out int parsedPort))
                {
                    port = parsedPort;
                }
            }
            else if (args[i].Equals("--host", StringComparison.OrdinalIgnoreCase) || args[i].Equals("-h", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 < args.Length)
                {
                    host = args[i + 1];
                }
            }
        }

        // 配置 Kestrel 监听地址
        builder.WebHost.ConfigureKestrel(serverOptions =>
        {
            if (host == "0.0.0.0")
            {
                serverOptions.ListenAnyIP(port);
            }
            else
            {
                serverOptions.ListenLocalhost(port);
            }
        });

        // 注册服务
        builder.Services
            .AddSingleton<IDesktopService, DesktopService>()
            .AddSingleton<IFileSystemService, FileSystemService>()
            .AddSingleton<IOcrService, OcrService>()
            .AddSingleton<ISystemControlService, SystemControlService>()
            .AddSingleton<UIAutomationService>()
            .AddSingleton<GetTextCoordinatesTool>()
            .AddSingleton<Tools.Desktop.ScreenshotTool>()
            .AddSingleton<Tools.Desktop.SwitchTool>()
            .AddSingleton<Tools.Desktop.GetWindowInfoTool>()
            .AddSingleton<Tools.Desktop.ResizeTool>()
            .AddSingleton<Tools.Desktop.UIElementTool>();

        // 构建应用
        var app = builder.Build();

        // 添加中间件处理 MCP 请求
        app.Use(async (context, next) =>
        {
            // 处理根路径的POST请求（普通JSON-RPC）
            if (context.Request.Path == "/" && context.Request.Method == "POST")
            {
                // 设置正确的 MCP 响应头
                context.Response.Headers.Append("Content-Type", "application/json");
                context.Response.Headers.Append("Cache-Control", "no-cache");
                context.Response.Headers.Append("Connection", "keep-alive");
                
                await ProcessRequestAsync(context, app.Services);
                // 不调用 next()，直接处理请求
                return;
            }
            
            // 处理工具发现请求
            if (context.Request.Path == "/tools" && context.Request.Method == "GET")
            {
                var tools = DiscoverAllTools();
                var response = JsonSerializer.Serialize(new
                {
                    tools = tools
                });
                
                context.Response.ContentType = "application/json; charset=utf-8";
                await context.Response.WriteAsync(response);
                return;
            }
            
            // 对于其他请求，继续管道处理
            await next();
        });

        // 启动服务
        Log.Information("Kestrel server started on {Url}", $"http://{host}:{port}");
        app.Run();
    }
    else
    {
        // 本地模式 - 使用 stdio 传输
        Log.Information("Starting Windows MCP Server in LOCAL mode (stdio transport)");

        var builder = Host.CreateApplicationBuilder(args);

        // 配置日志输出到 stderr
        builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);

        // 注册服务
        builder.Services
            .AddSingleton<IDesktopService, DesktopService>()
            .AddSingleton<IFileSystemService, FileSystemService>()
            .AddSingleton<IOcrService, OcrService>()
            .AddSingleton<ISystemControlService, SystemControlService>()
            .AddMcpServer()
            .WithStdioServerTransport()
            .WithToolsFromAssembly(Assembly.GetExecutingAssembly());

        await builder.Build().RunAsync();
    }
}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
}
finally
{
    Log.CloseAndFlush();
}

static async Task ProcessRequestAsync(HttpContext context, IServiceProvider serviceProvider)
{
    try
    {
        // 检查请求方法是否为POST
        if (context.Request.Method != "POST")
        {
            context.Response.StatusCode = 405; // Method Not Allowed
            await context.Response.WriteAsync("Method Not Allowed");
            return;
        }

        // 检查Content-Type是否为application/json
        if (!context.Request.Headers.ContainsKey("Content-Type") || 
            !context.Request.Headers["Content-Type"].ToString().Contains("application/json"))
        {
            context.Response.StatusCode = 415; // Unsupported Media Type
            await context.Response.WriteAsync("Unsupported Media Type");
            return;
        }

        // 读取请求体，使用UTF-8编码并处理BOM
        using var reader = new StreamReader(context.Request.Body, Encoding.UTF8, true, 1024, true);
        var body = await reader.ReadToEndAsync();
        Log.Information("Received request body: {Body}", body);  // 添加这行来记录请求体内容

        // 解析JSON-RPC请求
        var requestBody = body;  // 使用已读取的body变量

        Log.Debug("Received request: {Method} {Path}", context.Request.Method, context.Request.Path);
        Log.Debug("Request body: {RequestBody}", requestBody);

        // 处理 MCP 请求
        var response = await ProcessMcpRequest(requestBody, serviceProvider);

        // 发送响应
        context.Response.ContentType = "application/json; charset=utf-8";
        var responseBytes = Encoding.UTF8.GetBytes(response);
        await context.Response.Body.WriteAsync(responseBytes);
    }
    catch (Exception ex)
    {
        Log.Error(ex, "Error processing request");
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Internal Server Error");
    }
}

static async Task<string> ProcessMcpRequest(string requestBody, IServiceProvider serviceProvider)
{
    try
    {
        using var doc = JsonDocument.Parse(requestBody);
        var root = doc.RootElement;

        // 检查JSON-RPC版本
        if (!root.TryGetProperty("jsonrpc", out var jsonrpcElement) || 
            jsonrpcElement.GetString() != "2.0")
        {
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                error = new { code = -32600, message = "Invalid Request: jsonrpc version must be 2.0" },
                id = (long?)null
            });
        }

        // 检查method字段
        if (!root.TryGetProperty("method", out var methodElement))
        {
            var errorId = root.TryGetProperty("id", out var errorIdElement) ? 
                (errorIdElement.ValueKind == JsonValueKind.Number ? (object)errorIdElement.GetInt64() : 
                 errorIdElement.ValueKind == JsonValueKind.String ? (object)errorIdElement.GetString() : null) : null;
                 
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                error = new { code = -32600, message = "Invalid Request: method is required" },
                id = errorId
            });
        }

        var method = methodElement.GetString();
        var id = root.TryGetProperty("id", out var idElement) ? 
            (idElement.ValueKind == JsonValueKind.Number ? (object)idElement.GetInt64() : 
             idElement.ValueKind == JsonValueKind.String ? (object)idElement.GetString() : null) : null;

        // 处理 MCP 初始化请求
        if (method == "initialize")
        {
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                result = new
                {
                    protocolVersion = "2024-11-05",
                    capabilities = new
                    {
                        logging = new { },
                        progress = new { },
                        tools = new { },
                        prompts = new { }
                    },
                    serverInfo = new
                    {
                        name = "Windows-MCP.Net",
                        version = "1.0.0"
                    }
                },
                id = id
            });
        }
        // 处理工具列表请求
        else if (method == "tools/list")
        {
            // 自动发现所有可用的工具
            var tools = DiscoverAllTools();
            
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                result = new
                {
                    tools = tools
                },
                id = id
            });
        }
        else if (method == "tools/call")
        {
            // 处理工具调用请求，从params中提取工具名称
            if (!root.TryGetProperty("params", out var paramsElement))
            {
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    error = new { code = -32600, message = "Invalid Request: params is required for tools/call" },
                    id = id
                });
            }
            
            // 从params中提取工具名称
            if (!paramsElement.TryGetProperty("name", out var nameElement))
            {
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    error = new { code = -32600, message = "Invalid Request: name is required in params for tools/call" },
                    id = id
                });
            }
            
            var toolName = nameElement.GetString();
            return await HandleToolCall(toolName, root, serviceProvider);
        }
        else
        {
            // 未知方法
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                error = new { code = -32601, message = "Method not found" },
                id = id
            });
        }
    }
    catch (JsonException)
    {
        // JSON 解析错误
        return JsonSerializer.Serialize(new
        {
            jsonrpc = "2.0",
            error = new { code = -32700, message = "Parse error" },
            id = (object)null
        });
    }
}

static async Task<string> HandleToolCall(string method, JsonElement root, IServiceProvider serviceProvider)
{
    try
    {
        var id = root.TryGetProperty("id", out var idElement) ? 
            (idElement.ValueKind == JsonValueKind.Number ? (object)idElement.GetInt64() : 
             idElement.ValueKind == JsonValueKind.String ? (object)idElement.GetString() : null) : null;
        
        // 检查params字段
        if (!root.TryGetProperty("params", out var paramsElement))
        {
            return JsonSerializer.Serialize(new
            {
                jsonrpc = "2.0",
                error = new { code = -32600, message = "Invalid Request: params is required" },
                id = id
            });
        }

        // 根据方法名调用相应的工具
        switch (method)
        {
            case "open_browser":
                // 从arguments对象中获取url参数
                var url = "";
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement))
                {
                    url = argumentsElement.TryGetProperty("url", out var urlElement) ? urlElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    url = paramsElement.TryGetProperty("url", out var urlElement) ? urlElement.GetString() : null;
                }
                
                // 添加调试日志
                Log.Information("Received URL in open_browser tool call: {Url}", url ?? "null");
                
                // 通过依赖注入获取服务实例
                var desktopService = serviceProvider.GetRequiredService<IDesktopService>();
                var browserResult = await desktopService.OpenBrowserAsync(url, null);
                
                Log.Information("Result from OpenBrowserAsync: {Result}", browserResult);
                
                // 根据MCP协议，返回包含content字段的数组
                var resultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = browserResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = resultContent },
                    id = id
                });

            case "type":
                // 从arguments对象中获取参数
                var x = 0;
                var y = 0;
                var text = "";
                var clear = false;
                var pressEnter = false;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement1))
                {
                    x = argumentsElement1.TryGetProperty("x", out var xElement) ? xElement.GetInt32() : 0;
                    y = argumentsElement1.TryGetProperty("y", out var yElement) ? yElement.GetInt32() : 0;
                    text = argumentsElement1.TryGetProperty("text", out var textElement) ? textElement.GetString() : null;
                    clear = argumentsElement1.TryGetProperty("clear", out var clearElement) ? clearElement.GetBoolean() : false;
                    pressEnter = argumentsElement1.TryGetProperty("pressEnter", out var enterElement) ? enterElement.GetBoolean() : false;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    x = paramsElement.TryGetProperty("x", out var xElement) ? xElement.GetInt32() : 0;
                    y = paramsElement.TryGetProperty("y", out var yElement) ? yElement.GetInt32() : 0;
                    text = paramsElement.TryGetProperty("text", out var textElement) ? textElement.GetString() : null;
                    clear = paramsElement.TryGetProperty("clear", out var clearElement) ? clearElement.GetBoolean() : false;
                    pressEnter = paramsElement.TryGetProperty("pressEnter", out var enterElement) ? enterElement.GetBoolean() : false;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService2 = serviceProvider.GetRequiredService<IDesktopService>();
                await desktopService2.TypeAsync(x, y, text, clear, pressEnter);
                
                // 根据MCP协议，返回包含content字段的数组
                var typeResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = "Text typed successfully"
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = typeResultContent },
                    id = id
                });

            case "get_desktop_state":
                // 从arguments对象中获取参数
                var useVision = false;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement2))
                {
                    useVision = argumentsElement2.TryGetProperty("useVision", out var visionElement) ? visionElement.GetBoolean() : false;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    useVision = paramsElement.TryGetProperty("useVision", out var visionElement) ? visionElement.GetBoolean() : false;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService3 = serviceProvider.GetRequiredService<IDesktopService>();
                var state = await desktopService3.GetDesktopStateAsync();
                
                // 根据MCP协议，返回包含content字段的数组
                var getDesktopStateResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = "Desktop state retrieved successfully"
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = getDesktopStateResultContent },
                    id = id
                });

            // 添加其他工具调用处理...
            case "click":
                // 从arguments对象中获取参数
                var clickX = 0;
                var clickY = 0;
                var button = "left";
                var clicks = 1;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement3))
                {
                    clickX = argumentsElement3.TryGetProperty("x", out var clickXElement) ? clickXElement.GetInt32() : 0;
                    clickY = argumentsElement3.TryGetProperty("y", out var clickYElement) ? clickYElement.GetInt32() : 0;
                    button = argumentsElement3.TryGetProperty("button", out var buttonElement) ? buttonElement.GetString() : "left";
                    clicks = argumentsElement3.TryGetProperty("clicks", out var clicksElement) ? clicksElement.GetInt32() : 1;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    clickX = paramsElement.TryGetProperty("x", out var clickXElement) ? clickXElement.GetInt32() : 0;
                    clickY = paramsElement.TryGetProperty("y", out var clickYElement) ? clickYElement.GetInt32() : 0;
                    button = paramsElement.TryGetProperty("button", out var buttonElement) ? buttonElement.GetString() : "left";
                    clicks = paramsElement.TryGetProperty("clicks", out var clicksElement) ? clicksElement.GetInt32() : 1;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService4 = serviceProvider.GetRequiredService<IDesktopService>();
                var clickResult = await desktopService4.ClickAsync(clickX, clickY, button, clicks);
                
                // 根据MCP协议，返回包含content字段的数组
                var clickResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = clickResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = clickResultContent },
                    id = id
                });

            case "move":
                // 从arguments对象中获取参数
                var moveX = 0;
                var moveY = 0;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement4))
                {
                    moveX = argumentsElement4.TryGetProperty("x", out var moveXElement) ? moveXElement.GetInt32() : 0;
                    moveY = argumentsElement4.TryGetProperty("y", out var moveYElement) ? moveYElement.GetInt32() : 0;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    moveX = paramsElement.TryGetProperty("x", out var moveXElement) ? moveXElement.GetInt32() : 0;
                    moveY = paramsElement.TryGetProperty("y", out var moveYElement) ? moveYElement.GetInt32() : 0;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService5 = serviceProvider.GetRequiredService<IDesktopService>();
                var moveResult = await desktopService5.MoveAsync(moveX, moveY);
                
                // 根据MCP协议，返回包含content字段的数组
                var moveResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = moveResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = moveResultContent },
                    id = id
                });

            case "clipboard":
                // 从arguments对象中获取参数
                var mode = "";
                var clipboardText = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement5))
                {
                    mode = argumentsElement5.TryGetProperty("mode", out var modeElement) ? modeElement.GetString() : null;
                    clipboardText = argumentsElement5.TryGetProperty("text", out var textElement2) ? textElement2.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    mode = paramsElement.TryGetProperty("mode", out var modeElement) ? modeElement.GetString() : null;
                    clipboardText = paramsElement.TryGetProperty("text", out var textElement2) ? textElement2.GetString() : null;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService6 = serviceProvider.GetRequiredService<IDesktopService>();
                var clipboardResult = await desktopService6.ClipboardOperationAsync(mode, clipboardText);
                
                // 根据MCP协议，返回包含content字段的数组
                var clipboardResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = clipboardResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = clipboardResultContent },
                    id = id
                });

            case "powershell":
                // 从arguments对象中获取参数
                var command = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement6))
                {
                    command = argumentsElement6.TryGetProperty("command", out var commandElement) ? commandElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    command = paramsElement.TryGetProperty("command", out var commandElement) ? commandElement.GetString() : null;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService7 = serviceProvider.GetRequiredService<IDesktopService>();
                var (psResponse, psStatus) = await desktopService7.ExecuteCommandAsync(command);
                
                // 根据MCP协议，返回包含content字段的数组
                var powershellResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = psResponse
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = powershellResultContent },
                    id = id
                });

            case "launch":
                // 从arguments对象中获取参数
                var appName = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement7))
                {
                    appName = argumentsElement7.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    appName = paramsElement.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : null;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService8 = serviceProvider.GetRequiredService<IDesktopService>();
                var (launchResponse, launchStatus) = await desktopService8.LaunchAppAsync(appName);
                
                // 根据MCP协议，返回包含content字段的数组
                var launchResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = launchResponse
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = launchResultContent },
                    id = id
                });

            case "shortcut":
                var keysArray = new List<string>();
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement8))
                {
                    if (argumentsElement8.TryGetProperty("keys", out var keysElement) && keysElement.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var keyElement in keysElement.EnumerateArray())
                        {
                            keysArray.Add(keyElement.GetString());
                        }
                    }
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    if (paramsElement.TryGetProperty("keys", out var keysElement) && keysElement.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var keyElement in keysElement.EnumerateArray())
                        {
                            keysArray.Add(keyElement.GetString());
                        }
                    }
                }
                
                // 通过依赖注入获取服务实例
                var desktopService9 = serviceProvider.GetRequiredService<IDesktopService>();
                var shortcutResult = await desktopService9.ShortcutAsync(keysArray.ToArray());
                
                // 根据MCP协议，返回包含content字段的数组
                var shortcutResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = shortcutResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = shortcutResultContent },
                    id = id
                });

            case "key":
                // 从arguments对象中获取参数
                var key = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement9))
                {
                    key = argumentsElement9.TryGetProperty("key", out var keyParamElement) ? keyParamElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    key = paramsElement.TryGetProperty("key", out var keyParamElement) ? keyParamElement.GetString() : null;
                }
                
                // 通过依赖注入获取服务实例
                var desktopService10 = serviceProvider.GetRequiredService<IDesktopService>();
                var keyResult = await desktopService10.KeyAsync(key);
                
                // 根据MCP协议，返回包含content字段的数组
                var keyResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = keyResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = keyResultContent },
                    id = id
                });

            // 文件系统工具
            case "list_directory":
                // 从arguments对象中获取参数
                var path = "";
                var includeFiles = true;
                var includeDirectories = true;
                var recursive = false;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement10))
                {
                    path = argumentsElement10.TryGetProperty("path", out var pathElement) ? pathElement.GetString() : null;
                    includeFiles = argumentsElement10.TryGetProperty("includeFiles", out var includeFilesElement) ? includeFilesElement.GetBoolean() : true;
                    includeDirectories = argumentsElement10.TryGetProperty("includeDirectories", out var includeDirectoriesElement) ? includeDirectoriesElement.GetBoolean() : true;
                    recursive = argumentsElement10.TryGetProperty("recursive", out var recursiveElement) ? recursiveElement.GetBoolean() : false;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    path = paramsElement.TryGetProperty("path", out var pathElement) ? pathElement.GetString() : null;
                    includeFiles = paramsElement.TryGetProperty("includeFiles", out var includeFilesElement) ? includeFilesElement.GetBoolean() : true;
                    includeDirectories = paramsElement.TryGetProperty("includeDirectories", out var includeDirectoriesElement) ? includeDirectoriesElement.GetBoolean() : true;
                    recursive = paramsElement.TryGetProperty("recursive", out var recursiveElement) ? recursiveElement.GetBoolean() : false;
                }
                
                var fileSystemService = serviceProvider.GetRequiredService<IFileSystemService>();
                var (listing, listStatus) = await fileSystemService.ListDirectoryAsync(path, includeFiles, includeDirectories, recursive);
                
                // 根据MCP协议，返回包含content字段的数组
                var listDirectoryResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = listing
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = listDirectoryResultContent },
                    id = id
                });

            case "read_file":
                // 从arguments对象中获取参数
                var readPath = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement11))
                {
                    readPath = argumentsElement11.TryGetProperty("path", out var readPathElement) ? readPathElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    readPath = paramsElement.TryGetProperty("path", out var readPathElement) ? readPathElement.GetString() : null;
                }
                
                var fileSystemService2 = serviceProvider.GetRequiredService<IFileSystemService>();
                var (content, readStatus) = await fileSystemService2.ReadFileAsync(readPath);
                
                // 根据MCP协议，返回包含content字段的数组
                var readFileResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = content
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = readFileResultContent },
                    id = id
                });

            case "write_file":
                // 从arguments对象中获取参数
                var writePath = "";
                var fileContent = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement12))
                {
                    writePath = argumentsElement12.TryGetProperty("path", out var writePathElement) ? writePathElement.GetString() : null;
                    fileContent = argumentsElement12.TryGetProperty("content", out var contentElement) ? contentElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    writePath = paramsElement.TryGetProperty("path", out var writePathElement) ? writePathElement.GetString() : null;
                    fileContent = paramsElement.TryGetProperty("content", out var contentElement) ? contentElement.GetString() : null;
                }
                
                var fileSystemService3 = serviceProvider.GetRequiredService<IFileSystemService>();
                var (writeResponse, writeStatus) = await fileSystemService3.WriteFileAsync(writePath, fileContent);
                
                // 根据MCP协议，返回包含content字段的数组
                var writeFileResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = writeResponse
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = writeFileResultContent },
                    id = id
                });

            case "create_directory":
                // 从arguments对象中获取参数
                var dirPath = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement13))
                {
                    dirPath = argumentsElement13.TryGetProperty("path", out var dirPathElement) ? dirPathElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    dirPath = paramsElement.TryGetProperty("path", out var dirPathElement) ? dirPathElement.GetString() : null;
                }
                
                var fileSystemService4 = serviceProvider.GetRequiredService<IFileSystemService>();
                var (dirResponse, dirStatus) = await fileSystemService4.CreateDirectoryAsync(dirPath);
                
                // 根据MCP协议，返回包含content字段的数组
                var createDirectoryResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = dirResponse
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = createDirectoryResultContent },
                    id = id
                });

            // OCR工具
            case "extract_text_from_screen":
                // 从arguments对象中获取参数
                var extractX = 0;
                var extractY = 0;
                var extractWidth = 0;
                var extractHeight = 0;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement14))
                {
                    extractX = argumentsElement14.TryGetProperty("x", out var extractXElement) ? extractXElement.GetInt32() : 0;
                    extractY = argumentsElement14.TryGetProperty("y", out var extractYElement) ? extractYElement.GetInt32() : 0;
                    extractWidth = argumentsElement14.TryGetProperty("width", out var extractWidthElement) ? extractWidthElement.GetInt32() : 0;
                    extractHeight = argumentsElement14.TryGetProperty("height", out var extractHeightElement) ? extractHeightElement.GetInt32() : 0;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    extractX = paramsElement.TryGetProperty("x", out var extractXElement) ? extractXElement.GetInt32() : 0;
                    extractY = paramsElement.TryGetProperty("y", out var extractYElement) ? extractYElement.GetInt32() : 0;
                    extractWidth = paramsElement.TryGetProperty("width", out var extractWidthElement) ? extractWidthElement.GetInt32() : 0;
                    extractHeight = paramsElement.TryGetProperty("height", out var extractHeightElement) ? extractHeightElement.GetInt32() : 0;
                }
                
                var ocrService = serviceProvider.GetRequiredService<IOcrService>();
                var (extractedText, ocrStatus) = await ocrService.ExtractTextFromRegionAsync(extractX, extractY, extractWidth, extractHeight);
                
                // 根据MCP协议，返回包含content字段的数组
                var extractTextResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = extractedText
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = extractTextResultContent },
                    id = id
                });

            case "find_text_on_screen":
                // 从arguments对象中获取参数
                var searchText = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement15))
                {
                    searchText = argumentsElement15.TryGetProperty("text", out var searchTextElement) ? searchTextElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    searchText = paramsElement.TryGetProperty("text", out var searchTextElement) ? searchTextElement.GetString() : null;
                }
                
                var ocrService2 = serviceProvider.GetRequiredService<IOcrService>();
                var (found, findStatus) = await ocrService2.FindTextOnScreenAsync(searchText);
                
                // 根据MCP协议，返回包含content字段的数组
                var findTextResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = found ? "Text found" : "Text not found"
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = findTextResultContent },
                    id = id
                });

            case "get_text_coordinates":
                // 从arguments对象中获取参数
                var targetText = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement15a))
                {
                    targetText = argumentsElement15a.TryGetProperty("text", out var targetTextElement) ? targetTextElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    targetText = paramsElement.TryGetProperty("text", out var targetTextElement) ? targetTextElement.GetString() : null;
                }
                
                var getTextCoordinatesTool = serviceProvider.GetRequiredService<GetTextCoordinatesTool>();
                var coordinates = await getTextCoordinatesTool.GetTextCoordinatesAsync(targetText);
                
                // 根据MCP协议，返回包含content字段的数组
                var getTextCoordinatesResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = coordinates
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = getTextCoordinatesResultContent },
                    id = id
                });

            // 系统控制工具
            case "set_volume":
                // 从arguments对象中获取参数
                var volumeLevel = 50;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement16))
                {
                    volumeLevel = argumentsElement16.TryGetProperty("level", out var volumeElement) ? volumeElement.GetInt32() : 50;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    volumeLevel = paramsElement.TryGetProperty("level", out var volumeElement) ? volumeElement.GetInt32() : 50;
                }
                
                var systemControlService = serviceProvider.GetRequiredService<ISystemControlService>();
                var volumeResult = await systemControlService.SetVolumePercentAsync(volumeLevel);
                
                // 根据MCP协议，返回包含content字段的数组
                var setVolumeResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = volumeResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = setVolumeResultContent },
                    id = id
                });

            case "set_brightness":
                // 从arguments对象中获取参数
                var brightnessLevel = 50;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement17))
                {
                    brightnessLevel = argumentsElement17.TryGetProperty("level", out var brightnessElement) ? brightnessElement.GetInt32() : 50;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    brightnessLevel = paramsElement.TryGetProperty("level", out var brightnessElement) ? brightnessElement.GetInt32() : 50;
                }
                
                var systemControlService2 = serviceProvider.GetRequiredService<ISystemControlService>();
                var brightnessResult = await systemControlService2.SetBrightnessPercentAsync(brightnessLevel);
                
                // 根据MCP协议，返回包含content字段的数组
                var setBrightnessResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = brightnessResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = setBrightnessResultContent },
                    id = id
                });

            case "screenshot":
                // 截图工具不需要参数
                var screenshotTool = serviceProvider.GetRequiredService<Tools.Desktop.ScreenshotTool>();
                var screenshotPath = await screenshotTool.TakeScreenshotAsync();
                
                // 根据MCP协议，返回包含content字段的数组
                var screenshotResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = screenshotPath
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = screenshotResultContent },
                    id = id
                });

            // 窗口管理工具
            case "switch_window":
                // 从arguments对象中获取参数
                var windowTitle = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement18))
                {
                    windowTitle = argumentsElement18.TryGetProperty("title", out var titleElement) ? titleElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    windowTitle = paramsElement.TryGetProperty("title", out var titleElement) ? titleElement.GetString() : null;
                }
                
                var switchTool = serviceProvider.GetRequiredService<Tools.Desktop.SwitchTool>();
                var switchResult = await switchTool.SwitchAppAsync(windowTitle);
                
                // 根据MCP协议，返回包含content字段的数组
                var switchWindowResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = switchResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = switchWindowResultContent },
                    id = id
                });

            case "get_window_info":
                // 从arguments对象中获取参数
                var infoWindowTitle = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement19))
                {
                    infoWindowTitle = argumentsElement19.TryGetProperty("title", out var titleElement) ? titleElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    infoWindowTitle = paramsElement.TryGetProperty("title", out var titleElement) ? titleElement.GetString() : null;
                }
                
                var getWindowInfoTool = serviceProvider.GetRequiredService<Tools.Desktop.GetWindowInfoTool>();
                var windowInfoResult = await getWindowInfoTool.GetWindowInfoAsync(infoWindowTitle);
                
                // 根据MCP协议，返回包含content字段的数组
                var getWindowInfoResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = windowInfoResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = getWindowInfoResultContent },
                    id = id
                });

            case "resize_window":
                // 从arguments对象中获取参数
                var resizeWindowTitle = "";
                var newWidth = 0;
                var newHeight = 0;
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement20))
                {
                    resizeWindowTitle = argumentsElement20.TryGetProperty("title", out var titleElement) ? titleElement.GetString() : null;
                    newWidth = argumentsElement20.TryGetProperty("width", out var widthElement) ? widthElement.GetInt32() : 0;
                    newHeight = argumentsElement20.TryGetProperty("height", out var heightElement) ? heightElement.GetInt32() : 0;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    resizeWindowTitle = paramsElement.TryGetProperty("title", out var titleElement) ? titleElement.GetString() : null;
                    newWidth = paramsElement.TryGetProperty("width", out var widthElement) ? widthElement.GetInt32() : 0;
                    newHeight = paramsElement.TryGetProperty("height", out var heightElement) ? heightElement.GetInt32() : 0;
                }
                
                var resizeWindowTool = serviceProvider.GetRequiredService<Tools.Desktop.ResizeTool>();
                var resizeResult = await resizeWindowTool.ResizeAppAsync(resizeWindowTitle, newWidth, newHeight);
                
                // 根据MCP协议，返回包含content字段的数组
                var resizeWindowResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = resizeResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = resizeWindowResultContent },
                    id = id
                });

            case "ui_element":
                // 从arguments对象中获取参数
                var elementName = "";
                
                if (paramsElement.TryGetProperty("arguments", out var argumentsElement21))
                {
                    elementName = argumentsElement21.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : null;
                }
                else
                {
                    // 向后兼容：如果没有arguments对象，直接从params中获取
                    elementName = paramsElement.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : null;
                }
                
                var uiElementTool = serviceProvider.GetRequiredService<Tools.Desktop.UIElementTool>();
                var uiElementResult = await uiElementTool.FindElementByTextAsync(elementName);
                
                // 根据MCP协议，返回包含content字段的数组
                var uiElementResultContent = new[]
                {
                    new
                    {
                        type = "text",
                        text = uiElementResult
                    }
                };
                
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    result = new { content = uiElementResultContent },
                    id = id
                });

            default:
                return JsonSerializer.Serialize(new
                {
                    jsonrpc = "2.0",
                    error = new { code = -32601, message = "Method not found" },
                    id = id
                });
        }
    }
    catch (Exception ex)
    {
        var id = root.TryGetProperty("id", out var idElement) ? 
            (idElement.ValueKind == JsonValueKind.Number ? (object)idElement.GetInt64() : 
             idElement.ValueKind == JsonValueKind.String ? (object)idElement.GetString() : null) : null;
        
        return JsonSerializer.Serialize(new
        {
            jsonrpc = "2.0",
            error = new { code = -32000, message = "Internal error", data = ex.Message },
            id = id
        });
    }
}

// 自动发现所有工具的方法
static object[] DiscoverAllTools()
{
    var tools = new List<object>();
    
    // 添加 open_browser 工具
    tools.Add(new 
    { 
        name = "open_browser", 
        description = "Open a URL in the default browser",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                url = new { type = "string", description = "The URL to open" }
            }
        }
    });

    // 添加 type 工具
    tools.Add(new
    {
        name = "type",
        description = "Type text at specified coordinates",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                x = new { type = "integer", description = "X coordinate" },
                y = new { type = "integer", description = "Y coordinate" },
                text = new { type = "string", description = "Text to type" },
                clear = new { type = "boolean", description = "Clear existing text" },
                pressEnter = new { type = "boolean", description = "Press Enter after typing" }
            }
        }
    });

    // 添加 get_desktop_state 工具
    tools.Add(new
    {
        name = "get_desktop_state",
        description = "Get current desktop state",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                useVision = new { type = "boolean", description = "Use vision for state detection" }
            }
        }
    });

    // 添加 click 工具
    tools.Add(new
    {
        name = "click",
        description = "Click at specified coordinates",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                x = new { type = "integer", description = "X coordinate" },
                y = new { type = "integer", description = "Y coordinate" },
                button = new { type = "string", description = "Mouse button (left, right, middle)" },
                clicks = new { type = "integer", description = "Number of clicks" }
            }
        }
    });

    // 添加 move 工具
    tools.Add(new
    {
        name = "move",
        description = "Move mouse to specified coordinates",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                x = new { type = "integer", description = "X coordinate" },
                y = new { type = "integer", description = "Y coordinate" }
            }
        }
    });

    // 添加 clipboard 工具
    tools.Add(new
    {
        name = "clipboard",
        description = "Perform clipboard operations",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                mode = new { type = "string", description = "Operation mode (get, set, clear)" },
                text = new { type = "string", description = "Text to set (for set mode)" }
            }
        }
    });

    // 添加 powershell 工具
    tools.Add(new
    {
        name = "powershell",
        description = "Execute PowerShell command",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                command = new { type = "string", description = "PowerShell command to execute" }
            }
        }
    });

    // 添加 launch 工具
    tools.Add(new
    {
        name = "launch",
        description = "Launch application",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                name = new { type = "string", description = "Application name or path" }
            }
        }
    });

    // 添加 shortcut 工具
    tools.Add(new
    {
        name = "shortcut",
        description = "Execute keyboard shortcut",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                keys = new { type = "array", items = new { type = "string" }, description = "Array of keys to press" }
            }
        }
    });

    // 添加 key 工具
    tools.Add(new
    {
        name = "key",
        description = "Press a single key",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                key = new { type = "string", description = "Key to press" }
            }
        }
    });

    // 添加文件系统工具
    tools.Add(new
    {
        name = "list_directory",
        description = "List contents of a directory",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                path = new { type = "string", description = "The directory path to list" },
                includeFiles = new { type = "boolean", description = "Whether to include files in the listing" },
                includeDirectories = new { type = "boolean", description = "Whether to include directories in the listing" },
                recursive = new { type = "boolean", description = "Whether to list recursively" }
            }
        }
    });

    tools.Add(new
    {
        name = "read_file",
        description = "Read contents of a file",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                path = new { type = "string", description = "The file path to read" }
            }
        }
    });

    tools.Add(new
    {
        name = "write_file",
        description = "Write content to a file",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                path = new { type = "string", description = "The file path to write to" },
                content = new { type = "string", description = "The content to write" }
            }
        }
    });

    tools.Add(new
    {
        name = "create_directory",
        description = "Create a new directory",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                path = new { type = "string", description = "The directory path to create" }
            }
        }
    });

    // 添加OCR工具
    tools.Add(new
    {
        name = "extract_text_from_screen",
        description = "Extract text from screen region",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                x = new { type = "integer", description = "X coordinate of region" },
                y = new { type = "integer", description = "Y coordinate of region" },
                width = new { type = "integer", description = "Width of region" },
                height = new { type = "integer", description = "Height of region" }
            }
        }
    });

    tools.Add(new
    {
        name = "find_text_on_screen",
        description = "Find text on screen and return coordinates",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                text = new { type = "string", description = "Text to find" }
            }
        }
    });

    tools.Add(new
    {
        name = "get_text_coordinates",
        description = "Get the coordinates of specific text on the screen using OCR",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                text = new { type = "string", description = "The text to locate on the screen" }
            }
        }
    });

    // 添加系统控制工具
    tools.Add(new
    {
        name = "set_volume",
        description = "Set system volume level",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                level = new { type = "integer", description = "Volume level (0-100)" }
            }
        }
    });

    tools.Add(new
    {
        name = "set_brightness",
        description = "Set display brightness level",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                level = new { type = "integer", description = "Brightness level (0-100)" }
            }
        }
    });

    // 添加截图工具
    tools.Add(new
    {
        name = "screenshot",
        description = "Take a screenshot and save it to the temp directory",
        inputSchema = new
        {
            type = "object",
            properties = new { }
        }
    });

    // 添加窗口管理工具
    tools.Add(new
    {
        name = "switch_window",
        description = "Switch to a specific application window and bring it to foreground",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                title = new { type = "string", description = "The title of the window to switch to" }
            }
        }
    });

    tools.Add(new
    {
        name = "get_window_info",
        description = "Get information about application windows",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                title = new { type = "string", description = "The title of the window to get info for" }
            }
        }
    });

    tools.Add(new
    {
        name = "resize_window",
        description = "Resize an application window",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                title = new { type = "string", description = "The title of the window to resize" },
                width = new { type = "integer", description = "New width of the window" },
                height = new { type = "integer", description = "New height of the window" }
            }
        }
    });

    tools.Add(new
    {
        name = "ui_element",
        description = "Get UI elements from the current window",
        inputSchema = new
        {
            type = "object",
            properties = new
            {
                name = new { type = "string", description = "The name of the UI element to find" }
            }
        }
    });

    return tools.ToArray();
}

static async Task RunOcrTest()
{
    try
    {
        Log.Information("开始测试OCR服务功能...");
        
        // 创建OCR服务实例 - 使用默认构造函数（无参数）
        var ocrService = new OcrService();
        
        // 测试1: 获取操作统计信息
        Log.Information("=== 测试1: 获取操作统计信息 ===");
        var stats = ocrService.GetOperationStatistics();
        Log.Information($"操作统计: 总操作数={stats.Total}, 失败操作数={stats.Failed}, 并发操作数={stats.Concurrent}");
        
        // 测试2: 创建测试图像并提取文本
        Log.Information("=== 测试2: 图像文本提取 ===");
        await TestImageTextExtraction(ocrService);
        
        // 测试3: 再次获取操作统计信息
        Log.Information("=== 测试3: 再次获取操作统计信息 ===");
        stats = ocrService.GetOperationStatistics();
        Log.Information($"操作统计: 总操作数={stats.Total}, 失败操作数={stats.Failed}, 并发操作数={stats.Concurrent}");
        
        Log.Information("=== 测试完成 ===");
    }
    catch (Exception ex)
    {
        Log.Error(ex, "测试过程中发生错误");
    }
}

static async Task TestImageTextExtraction(OcrService ocrService)
{
    try
    {
        // 创建一个简单的测试图像（包含文本）
        using var bitmap = new System.Drawing.Bitmap(200, 50);
        using var graphics = System.Drawing.Graphics.FromImage(bitmap);
        graphics.Clear(System.Drawing.Color.White);
        using var font = new System.Drawing.Font("Arial", 12);
        using var brush = new System.Drawing.SolidBrush(System.Drawing.Color.Black);
        graphics.DrawString("Hello World", font, brush, new System.Drawing.PointF(10, 10));
        
        // 将图像保存到内存流
        using var memoryStream = new System.IO.MemoryStream();
        bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
        memoryStream.Position = 0;
        
        Log.Information("开始从测试图像提取文本...");
        
        // 调用OCR服务
        var result = await ocrService.ExtractTextFromImageAsync(memoryStream);
        
        Log.Information($"OCR结果: 文本='{result.Text}', 状态={result.Status}");
        
        if (result.Status == 0 && !string.IsNullOrEmpty(result.Text))
        {
            Log.Information("✅ 图像文本提取测试成功");
        }
        else
        {
            Log.Warning("⚠️ 图像文本提取测试结果不理想");
        }
    }
    catch (Exception ex)
    {
        Log.Error(ex, "图像文本提取测试失败");
    }
}