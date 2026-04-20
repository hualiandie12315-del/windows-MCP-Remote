using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Extensions.Logging;
using OpenCvSharp;
using Sdcb.OpenVINO.PaddleOCR;
using Sdcb.OpenVINO.PaddleOCR.Models.Online;
using Sdcb.OpenVINO.PaddleOCR.Models;
using System.Runtime.InteropServices;
using Interface;


namespace WindowsMCP.Net.Services;

/// <summary>
/// Implementation of OCR services using PaddleOCR.
/// </summary>
public class OcrService : IOcrService, IDisposable
{
    private readonly ILogger<OcrService> _logger;
    private static readonly Lazy<OcrService> _instance = new Lazy<OcrService>(() => new OcrService());
    private static readonly object _lock = new object();
    private static FullOcrModel? _model = null;
    private bool _disposed = false;
    
    // 并发控制
    private readonly SemaphoreSlim _ocrSemaphore = new SemaphoreSlim(2, 2); // 限制同时运行的OCR任务数量
    private static readonly TimeSpan _operationTimeout = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan _minOperationInterval = TimeSpan.FromMilliseconds(500); // 最小操作间隔
    private static DateTime _lastOperationTime = DateTime.MinValue;
    
    // 资源监控
    private static int _concurrentOperations = 0;
    private static int _totalOperations = 0;
    private static int _failedOperations = 0;

    // Windows API imports for screen capture
    [DllImport("user32.dll")]
    private static extern IntPtr GetDesktopWindow();

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindowDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hDC);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth, int nHeight);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hDC, IntPtr hGDIObj);

    [DllImport("gdi32.dll")]
    private static extern bool BitBlt(IntPtr hDestDC, int x, int y, int nWidth, int nHeight, IntPtr hSrcDC, int xSrc, int ySrc, int dwRop);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hDC);

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    // DPI-related API imports
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("gdi32.dll")]
    private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    // System metrics constants
    private const int SM_CXSCREEN = 0; // Width of the screen
    private const int SM_CYSCREEN = 1; // Height of the screen

    // DPI-related constants
    private const int DESKTOPHORZRES = 118; // Physical screen width
    private const int DESKTOPVERTRES = 117; // Physical screen height
    private const int LOGPIXELSX = 88; // Logical pixels per inch in X
    private const int LOGPIXELSY = 90; // Logical pixels per inch in Y

    /// <summary>
    /// 获取单例实例
    /// </summary>
    public static OcrService Instance => _instance.Value;

    /// <summary>
    /// 初始化OCR服务
    /// </summary>
    /// <param name="logger">日志记录器</param>
    public OcrService(ILogger<OcrService> logger = null)
    {
        _logger = logger ?? Microsoft.Extensions.Logging.Abstractions.NullLogger<OcrService>.Instance;
    }

    /// <summary>
    /// 私有构造函数，用于单例模式
    /// </summary>
    private OcrService()
    {
        _logger = Microsoft.Extensions.Logging.Abstractions.NullLogger<OcrService>.Instance;
    }

    /// <summary>
    /// Initialize the OCR model asynchronously.
    /// </summary>
    private async Task InitializeModelAsync()
        {
            if (_model == null)
            {
                lock (_lock)
                {
                    if (_model == null)
                    {
                        try
                        {
                            _logger.LogInformation("Initializing OCR model...");
                            
                            // 诊断信息：检查OpenCV是否可用
                            try
                            {
                                var version = Cv2.GetVersionString();
                                _logger.LogInformation("OpenCV version: {Version}", version);
                            }
                            catch (Exception cvEx)
                            {
                                _logger.LogError(cvEx, "Failed to get OpenCV version - native libraries may not be loaded correctly");
                                throw new InvalidOperationException("OpenCV native libraries are not available. This may be due to missing runtime dependencies in the dnx environment.", cvEx);
                            }
                            
                            // 诊断信息：检查运行时环境
                            var runtimeInfo = System.Runtime.InteropServices.RuntimeInformation.RuntimeIdentifier;
                            var architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture;
                            _logger.LogInformation("Runtime: {Runtime}, Architecture: {Architecture}", runtimeInfo, architecture);
                            
                            _logger.LogInformation("Downloading OCR model...");
                            _model = OnlineFullModels.ChineseV4.DownloadAsync().Result;
                            _logger.LogInformation("OCR model downloaded and initialized successfully");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Failed to initialize OCR model. Error: {Message}", ex.Message);
                            throw;
                        }
                    }
                }
            }
        }

    /// <summary>
    /// 安全地处理OpenCV Mat对象，确保内存正确释放
    /// </summary>
    private Mat SafeImDecode(byte[] imageData, ImreadModes mode = ImreadModes.Color)
    {
        try
        {
            if (imageData == null || imageData.Length == 0)
            {
                throw new ArgumentException("Image data is null or empty");
            }
            
            // 使用缓冲区复制来隔离内存
            byte[] bufferCopy = new byte[imageData.Length];
            Buffer.BlockCopy(imageData, 0, bufferCopy, 0, imageData.Length);
            
            Mat src = Cv2.ImDecode(bufferCopy, mode);
            
            if (src.Empty())
            {
                src?.Dispose();
                throw new InvalidOperationException("Failed to decode image - result image is empty");
            }
            
            return src;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in SafeImDecode: {Message}", ex.Message);
            throw;
        }
    }

    /// <summary>
    /// 安全地运行OCR处理，确保资源正确释放
    /// </summary>
    private PaddleOcrResult SafeRunOcr(Mat src, PaddleOcrAll ocrEngine)
    {
        if (src == null || src.Empty())
        {
            throw new ArgumentException("Source image is null or empty");
        }
        
        if (ocrEngine == null)
        {
            throw new ArgumentNullException(nameof(ocrEngine));
        }
        
        try
        {
            // 创建图像副本以避免原始图像被修改
            using var srcCopy = new Mat();
            src.CopyTo(srcCopy);
            
            return ocrEngine.Run(srcCopy);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error in SafeRunOcr: {Message}", ex.Message);
            throw;
        }
    }

    /// <summary>
    /// 带重试机制的OCR操作执行
    /// </summary>
    private async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> operation, string operationName, int maxRetries = 2, int baseDelayMs = 1000)
    {
        int attempt = 0;
        Exception lastException = null;
        
        while (attempt <= maxRetries)
        {
            try
            {
                if (attempt > 0)
                {
                    _logger.LogWarning("Retry attempt {Attempt} for {OperationName}", attempt, operationName);
                    
                    // 指数退避策略
                    int delayMs = baseDelayMs * (int)Math.Pow(2, attempt - 1);
                    await Task.Delay(delayMs);
                }
                
                return await operation();
            }
            catch (AccessViolationException ex)
            {
                lastException = ex;
                _logger.LogWarning(ex, "AccessViolationException in {OperationName} (attempt {Attempt})", operationName, attempt + 1);
                
                // 对于内存访问违规，尝试强制垃圾回收
                if (attempt < maxRetries)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            catch (Exception ex) when (ex.Message.Contains("内存") || ex.Message.Contains("memory"))
            {
                lastException = ex;
                _logger.LogWarning(ex, "Memory-related exception in {OperationName} (attempt {Attempt})", operationName, attempt + 1);
            }
            catch (Exception ex) when (ex.InnerException is AccessViolationException)
            {
                lastException = ex;
                _logger.LogWarning(ex, "Inner AccessViolationException in {OperationName} (attempt {Attempt})", operationName, attempt + 1);
            }
            catch (Exception ex)
            {
                lastException = ex;
                _logger.LogWarning(ex, "Exception in {OperationName} (attempt {Attempt})", operationName, attempt + 1);
            }
            
            attempt++;
        }
        
        _logger.LogError(lastException, "Failed to execute {OperationName} after {MaxRetries} retries", operationName, maxRetries);
        throw new InvalidOperationException($"Failed to execute {operationName} after {maxRetries} retries", lastException);
    }

    /// <summary>
    /// 优雅降级处理 - 当OCR失败时提供基本功能
    /// </summary>
    private (string? Text, int Status) HandleOcrFailure(Exception ex, string operationName)
    {
        _logger.LogError(ex, "OCR operation {OperationName} failed, using fallback", operationName);
        
        // 记录失败统计
        Interlocked.Increment(ref _failedOperations);
        
        // 根据错误类型提供不同的降级策略
        if (ex is AccessViolationException || ex.InnerException is AccessViolationException)
        {
            // 内存访问违规 - 建议重启服务或等待资源释放
            _logger.LogWarning("Memory access violation detected, consider restarting the OCR service");
            return (null, 2); // 特殊错误码表示需要重启
        }
        else if (ex.Message.Contains("内存") || ex.Message.Contains("memory"))
        {
            // 内存相关错误
            return (null, 3); // 内存错误
        }
        
        // 通用错误
        return (null, 1);
    }

    /// <summary>
    /// Extract text from a specific region of the screen.
    /// </summary>
    public async Task<(string Text, int Status)> ExtractTextFromRegionAsync(int x, int y, int width, int height, CancellationToken cancellationToken = default)
    {
        return await ExecuteWithConcurrencyControlAsync(async () =>
        {
            try
            {
                _logger.LogInformation("Extracting text from screen region: ({X}, {Y}, {Width}, {Height})", x, y, width, height);
                
                await InitializeModelAsync();
                
                using var bitmap = CaptureScreenRegion(x, y, width, height);
                using var stream = new MemoryStream();
                bitmap.Save(stream, ImageFormat.Png);
                stream.Position = 0;
                
                var (text, status) = await ExtractTextFromImageAsync(stream, cancellationToken);
                return (text, status);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting text from screen region");
                return (string.Empty, 1);
            }
        }, "ExtractTextFromRegionAsync");
    }

    /// <summary>
    /// Extract text from the entire screen.
    /// </summary>
    public async Task<(string Text, int Status)> ExtractTextFromScreenAsync(CancellationToken cancellationToken = default)
    {
        return await ExecuteWithConcurrencyControlAsync(async () =>
        {
            try
            {
                _logger.LogInformation("Extracting text from entire screen");
                
                // 使用DPI感知的屏幕尺寸
                var dpiAwareScreenSize = GetDpiAwareScreenSize();
                int screenWidth = dpiAwareScreenSize.Width;
                int screenHeight = dpiAwareScreenSize.Height;
                
                _logger.LogInformation("使用DPI感知的屏幕尺寸: {Width}x{Height}", screenWidth, screenHeight);
                return await ExtractTextFromRegionAsync(0, 0, screenWidth, screenHeight, cancellationToken);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting text from screen");
                return (string.Empty, 1);
            }
        }, "ExtractTextFromScreenAsync");
    }

    /// <summary>
    /// Find specific text on the screen.
    /// </summary>
    public async Task<(bool Found, int Status)> FindTextOnScreenAsync(string text, CancellationToken cancellationToken = default)
    {
        try
        {
            _logger.LogInformation("Searching for text on screen: {Text}", text);
            
            var (extractedText, status) = await ExtractTextFromScreenAsync(cancellationToken);
            if (status != 0)
            {
                return (false, status);
            }
            
            bool found = extractedText.Contains(text, StringComparison.OrdinalIgnoreCase);
            return (found, 0);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error searching for text on screen");
            return (false, 1);
        }
    }

    /// <summary>
    /// Get coordinates of specific text on the screen.
    /// </summary>
    public async Task<(System.Drawing.Point? Coordinates, int Status)> GetTextCoordinatesAsync(string text, CancellationToken cancellationToken = default)
    {
        return await ExecuteWithConcurrencyControlAsync(async () =>
        {
            try
            {
                _logger.LogInformation("Getting coordinates for text: {Text}", text);
                
                await InitializeModelAsync();
                
                // 获取考虑DPI缩放的屏幕尺寸
                var dpiAwareScreenSize = GetDpiAwareScreenSize();
                int screenWidth = dpiAwareScreenSize.Width;
                int screenHeight = dpiAwareScreenSize.Height;
                
                _logger.LogInformation("使用DPI感知的屏幕尺寸: {Width}x{Height}", screenWidth, screenHeight);
                
                using var bitmap = CaptureScreenRegion(0, 0, screenWidth, screenHeight);
                using var stream = new MemoryStream();
                bitmap.Save(stream, ImageFormat.Png);
                stream.Position = 0;
                
                using (PaddleOcrAll all = new(_model)
                {
                    AllowRotateDetection = true,
                    Enable180Classification = true
                })
                {
                    using var src = SafeImDecode(StreamToByte(stream), ImreadModes.Color);
                    PaddleOcrResult result = SafeRunOcr(src, all);
                    
                    _logger.LogInformation("OCR识别结果: " + (result.Text ?? "null"));
                    _logger.LogInformation("识别到的区域数量: " + result.Regions.Count().ToString());
                    
                    // 查找包含指定文本的区域
                    var bestMatch = (region: (PaddleOcrResultRegion?)null, similarity: 0.0, confidence: 0);
                    
                    foreach (var region in result.Regions)
                    {
                        string regionText = region.Text?.Trim() ?? "";
                        string searchText = text?.Trim() ?? "";
                        
                        _logger.LogInformation("识别到的文本区域: '" + regionText + "' at (" + (int)(region.Rect.Center.X + region.Rect.Size.Width / 2) + ", " + (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2) + ")");
                        
                        // 改进的文本匹配逻辑：多级匹配策略
                        int matchLevel = 0;
                        
                        // 1. 首先尝试精确匹配
                        if (string.Equals(regionText, searchText, StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogInformation("通过精确匹配找到文本");
                            var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                    var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                            var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                            return (calibratedPoint, 0);
                        }
                        matchLevel = 1;
                        
                        // 2. 尝试包含匹配（去除空格和特殊字符）
                        string cleanRegionText = System.Text.RegularExpressions.Regex.Replace(regionText, @"\s+", "");
                        string cleanSearchText = System.Text.RegularExpressions.Regex.Replace(searchText, @"\s+", "");
                        
                        if (cleanRegionText.Contains(cleanSearchText, StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogInformation("通过清理后包含匹配找到文本");
                            var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                             var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                            var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                            return (calibratedPoint, 0);
                        }
                        matchLevel = 2;
                        
                        // 3. 尝试部分匹配（处理OCR可能的分词问题）
                        if (!string.IsNullOrEmpty(regionText) && !string.IsNullOrEmpty(searchText) && 
                            regionText.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogInformation("通过部分匹配找到文本");
                            var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                             var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                            var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                            return (calibratedPoint, 0);
                        }
                        matchLevel = 3;
                        
                        // 4. 尝试搜索文本包含在识别文本中
                        if (!string.IsNullOrEmpty(regionText) && !string.IsNullOrEmpty(searchText) && 
                            searchText.Contains(regionText, StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogInformation("通过反向包含找到文本");
                            var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                             var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                            var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                            return (calibratedPoint, 0);
                        }
                        matchLevel = 4;
                        
                        // 5. 处理OCR可能的分词问题：将搜索文本拆分为单词进行匹配
                        var searchWords = searchText.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                        var regionWords = regionText.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                        
                        if (searchWords.Length > 0 && regionWords.Length > 0)
                        {
                            bool allWordsFound = searchWords.All(word => 
                                regionWords.Any(rw => rw.Contains(word, StringComparison.OrdinalIgnoreCase)));
                            
                            if (allWordsFound)
                            {
                                _logger.LogInformation("通过单词匹配找到文本");
                            var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                             var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                            var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                            return (calibratedPoint, 0);
                            }
                        }
                        matchLevel = 5;
                        
                        // 6. 使用相似度算法进行模糊匹配
                        double similarity = CalculateTextSimilarity(regionText, searchText);
                        if (similarity > 0.7) // 相似度阈值
                        {
                            _logger.LogInformation($"通过相似度匹配找到文本，相似度: {similarity:F2}");
                            var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                             var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                            var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                            return (calibratedPoint, 0);
                        }
                        
                        // 记录最佳匹配
                        if (similarity > bestMatch.similarity)
                        {
                            bestMatch = (region, similarity, matchLevel);
                        }
                    }
                    
                    // 如果没有精确匹配，但存在相似度较高的匹配，返回最佳匹配
                    if (bestMatch.region != null && bestMatch.similarity > 0.5)
                    {
                        _logger.LogInformation($"返回最佳匹配文本，相似度: {bestMatch.similarity:F2}, 匹配级别: {bestMatch.confidence}");
                        var region = bestMatch.region.Value; // 获取非空值
                        var centerX = (int)(region.Rect.Center.X + region.Rect.Size.Width / 2);
                             var centerY = (int)(region.Rect.Center.Y + region.Rect.Size.Height / 2);
                        var calibratedPoint = CalibrateCoordinates(centerX, centerY);
                        return (calibratedPoint, 0);
                    }
                }
                
                _logger.LogInformation("未找到文本: {Text}", text);
                return (System.Drawing.Point.Empty, 0); // 未找到文本，返回空点
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting text coordinates");
                return (System.Drawing.Point.Empty, 1);
            }
        }, "GetTextCoordinatesAsync");
    }

    /// <summary>
    /// Extract text from an image stream.
    /// </summary>
    public async Task<(string Text, int Status)> ExtractTextFromImageAsync(Stream imageStream, CancellationToken cancellationToken = default)
    {
        return await ExecuteWithConcurrencyControlAsync(async () =>
        {
            try
            {
                _logger.LogInformation("Extracting text from image stream");
                
                await InitializeModelAsync();

                using (PaddleOcrAll all = new(_model)
                {
                    AllowRotateDetection = true,
                    Enable180Classification = true
                })
                {
                    using var src = SafeImDecode(StreamToByte(imageStream), ImreadModes.Color);
                    PaddleOcrResult result = SafeRunOcr(src, all);
                    
                    _logger.LogInformation("OCR识别结果: " + (result.Text ?? "null"));
                    
                    return (result.Text, 0);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting text from image");
                return (string.Empty, 1);
            }
        }, "ExtractTextFromImageAsync");
    }

    /// <summary>
    /// 获取考虑DPI缩放的屏幕尺寸
    /// </summary>
    private System.Drawing.Size GetDpiAwareScreenSize()
    {
        try
        {
            IntPtr desktopDC = GetDC(IntPtr.Zero);
            if (desktopDC != IntPtr.Zero)
            {
                // 获取物理屏幕尺寸（考虑DPI缩放）
                int physicalWidth = GetDeviceCaps(desktopDC, DESKTOPHORZRES);
                int physicalHeight = GetDeviceCaps(desktopDC, DESKTOPVERTRES);
                ReleaseDC(IntPtr.Zero, desktopDC);
                
                // 如果获取物理尺寸失败，则回退到系统指标
                if (physicalWidth > 0 && physicalHeight > 0)
                {
                    _logger.LogInformation("使用物理屏幕尺寸: {Width}x{Height}", physicalWidth, physicalHeight);
                    return new System.Drawing.Size(physicalWidth, physicalHeight);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "获取DPI感知屏幕尺寸失败，使用系统指标");
        }
        
        // 回退到系统指标
        int screenWidth = GetSystemMetrics(SM_CXSCREEN);
        int screenHeight = GetSystemMetrics(SM_CYSCREEN);
        _logger.LogInformation("使用系统指标屏幕尺寸: {Width}x{Height}", screenWidth, screenHeight);
        return new System.Drawing.Size(screenWidth, screenHeight);
    }

    /// <summary>
    /// 计算文本相似度（Levenshtein距离）
    /// </summary>
    private double CalculateTextSimilarity(string text1, string text2)
    {
        if (string.IsNullOrEmpty(text1) || string.IsNullOrEmpty(text2))
            return 0.0;

        if (text1 == text2)
            return 1.0;

        int maxLength = Math.Max(text1.Length, text2.Length);
        if (maxLength == 0)
            return 1.0;

        // 计算Levenshtein距离
        int[,] distance = new int[text1.Length + 1, text2.Length + 1];

        for (int i = 0; i <= text1.Length; i++)
            distance[i, 0] = i;
        for (int j = 0; j <= text2.Length; j++)
            distance[0, j] = j;

        for (int i = 1; i <= text1.Length; i++)
        {
            for (int j = 1; j <= text2.Length; j++)
            {
                int cost = (text1[i - 1] == text2[j - 1]) ? 0 : 1;
                distance[i, j] = Math.Min(
                    Math.Min(distance[i - 1, j] + 1, distance[i, j - 1] + 1),
                    distance[i - 1, j - 1] + cost);
            }
        }

        int levenshteinDistance = distance[text1.Length, text2.Length];
        return 1.0 - (double)levenshteinDistance / maxLength;
    }

    /// <summary>
    /// 校准坐标，考虑DPI缩放和屏幕缩放因素
    /// </summary>
    private System.Drawing.Point CalibrateCoordinates(int x, int y)
    {
        try
        {
            // 获取系统DPI缩放比例
            IntPtr desktopDC = GetDC(IntPtr.Zero);
            if (desktopDC != IntPtr.Zero)
            {
                int dpiX = GetDeviceCaps(desktopDC, LOGPIXELSX);
                int dpiY = GetDeviceCaps(desktopDC, LOGPIXELSY);
                ReleaseDC(IntPtr.Zero, desktopDC);
                
                // 标准DPI为96，计算缩放比例
                double scaleX = dpiX / 96.0;
                double scaleY = dpiY / 96.0;
                
                _logger.LogInformation("DPI缩放比例: X={ScaleX:F2}, Y={ScaleY:F2}", scaleX, scaleY);
                
                // 应用缩放校准
                int calibratedX = (int)(x / scaleX);
                int calibratedY = (int)(y / scaleY);
                
                _logger.LogInformation("坐标校准: ({X}, {Y}) -> ({CalibratedX}, {CalibratedY})", x, y, calibratedX, calibratedY);
                
                return new System.Drawing.Point(calibratedX, calibratedY);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "坐标校准失败，使用原始坐标");
        }
        
        // 回退到原始坐标
        return new System.Drawing.Point(x, y);
    }

    /// <summary>
    /// 截取屏幕指定区域（增强资源管理版本）
    /// </summary>
    private Bitmap CaptureScreenRegion(int x, int y, int width, int height)
    {
        IntPtr desktopWindow = IntPtr.Zero;
        IntPtr desktopDC = IntPtr.Zero;
        IntPtr memoryDC = IntPtr.Zero;
        IntPtr bitmap = IntPtr.Zero;
        IntPtr oldBitmap = IntPtr.Zero;
        
        try
        {
            desktopWindow = GetDesktopWindow();
            desktopDC = GetWindowDC(desktopWindow);
            
            if (desktopDC == IntPtr.Zero)
            {
                throw new InvalidOperationException("Failed to get desktop device context");
            }
            
            memoryDC = CreateCompatibleDC(desktopDC);
            if (memoryDC == IntPtr.Zero)
            {
                throw new InvalidOperationException("Failed to create compatible device context");
            }
            
            bitmap = CreateCompatibleBitmap(desktopDC, width, height);
            if (bitmap == IntPtr.Zero)
            {
                throw new InvalidOperationException("Failed to create compatible bitmap");
            }
            
            oldBitmap = SelectObject(memoryDC, bitmap);
            
            bool bitBltResult = BitBlt(memoryDC, 0, 0, width, height, desktopDC, x, y, 0x00CC0020); // SRCCOPY
            if (!bitBltResult)
            {
                _logger.LogWarning("BitBlt operation may have failed");
            }
            
            // 恢复原始位图选择
            if (oldBitmap != IntPtr.Zero)
            {
                SelectObject(memoryDC, oldBitmap);
            }
            
            // 创建Bitmap对象
            Bitmap result = Image.FromHbitmap(bitmap);
            
            _logger.LogDebug("Screen capture completed successfully for region: ({X}, {Y}, {Width}, {Height})", x, y, width, height);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during screen capture for region: ({X}, {Y}, {Width}, {Height})", x, y, width, height);
            throw;
        }
        finally
        {
            // 确保所有GDI资源都被正确释放
            if (oldBitmap != IntPtr.Zero && memoryDC != IntPtr.Zero)
            {
                SelectObject(memoryDC, oldBitmap);
            }
            
            if (bitmap != IntPtr.Zero)
            {
                DeleteObject(bitmap);
            }
            
            if (memoryDC != IntPtr.Zero)
            {
                DeleteDC(memoryDC);
            }
            
            if (desktopDC != IntPtr.Zero && desktopWindow != IntPtr.Zero)
            {
                ReleaseDC(desktopWindow, desktopDC);
            }
        }
    }

    /// <summary>
    /// 将Stream转换为byte数组
    /// </summary>
    private byte[] StreamToByte(Stream stream)
    {
        using (var memoryStream = new MemoryStream())
        {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                memoryStream.Write(buffer, 0, bytesRead);
            }
            return memoryStream.ToArray();
        }
    }

    /// <summary>
    /// 执行带并发控制的OCR操作
    /// </summary>
    private async Task<T> ExecuteWithConcurrencyControlAsync<T>(Func<Task<T>> operation, string operationName)
    {
        Interlocked.Increment(ref _totalOperations);
        
        try
        {
            // 检查并发限制
            if (!await _ocrSemaphore.WaitAsync(_operationTimeout))
            {
                Interlocked.Increment(ref _failedOperations);
                throw new TimeoutException($"Operation {operationName} timed out waiting for concurrency slot");
            }
            
            try
            {
                // 强制执行操作间隔
                await EnforceOperationIntervalAsync();
                
                // 记录操作开始时间
                var operationStartTime = DateTime.UtcNow;
                Interlocked.Increment(ref _concurrentOperations);
                
                _logger.LogDebug("Starting {OperationName} (concurrent: {ConcurrentCount})", 
                    operationName, _concurrentOperations);
                
                // 执行带重试机制的操作
                var result = await ExecuteWithRetryAsync(operation, operationName);
                
                // 记录操作完成时间
                var operationDuration = DateTime.UtcNow - operationStartTime;
                _logger.LogDebug("Completed {OperationName} in {Duration}ms", 
                    operationName, operationDuration.TotalMilliseconds);
                
                return result;
            }
            finally
            {
                Interlocked.Decrement(ref _concurrentOperations);
                _ocrSemaphore.Release();
            }
        }
        catch (Exception ex)
        {
            Interlocked.Increment(ref _failedOperations);
            
            // 对于特定类型的错误，提供优雅降级
            if (typeof(T) == typeof((string?, int)))
            {
                // OCR操作返回类型，使用降级处理
                var fallbackResult = HandleOcrFailure(ex, operationName);
                return (T)(object)fallbackResult;
            }
            
            _logger.LogError(ex, "Error in {OperationName}: {Message}", operationName, ex.Message);
            throw;
        }
    }

    /// <summary>
    /// 强制执行操作间隔
    /// </summary>
    private async Task EnforceOperationIntervalAsync()
    {
        var now = DateTime.UtcNow;
        var timeSinceLastOperation = now - _lastOperationTime;
        
        if (timeSinceLastOperation < _minOperationInterval)
        {
            var delayTime = _minOperationInterval - timeSinceLastOperation;
            _logger.LogDebug("强制执行操作间隔，等待: {DelayMs}ms", delayTime.TotalMilliseconds);
            await Task.Delay(delayTime);
        }
        
        _lastOperationTime = DateTime.UtcNow;
    }

    /// <summary>
    /// 获取操作统计信息
    /// </summary>
    public (int Total, int Failed, int Concurrent) GetOperationStatistics()
    {
        return (_totalOperations, _failedOperations, _concurrentOperations);
    }

    /// <summary>
    /// 释放资源
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// 释放资源的具体实现
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                lock (_lock)
                {
                    _model = null;
                }
            }
            _disposed = true;
        }
    }
}