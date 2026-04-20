using Microsoft.Extensions.Logging;
using ModelContextProtocol.Server;
using System.ComponentModel;
using Interface;

namespace Tools.Desktop;

/// <summary>
/// MCP tool for opening URLs in the default browser.
/// </summary>
[McpServerToolType]
public class OpenBrowserTool
{
    private readonly IDesktopService _desktopService;
    private readonly ILogger<OpenBrowserTool> _logger;

    public OpenBrowserTool(IDesktopService desktopService, ILogger<OpenBrowserTool> logger)
    {
        _desktopService = desktopService;
        _logger = logger;
    }

    /// <summary>
    /// Open a URL in the default browser.
    /// </summary>
    /// <param name="url">The URL to open. URL is required and must be valid</param>
    /// <param name="searchQuery">Optional search query to use Google search</param>
    /// <returns>Result message indicating success or failure</returns>
    [McpServerTool, Description("Open a URL in the default browser")]
    public async Task<string> OpenBrowserAsync(
        [Description("The URL to open (required, must be valid HTTP/HTTPS URL)")] string? url = null,
        [Description("Optional search query to use Google search")] string? searchQuery = null)
    {
        _logger.LogInformation("Opening browser with URL: {Url}, SearchQuery: {SearchQuery}", url ?? "default", searchQuery ?? "none");
        
        return await _desktopService.OpenBrowserAsync(url, searchQuery);
    }
}