using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using Microsoft.Extensions.Logging;
using System.Windows.Automation.Text;

namespace WindowsMCP.Net.Services;

/// <summary>
/// UI Automation service for finding and interacting with UI elements.
/// Uses Windows built-in UI Automation framework to find elements by text, class name, automation ID.
/// </summary>
public class UIAutomationService
{
    private readonly ILogger<UIAutomationService> _logger;

    public UIAutomationService(ILogger<UIAutomationService> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Find all UI elements containing the specified text.
    /// </summary>
    /// <param name="text">Text to search for</param>
    /// <returns>List of elements found</returns>
    public List<AutomationElement> FindElementsByText(string text)
    {
        var results = new List<AutomationElement>();
        var root = AutomationElement.RootElement;
        
        FindElementsByText(root, text, results);
        _logger.LogInformation("Found {Count} elements containing text: {Text}", results.Count, text);
        
        return results;
    }

    /// <summary>
    /// Find a UI element by its class name.
    /// </summary>
    /// <param name="className">Class name to search for</param>
    /// <returns>List of elements found</returns>
    public List<AutomationElement> FindElementsByClassName(string className)
    {
        var results = new List<AutomationElement>();
        var root = AutomationElement.RootElement;
        
        FindElementsByProperty(root, AutomationElement.ClassNameProperty, className, results);
        _logger.LogInformation("Found {Count} elements with class name: {ClassName}", results.Count, className);
        
        return results;
    }

    /// <summary>
    /// Find a UI element by its automation ID.
    /// </summary>
    /// <param name="automationId">Automation ID to search for</param>
    /// <returns>List of elements found</returns>
    public List<AutomationElement> FindElementsByAutomationId(string automationId)
    {
        var results = new List<AutomationElement>();
        var root = AutomationElement.RootElement;
        
        FindElementsByProperty(root, AutomationElement.AutomationIdProperty, automationId, results);
        _logger.LogInformation("Found {Count} elements with automation ID: {AutomationId}", results.Count, automationId);
        
        return results;
    }

    /// <summary>
    /// Get the bounding rectangle of an element.
    /// </summary>
    /// <param name="element">The UI element</param>
    /// <returns>Bounding rectangle (x, y, width, height)</returns>
    public Rectangle GetElementBoundingRectangle(AutomationElement element)
    {
        try
        {
            var rect = element.Current.BoundingRectangle;
            return new Rectangle((int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height);
        }
        catch (ElementNotEnabledException ex)
        {
            _logger.LogError(ex, "Element is not enabled");
            return Rectangle.Empty;
        }
        catch (ElementNotAvailableException ex)
        {
            _logger.LogError(ex, "Element is not available");
            return Rectangle.Empty;
        }
    }

    /// <summary>
    /// Click a UI element using the Invoke pattern.
    /// </summary>
    /// <param name="element">Element to click</param>
    /// <returns>True if successful</returns>
    public bool ClickElement(AutomationElement element)
    {
        try
        {
            if (element.TryGetCurrentPattern(InvokePattern.Pattern, out object patternObj))
            {
                var invokePattern = (InvokePattern)patternObj;
                invokePattern.Invoke();
                _logger.LogInformation("Successfully clicked element");
                return true;
            }
            
            _logger.LogWarning("Element does not support Invoke pattern");
            return false;
        }
        catch (ElementNotEnabledException ex)
        {
            _logger.LogError(ex, "Element is not enabled");
            return false;
        }
        catch (ElementNotAvailableException ex)
        {
            _logger.LogError(ex, "Element is not available");
            return false;
        }
    }

    /// <summary>
    /// Wait for an element to appear on screen.
    /// </summary>
    /// <param name="property">Property to check</param>
    /// <param name="value">Value to wait for</param>
    /// <param name="timeoutMs">Timeout in milliseconds</param>
    /// <returns>The element if found, null otherwise</returns>
    public AutomationElement WaitForElement(AutomationProperty property, string value, int timeoutMs)
    {
        var startTime = Environment.TickCount;
        
        while (Environment.TickCount - startTime < timeoutMs)
        {
            var root = AutomationElement.RootElement;
            var results = new List<AutomationElement>();
            FindElementsByProperty(root, property, value, results);
            
            if (results.Count > 0)
            {
                _logger.LogInformation("Element found after {Ms}ms", Environment.TickCount - startTime);
                return results[0];
            }
            
            System.Threading.Thread.Sleep(500);
        }
        
        _logger.LogWarning("Timeout waiting for element: {Value}", value);
        return null;
    }

    /// <summary>
    /// Get all properties of an element as a dictionary.
    /// </summary>
    /// <param name="element">The UI element</param>
    /// <returns>Dictionary of property names and values</returns>
    public Dictionary<string, object> GetElementProperties(AutomationElement element)
    {
        var properties = new Dictionary<string, object>();
        
        try
        {
            var name = element.Current.Name;
            if (!string.IsNullOrEmpty(name))
                properties["name"] = name;

            var className = element.Current.ClassName;
            if (!string.IsNullOrEmpty(className))
                properties["className"] = className;

            var automationId = element.Current.AutomationId;
            if (!string.IsNullOrEmpty(automationId))
                properties["automationId"] = automationId;

            var controlType = element.Current.ControlType.ProgrammaticName;
            properties["controlType"] = controlType;

            var rect = element.Current.BoundingRectangle;
            properties["boundingRectangle"] = new { x = rect.X, y = rect.Y, width = rect.Width, height = rect.Height };

            var isEnabled = element.Current.IsEnabled;
            properties["isEnabled"] = isEnabled;

            _logger.LogDebug("Extracted {Count} properties for element", properties.Count);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting element properties");
        }

        return properties;
    }

    private void FindElementsByText(AutomationElement root, string text, List<AutomationElement> results)
    {
        var elementName = root.Current.Name;
        if (!string.IsNullOrEmpty(elementName) && elementName.Contains(text, StringComparison.OrdinalIgnoreCase))
        {
            results.Add(root);
        }

        try
        {
            foreach (AutomationElement child in root.FindAll(TreeScope.Children, Condition.TrueCondition))
            {
                FindElementsByText(child, text, results);
            }
        }
        catch (ElementNotAvailableException)
        {
            // Element disappeared during enumeration, skip
        }
    }

    private void FindElementsByProperty(AutomationElement root, AutomationProperty property, string value, List<AutomationElement> results)
    {
        try
        {
            var currentValue = root.GetCurrentPropertyValue(property) as string;
            if (!string.IsNullOrEmpty(currentValue) && currentValue.Equals(value, StringComparison.OrdinalIgnoreCase))
            {
                results.Add(root);
            }

            foreach (AutomationElement child in root.FindAll(TreeScope.Children, Condition.TrueCondition))
            {
                FindElementsByProperty(child, property, value, results);
            }
        }
        catch (ElementNotAvailableException)
        {
            // Element disappeared during enumeration, skip
        }
    }
}
