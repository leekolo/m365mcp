using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace NewMCP.Tools;

/// <summary>
/// Role-based access control service for MCP tools.
/// Enforces per-user access control based on Entra App Roles.
/// </summary>
public class RoleBasedAccessControl
{
    // App role definitions
    public const string RoleSharePointOnly = "SharePoint.Only";
    public const string RoleSharePointAndOutlook = "SharePoint.And.Outlook";

    // Role to tool mapping
    private static readonly Dictionary<string, string[]> RoleToolPermissions = new()
    {
        { RoleSharePointOnly, new[] { "SharePoint" } },
        { RoleSharePointAndOutlook, new[] { "SharePoint", "Email" } }
    };

    private readonly string? _roleMappingJson;
    private Dictionary<string, string>? _roleMapping;

    public RoleBasedAccessControl()
    {
        _roleMappingJson = Environment.GetEnvironmentVariable("USER_ROLE_MAPPING");
        LoadRoleMapping();
    }

    /// <summary>
    /// Loads role mapping from environment variable.
    /// Format: {"user1@contoso.com": "SharePoint.Only", "user2@contoso.com": "SharePoint.And.Outlook"}
    /// </summary>
    private void LoadRoleMapping()
    {
        if (!string.IsNullOrEmpty(_roleMappingJson))
        {
            try
            {
                _roleMapping = JsonSerializer.Deserialize<Dictionary<string, string>>(_roleMappingJson);
            }
            catch
            {
                _roleMapping = null;
            }
        }
    }

    /// <summary>
    /// Gets the role for a specific user.
    /// </summary>
    public string GetUserRole(string userId)
    {
        if (_roleMapping != null && _roleMapping.TryGetValue(userId, out var role))
        {
            return role;
        }
        // Default role if not configured
        return RoleSharePointOnly;
    }

    /// <summary>
    /// Checks if a user has access to a specific tool category.
    /// </summary>
    public bool HasAccess(string userId, string toolCategory)
    {
        var userRole = GetUserRole(userId);
        
        if (RoleToolPermissions.TryGetValue(userRole, out var allowedTools))
        {
            return allowedTools.Contains(toolCategory, StringComparer.OrdinalIgnoreCase);
        }
        
        return false;
    }

    /// <summary>
    /// Gets the list of tools a user has access to.
    /// </summary>
    public string[] GetAllowedTools(string userId)
    {
        var userRole = GetUserRole(userId);
        
        if (RoleToolPermissions.TryGetValue(userRole, out var allowedTools))
        {
            return allowedTools;
        }
        
        return Array.Empty<string>();
    }

    /// <summary>
    /// Validates access and returns error message if access denied.
    /// </summary>
    public (bool HasAccess, string ErrorMessage) ValidateAccess(string userId, string toolCategory)
    {
        var hasAccess = HasAccess(userId, toolCategory);
        
        if (!hasAccess)
        {
            var userRole = GetUserRole(userId);
            var allowedTools = GetAllowedTools(userId);
            var allowedList = string.Join(", ", allowedTools);
            
            return (false, $"Access denied. User '{userId}' has role '{userRole}' with access to: {allowedList}. '{toolCategory}' tools are not permitted.");
        }
        
        return (true, string.Empty);
    }

    /// <summary>
    /// Gets all configured users and their roles.
    /// </summary>
    public string GetAllUserRoles()
    {
        if (_roleMapping == null || _roleMapping.Count == 0)
        {
            return JsonSerializer.Serialize(new
            {
                configured = false,
                message = "No role mapping configured. Set USER_ROLE_MAPPING environment variable.",
                defaultRole = RoleSharePointOnly,
                example = JsonSerializer.Serialize(new
                {
                    user1 = "SharePoint.Only",
                    user2 = "SharePoint.And.Outlook"
                })
            }, new JsonSerializerOptions { WriteIndented = true });
        }

        return JsonSerializer.Serialize(new
        {
            configured = true,
            users = _roleMapping,
            totalUsers = _roleMapping.Count
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// Updates role mapping (requires restart or cache refresh in production).
    /// </summary>
    [McpServerTool]
    [Description("Gets the current role assignments for all users.")]
    public string GetRoleAssignments()
    {
        return GetAllUserRoles();
    }

    /// <summary>
    /// Checks if a specific user has access to a tool category.
    /// </summary>
    [McpServerTool]
    [Description("Checks if a user has access to a specific tool category.")]
    public string CheckUserAccess(
        [Description("The user ID or email to check")] string userId,
        [Description("Tool category to check: SharePoint or Email")] string toolCategory)
    {
        var hasAccess = HasAccess(userId, toolCategory);
        var userRole = GetUserRole(userId);
        var allowedTools = GetAllowedTools(userId);

        return JsonSerializer.Serialize(new
        {
            userId = userId,
            toolCategory = toolCategory,
            hasAccess = hasAccess,
            userRole = userRole,
            allowedTools = allowedTools
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// Gets the access summary for a user.
    /// </summary>
    [McpServerTool]
    [Description("Gets access summary for a user.")]
    public string GetUserAccessSummary(
        [Description("The user ID or email to check")] string userId)
    {
        var userRole = GetUserRole(userId);
        var allowedTools = GetAllowedTools(userId);

        return JsonSerializer.Serialize(new
        {
            userId = userId,
            role = userRole,
            allowedToolCategories = allowedTools,
            canAccessSharePoint = allowedTools.Contains("SharePoint"),
            canAccessEmail = allowedTools.Contains("Email"),
            permissions = new
            {
                sharePoint = new
                {
                    sitesRead = userRole == RoleSharePointOnly || userRole == RoleSharePointAndOutlook,
                    sitesWrite = userRole == RoleSharePointAndOutlook
                },
                email = new
                {
                    mailRead = userRole == RoleSharePointAndOutlook,
                    mailSend = userRole == RoleSharePointAndOutlook
                }
            }
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// Lists all available roles.
    /// </summary>
    [McpServerTool]
    [Description("Lists all available roles and their permissions.")]
    public string ListRoles()
    {
        return JsonSerializer.Serialize(new
        {
            roles = new
            {
                SharePoint_Only = new
                {
                    description = "Access to SharePoint/OneDrive only",
                    tools = new[] { "SharePoint" },
                    scopes = new[] { "Sites.Read.All" }
                },
                SharePoint_And_Outlook = new
                {
                    description = "Access to SharePoint/OneDrive and Outlook email",
                    tools = new[] { "SharePoint", "Email" },
                    scopes = new[] { "Sites.Read.All", "Mail.Read", "Mail.Send" }
                }
            }
        }, new JsonSerializerOptions { WriteIndented = true });
    }
}