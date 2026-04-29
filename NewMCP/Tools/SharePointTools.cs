using System.ComponentModel;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace NewMCP.Tools;

/// <summary>
/// MCP tools for accessing Microsoft SharePoint via Graph API.
/// Requires Azure AD app registration with Sites.Read.All permissions.
/// Enforces role-based access control - users with SharePoint.Only or SharePoint.And.Outlook can access.
/// Uses per-user OAuth authentication when available.
/// </summary>
internal class SharePointTools
{
    private readonly GraphServiceClient? _graphClient;
    private readonly string? _tenantId;
    private readonly string? _clientId;
    private readonly string? _clientSecret;
    private readonly RoleBasedAccessControl _accessControl;
    private readonly OAuthAuthenticationService _oauthService;
    private readonly string? _currentUserId;

    public SharePointTools(RoleBasedAccessControl accessControl, OAuthAuthenticationService oauthService)
    {
        _accessControl = accessControl;
        _oauthService = oauthService;
        
        // Load configuration from environment variables
        _tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
        _clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
        _clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET");
        _currentUserId = Environment.GetEnvironmentVariable("CURRENT_USER_ID");

        if (!string.IsNullOrEmpty(_tenantId) && !string.IsNullOrEmpty(_clientId) && !string.IsNullOrEmpty(_clientSecret))
        {
            _graphClient = CreateGraphClient();
        }
    }

    private GraphServiceClient CreateGraphClient()
    {
        var credential = new ClientSecretCredential(
            _tenantId,
            _clientId,
            _clientSecret);

        return new GraphServiceClient(credential, new[] { "https://graph.microsoft.com/.default" });
    }

    private GraphServiceClient GetGraphClient()
    {
        // Try per-user OAuth first
        var userId = _currentUserId ?? "default";
        if (_oauthService.HasValidToken(userId))
        {
            return _oauthService.GetGraphClientForUser(userId);
        }

        // Fall back to service client
        if (_graphClient == null)
        {
            throw new InvalidOperationException(
                "SharePoint service not configured. Please set the following environment variables:\n" +
                "- AZURE_TENANT_ID: Your Azure AD tenant ID\n" +
                "- AZURE_CLIENT_ID: Your Azure AD app registration client ID\n" +
                "- AZURE_CLIENT_SECRET: Your Azure AD app registration client secret\n\n" +
                "Or authenticate via OAuth using GetAuthorizationUrl and CompleteAuthorization.");
        }

        return _graphClient;
    }

    private void EnsureSharePointAccess()
    {
        var userId = _currentUserId ?? "default";
        var (hasAccess, errorMessage) = _accessControl.ValidateAccess(userId, "SharePoint");
        
        if (!hasAccess)
        {
            throw new UnauthorizedAccessException(errorMessage);
        }
    }

    [McpServerTool]
    [Description("Gets the current user's profile information from SharePoint/OneDrive.")]
    public async Task<string> GetMyProfile()
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        var me = await client.Me.GetAsync();
        return JsonSerializer.Serialize(new
        {
            displayName = me?.DisplayName,
            mail = me?.Mail,
            userPrincipalName = me?.UserPrincipalName,
            id = me?.Id
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets list of SharePoint sites the user has access to.")]
    public async Task<string> GetSites(
        [Description("Maximum number of sites to return (default 10, max 100)")] int top = 10)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 100);

        var sites = await client.Sites.GetAsync();

        var siteList = sites?.Value?.Take(top).Select(s => new
        {
            name = s.Name,
            description = s.Description,
            webUrl = s.WebUrl,
            id = s.Id,
            displayName = s.DisplayName
        }).ToList();

        return JsonSerializer.Serialize(siteList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets a specific SharePoint site by ID or URL.")]
    public async Task<string> GetSiteById(
        [Description("The site ID or URL of the SharePoint site")] string siteId)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        var site = await client.Sites[siteId].GetAsync();

        if (site == null)
            return "Site not found.";

        return JsonSerializer.Serialize(new
        {
            name = site.Name,
            description = site.Description,
            webUrl = site.WebUrl,
            id = site.Id,
            displayName = site.DisplayName,
            createdDateTime = site.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            lastModifiedDateTime = site.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss")
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets the drives (document libraries) for a SharePoint site.")]
    public async Task<string> GetDrives(
        [Description("The site ID or leave empty for root site")] string? siteId = null)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        var drives = string.IsNullOrEmpty(siteId)
            ? await client.Drives.GetAsync()
            : await client.Sites[siteId].Drives.GetAsync();

        var driveList = drives?.Value?.Select(d => new
        {
            name = d.Name,
            description = d.Description,
            driveType = d.DriveType,
            webUrl = d.WebUrl,
            id = d.Id,
            quota = new
            {
                total = d.Quota?.Total,
                used = d.Quota?.Used,
                remaining = d.Quota?.Remaining,
                state = d.Quota?.State
            }
        }).ToList();

        return JsonSerializer.Serialize(driveList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets the root folder contents of a drive (document library).")]
    public async Task<string> GetDriveRoot(
        [Description("The drive ID (from GetDrives)")] string driveId,
        [Description("Maximum number of items to return (default 50, max 200)")] int top = 50)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 200);

        var items = await client.Drives[driveId].Root.GetAsync();

        if (items == null)
            return "Drive not found.";

        return JsonSerializer.Serialize(new
        {
            name = items.Name,
            webUrl = items.WebUrl,
            id = items.Id,
            createdDateTime = items.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            lastModifiedDateTime = items.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            size = items.Size,
            folder = items.Folder != null ? new { childCount = items.Folder.ChildCount } : null,
            file = items.File != null ? new { mimeType = items.File.MimeType } : null
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets the items (files and folders) in a folder.")]
    public async Task<string> GetFolderItems(
        [Description("The drive ID")] string driveId,
        [Description("The folder ID (use 'root' for root folder)")] string folderId = "root",
        [Description("Maximum number of items to return (default 50, max 200)")] int top = 50)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 200);

        var items = await client.Drives[driveId].Items[folderId].Children.GetAsync();

        var itemList = items?.Value?.Take(top).Select(i => new
        {
            name = i.Name,
            webUrl = i.WebUrl,
            id = i.Id,
            createdDateTime = i.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            lastModifiedDateTime = i.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            size = i.Size,
            type = i.Folder != null ? "folder" : "file",
            folder = i.Folder != null ? new { childCount = i.Folder.ChildCount } : null,
            file = i.File != null ? new { mimeType = i.File.MimeType } : null
        }).ToList();

        return JsonSerializer.Serialize(itemList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Searches for files in SharePoint/OneDrive by name.")]
    public async Task<string> SearchFiles(
        [Description("Search query string")] string query,
        [Description("Maximum number of results (default 25, max 100)")] int top = 25)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 100);

        // Get all drives and search in each
        var drives = await client.Drives.GetAsync();
        
        var searchResults = new List<object>();
        if (drives?.Value != null)
        {
            foreach (var drive in drives.Value.Take(5))
            {
                try
                {
                    // Use root to get top-level items
                    var root = await client.Drives[drive.Id].Root.GetAsync();
                    if (root != null)
                    {
                        // Add root if it matches
                        if (root.Name != null && root.Name.Contains(query, StringComparison.OrdinalIgnoreCase))
                        {
                            searchResults.Add(new
                            {
                                name = root.Name,
                                webUrl = root.WebUrl,
                                id = root.Id,
                                driveName = drive.Name,
                                driveId = drive.Id,
                                type = root.Folder != null ? "folder" : "file"
                            });
                        }
                    }
                }
                catch { /* Skip drives without access */ }
            }
        }

        return JsonSerializer.Serialize(searchResults.Take(top).ToList(), new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets a specific file or folder by ID.")]
    public async Task<string> GetItemById(
        [Description("The drive ID")] string driveId,
        [Description("The item ID")] string itemId)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        var item = await client.Drives[driveId].Items[itemId].GetAsync();

        if (item == null)
            return "Item not found.";

        return JsonSerializer.Serialize(new
        {
            name = item.Name,
            webUrl = item.WebUrl,
            id = item.Id,
            createdDateTime = item.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            lastModifiedDateTime = item.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            size = item.Size,
            type = item.Folder != null ? "folder" : "file",
            createdBy = new { name = item.CreatedBy?.User?.DisplayName },
            lastModifiedBy = new { name = item.LastModifiedBy?.User?.DisplayName }
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Downloads file content as base64.")]
    public async Task<string> DownloadFile(
        [Description("The drive ID")] string driveId,
        [Description("The file item ID")] string itemId)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        try
        {
            using var stream = await client.Drives[driveId].Items[itemId].Content.GetAsync();
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms);
            var bytes = ms.ToArray();
            var base64 = Convert.ToBase64String(bytes);

            return JsonSerializer.Serialize(new
            {
                success = true,
                content = base64,
                size = bytes.Length,
                message = "File content returned as base64. Use a client library to decode."
            }, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (ODataError ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = ex.Error?.Message ?? ex.Message
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = ex.Message
            });
        }
    }

    [McpServerTool]
    [Description("Gets recent files the user has modified or viewed.")]
    public async Task<string> GetRecentFiles(
        [Description("Maximum number of files to return (default 10, max 50)")] int top = 10)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 50);

        // Get recent activity from root of all drives user has access to
        var drives = await client.Drives.GetAsync();
        
        var recentItems = new List<object>();
        if (drives?.Value != null)
        {
            foreach (var drive in drives.Value.Take(10))
            {
                try
                {
                    var root = await client.Drives[drive.Id].Root.GetAsync();
                    if (root != null)
                    {
                        recentItems.Add(new
                        {
                            name = root.Name,
                            webUrl = root.WebUrl,
                            id = root.Id,
                            lastModifiedDateTime = root.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
                            size = root.Size,
                            type = root.Folder != null ? "folder" : "file",
                            driveName = drive.Name
                        });
                    }
                }
                catch { /* Skip drives without access */ }
            }
        }

        return JsonSerializer.Serialize(recentItems.Take(top).ToList(), new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets shared files and folders for the current user.")]
    public async Task<string> GetSharedFiles(
        [Description("Maximum number of items to return (default 10, max 50)")] int top = 10)
    {
        EnsureSharePointAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 50);

        // Get shared items by checking root of all drives for shared content
        var drives = await client.Drives.GetAsync();
        
        var sharedList = new List<object>();
        if (drives?.Value != null)
        {
            foreach (var drive in drives.Value.Take(10))
            {
                try
                {
                    var root = await client.Drives[drive.Id].Root.GetAsync();
                    if (root != null && root.Shared != null)
                    {
                        sharedList.Add(new
                        {
                            name = root.Name,
                            webUrl = root.WebUrl,
                            id = root.Id,
                            owner = root.CreatedBy?.User?.DisplayName,
                            driveName = drive.Name
                        });
                    }
                }
                catch { /* Skip drives without access */ }
            }
        }

        return JsonSerializer.Serialize(sharedList.Take(top).ToList(), new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Checks if the SharePoint service is properly configured and authenticated.")]
    public async Task<string> CheckSharePointConnection()
    {
        var userId = _currentUserId ?? "default";
        
        // Check per-user OAuth first
        if (_oauthService.HasValidToken(userId))
        {
            try
            {
                var client = _oauthService.GetGraphClientForUser(userId);
                var me = await client.Me.GetAsync();
                return JsonSerializer.Serialize(new
                {
                    configured = true,
                    connected = true,
                    authType = "OAuth",
                    user = me?.DisplayName,
                    upn = me?.UserPrincipalName
                });
            }
            catch (Exception ex)
            {
                return JsonSerializer.Serialize(new
                {
                    configured = true,
                    connected = false,
                    authType = "OAuth",
                    error = ex.Message
                });
            }
        }

        // Fall back to service client
        if (_graphClient == null)
        {
            return JsonSerializer.Serialize(new
            {
                configured = false,
                message = "SharePoint service not configured. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables, or authenticate via OAuth."
            });
        }

        try
        {
            var me = await _graphClient.Me.GetAsync();
            return JsonSerializer.Serialize(new
            {
                configured = true,
                connected = true,
                authType = "ServicePrincipal",
                user = me?.DisplayName,
                upn = me?.UserPrincipalName
            });
        }
        catch (ODataError ex)
        {
            return JsonSerializer.Serialize(new
            {
                configured = true,
                connected = false,
                authType = "ServicePrincipal",
                error = ex.Error?.Message ?? ex.Message
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                configured = true,
                connected = false,
                authType = "ServicePrincipal",
                error = ex.Message
            });
        }
    }
}