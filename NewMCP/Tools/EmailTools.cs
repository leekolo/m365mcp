using System.ComponentModel;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;

namespace NewMCP.Tools;

/// <summary>
/// MCP tools for accessing Microsoft 365 email via Graph API.
/// Requires Azure AD app registration with Mail.Read and Mail.Send permissions.
/// Enforces role-based access control - only users with SharePoint.And.Outlook role can access email.
/// Uses per-user OAuth authentication when available.
/// </summary>
internal class EmailTools
{
    private readonly GraphServiceClient? _graphClient;
    private readonly string? _tenantId;
    private readonly string? _clientId;
    private readonly string? _clientSecret;
    private readonly RoleBasedAccessControl _accessControl;
    private readonly OAuthAuthenticationService _oauthService;
    private readonly string? _currentUserId;

    public EmailTools(RoleBasedAccessControl accessControl, OAuthAuthenticationService oauthService)
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
                "Email service not configured. Please set the following environment variables:\n" +
                "- AZURE_TENANT_ID: Your Azure AD tenant ID\n" +
                "- AZURE_CLIENT_ID: Your Azure AD app registration client ID\n" +
                "- AZURE_CLIENT_SECRET: Your Azure AD app registration client secret\n\n" +
                "Or authenticate via OAuth using GetAuthorizationUrl and CompleteAuthorization.");
        }

        return _graphClient;
    }

    private void EnsureEmailAccess()
    {
        var userId = _currentUserId ?? "default";
        var (hasAccess, errorMessage) = _accessControl.ValidateAccess(userId, "Email");
        
        if (!hasAccess)
        {
            throw new UnauthorizedAccessException(errorMessage);
        }
    }

    [McpServerTool]
    [Description("Gets the current user's profile information from Microsoft 365.")]
    public async Task<string> GetMyProfile()
    {
        EnsureEmailAccess();
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
    [Description("Gets recent emails from the user's inbox. Default is 10, max is 50.")]
    public async Task<string> GetRecentEmails(
        [Description("Number of emails to retrieve (default 10, max 50)")] int top = 10)
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 50);

        var messages = await client.Me.Messages.GetAsync();

        var emailList = messages?.Value?.Take(top).Select(m => new
        {
            subject = m.Subject,
            from = m.From?.EmailAddress?.Name ?? m.From?.EmailAddress?.Address,
            to = m.ToRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
            received = m.ReceivedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            isRead = m.IsRead,
            bodyPreview = m.BodyPreview,
            id = m.Id
        }).ToList();

        return JsonSerializer.Serialize(emailList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets a specific email by its ID.")]
    public async Task<string> GetEmailById(
        [Description("The ID of the email to retrieve")] string messageId)
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        var message = await client.Me.Messages[messageId].GetAsync();

        if (message == null)
            return "Email not found.";

        return JsonSerializer.Serialize(new
        {
            subject = message.Subject,
            from = new { name = message.From?.EmailAddress?.Name, address = message.From?.EmailAddress?.Address },
            to = message.ToRecipients?.Select(r => new { name = r.EmailAddress?.Name, address = r.EmailAddress?.Address }),
            cc = message.CcRecipients?.Select(r => new { name = r.EmailAddress?.Name, address = r.EmailAddress?.Address }),
            received = message.ReceivedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            sent = message.SentDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            body = message.Body?.Content,
            bodyPreview = message.BodyPreview,
            isRead = message.IsRead,
            importance = message.Importance,
            hasAttachments = message.HasAttachments,
            id = message.Id
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Sends a new email message.")]
    public async Task<string> SendEmail(
        [Description("Recipients (comma-separated email addresses)")] string to,
        [Description("Subject of the email")] string subject,
        [Description("Body of the email (plain text)")] string body,
        [Description("CC recipients (comma-separated, optional)")] string? cc = null,
        [Description("Importance: low, normal, or high")] string importance = "normal")
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        var toRecipients = to.Split(',', StringSplitOptions.RemoveEmptyEntries)
            .Select(e => new Recipient
            {
                EmailAddress = new EmailAddress { Address = e.Trim() }
            }).ToList();

        List<Recipient>? ccRecipientsList = null;
        if (!string.IsNullOrEmpty(cc))
        {
            ccRecipientsList = cc.Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => new Recipient
                {
                    EmailAddress = new EmailAddress { Address = e.Trim() }
                }).ToList();
        }

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Text,
                Content = body
            },
            ToRecipients = toRecipients,
            CcRecipients = ccRecipientsList,
            Importance = importance?.ToLowerInvariant() switch
            {
                "low" => Importance.Low,
                "high" => Importance.High,
                _ => Importance.Normal
            }
        };

        // Use the SendMail POST request with proper body
        var requestBody = new Microsoft.Graph.Me.SendMail.SendMailPostRequestBody
        {
            Message = message
        };
        await client.Me.SendMail.PostAsync(requestBody);

        return JsonSerializer.Serialize(new { success = true, message = "Email sent successfully." });
    }

    [McpServerTool]
    [Description("Marks an email as read or unread.")]
    public async Task<string> MarkEmailAsRead(
        [Description("The ID of the email to update")] string messageId,
        [Description("True to mark as read, false to mark as unread")] bool isRead)
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        var message = new Message
        {
            IsRead = isRead
        };

        await client.Me.Messages[messageId].PatchAsync(message);

        return JsonSerializer.Serialize(new { success = true, message = $"Email marked as {(isRead ? "read" : "unread")}." });
    }

    [McpServerTool]
    [Description("Searches for emails matching a query string.")]
    public async Task<string> SearchEmails(
        [Description("Search query (searches in subject and body)")] string query,
        [Description("Maximum number of results (default 25, max 100)")] int top = 25)
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 100);

        // Sanitize query to prevent OData injection
        var safeQuery = query.Replace("'", "''");

        var messages = await client.Me.Messages
            .GetAsync(o => o.QueryParameters.Filter = $"contains(subject, '{safeQuery}') or contains(bodyPreview, '{safeQuery}')");

        var emailList = messages?.Value?.Take(top).Select(m => new
        {
            subject = m.Subject,
            from = m.From?.EmailAddress?.Name ?? m.From?.EmailAddress?.Address,
            to = m.ToRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
            received = m.ReceivedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            isRead = m.IsRead,
            bodyPreview = m.BodyPreview,
            id = m.Id
        }).ToList();

        return JsonSerializer.Serialize(emailList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets the user's mailbox folders.")]
    public async Task<string> GetMailFolders()
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        var folders = await client.Me.MailFolders.GetAsync();

        var folderList = folders?.Value?.Select(f => new
        {
            name = f.DisplayName,
            childFolders = f.ChildFolderCount,
            unread = f.UnreadItemCount,
            total = f.TotalItemCount,
            id = f.Id
        }).ToList();

        return JsonSerializer.Serialize(folderList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Gets emails from a specific mailbox folder.")]
    public async Task<string> GetEmailsFromFolder(
        [Description("The folder ID (use GetMailFolders to find IDs)")] string folderId,
        [Description("Number of emails to retrieve (default 10, max 50)")] int top = 10)
    {
        EnsureEmailAccess();
        var client = GetGraphClient();

        top = Math.Min(Math.Max(top, 1), 50);

        var messages = await client.Me.MailFolders[folderId].Messages.GetAsync();

        var emailList = messages?.Value?.Take(top).Select(m => new
        {
            subject = m.Subject,
            from = m.From?.EmailAddress?.Name ?? m.From?.EmailAddress?.Address,
            to = m.ToRecipients?.Select(r => r.EmailAddress?.Address).ToList(),
            received = m.ReceivedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            isRead = m.IsRead,
            bodyPreview = m.BodyPreview,
            id = m.Id
        }).ToList();

        return JsonSerializer.Serialize(emailList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [Description("Checks if the email service is properly configured and authenticated.")]
    public async Task<string> CheckEmailConnection()
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
                message = "Email service not configured. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables, or authenticate via OAuth."
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