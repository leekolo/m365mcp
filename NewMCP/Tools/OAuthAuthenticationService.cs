using System.ComponentModel;
using System.Net.Http.Json;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using ModelContextProtocol.Server;

namespace NewMCP.Tools;

/// <summary>
/// OAuth 2.0 Authentication service with full auth code flow.
/// Handles per-user authentication with token management.
/// </summary>
public class OAuthAuthenticationService
{
    private readonly string? _tenantId;
    private readonly string? _clientId;
    private readonly string? _clientSecret;
    private readonly string? _redirectUri;
    private readonly string[] _defaultScopes;

    // In-memory token cache: userId -> (accessToken, refreshToken, expiresAt)
    private readonly Dictionary<string, (string AccessToken, string? RefreshToken, DateTime ExpiresAt)> _tokenCache = new();

    public OAuthAuthenticationService()
    {
        _tenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
        _clientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
        _clientSecret = Environment.GetEnvironmentVariable("AZURE_CLIENT_SECRET");
        _redirectUri = Environment.GetEnvironmentVariable("OAUTH_REDIRECT_URI") ?? "http://localhost";

        _defaultScopes = new[] { "https://graph.microsoft.com/.default" };
    }

    /// <summary>
    /// Generates OAuth 2.0 authorization URL for user login.
    /// </summary>
    [McpServerTool]
    [Description("Generates OAuth 2.0 authorization URL for user login.")]
    public string GetAuthorizationUrl(
        [Description("The user identifier for tracking (e.g., email)")] string userId,
        [Description("Roles to request: SharePoint.Only or SharePoint.And.Outlook")] string roles = "SharePoint.Only")
    {
        if (string.IsNullOrEmpty(_tenantId) || string.IsNullOrEmpty(_clientId))
        {
            return JsonSerializer.Serialize(new
            {
                error = "OAuth not configured. Set AZURE_TENANT_ID and AZURE_CLIENT_ID."
            });
        }

        var scopes = GetScopesForRole(roles);
        var scopeString = string.Join(" ", scopes);

        // Generate state for CSRF protection
        var state = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(
            JsonSerializer.Serialize(new { userId, roles, timestamp = DateTime.UtcNow })));

        var authUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/authorize" +
            $"?client_id={_clientId}" +
            $"&response_type=code" +
            $"&redirect_uri={Uri.EscapeDataString(_redirectUri ?? "http://localhost")}" +
            $"&scope={Uri.EscapeDataString(scopeString)}" +
            $"&state={state}";

        return JsonSerializer.Serialize(new
        {
            authorizationUrl = authUrl,
            state = state,
            userId = userId,
            roles = roles,
            message = "Navigate to authorization URL, then call CompleteAuthorization with the code."
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// Exchanges authorization code for access token.
    /// </summary>
    [McpServerTool]
    [Description("Exchanges OAuth authorization code for access token.")]
    public async Task<string> CompleteAuthorization(
        [Description("The authorization code from OAuth redirect")] string authCode,
        [Description("The state parameter from authorization URL")] string state)
    {
        if (string.IsNullOrEmpty(_tenantId) || string.IsNullOrEmpty(_clientId) || string.IsNullOrEmpty(_clientSecret))
        {
            return JsonSerializer.Serialize(new
            {
                error = "OAuth not configured. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET."
            });
        }

        try
        {
            // Decode state to get user info
            var stateJson = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(state));
            var stateObj = JsonSerializer.Deserialize<JsonElement>(stateJson);
            var userId = stateObj.GetProperty("userId").GetString() ?? "unknown";
            var roles = stateObj.GetProperty("roles").GetString() ?? "SharePoint.Only";

            // Exchange auth code for token
            var token = await ExchangeCodeForToken(authCode, roles);

            // Cache token
            _tokenCache[userId] = (token.AccessToken, token.RefreshToken, token.ExpiresAt);

            return JsonSerializer.Serialize(new
            {
                success = true,
                userId = userId,
                roles = roles,
                expiresAt = token.ExpiresAt.ToString("yyyy-MM-dd HH:mm:ss"),
                message = "Authorization completed. User can now access resources based on their role."
            }, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Authorization failed",
                message = ex.Message
            });
        }
    }

    /// <summary>
    /// Refreshes access token using refresh token.
    /// </summary>
    [McpServerTool]
    [Description("Refreshes access token using refresh token.")]
    public async Task<string> RefreshToken(
        [Description("The user ID to refresh token for")] string userId)
    {
        if (!_tokenCache.TryGetValue(userId, out var cached) || string.IsNullOrEmpty(cached.RefreshToken))
        {
            return JsonSerializer.Serialize(new
            {
                error = "No refresh token found for user"
            });
        }

        try
        {
            var token = await RefreshAccessToken(cached.RefreshToken);
            _tokenCache[userId] = (token.AccessToken, token.RefreshToken, token.ExpiresAt);

            return JsonSerializer.Serialize(new
            {
                success = true,
                expiresAt = token.ExpiresAt.ToString("yyyy-MM-dd HH:mm:ss"),
                message = "Token refreshed successfully."
            }, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                error = "Token refresh failed",
                message = ex.Message
            });
        }
    }

    /// <summary>
    /// Gets Graph client for a specific user.
    /// </summary>
    public GraphServiceClient GetGraphClientForUser(string userId)
    {
        if (!_tokenCache.TryGetValue(userId, out var cached))
        {
            throw new InvalidOperationException($"No token found for user {userId}. Complete authorization first.");
        }

        if (cached.ExpiresAt < DateTime.UtcNow.AddMinutes(5))
        {
            throw new InvalidOperationException($"Token expired for user {userId}. Refresh token first.");
        }

        var credential = new Azure.Core.AccessToken(cached.AccessToken, cached.ExpiresAt);
        return new GraphServiceClient(new TokenCredentialAdapter(credential));
    }

    /// <summary>
    /// Gets service-level Graph client (fallback).
    /// </summary>
    public GraphServiceClient GetGraphClient()
    {
        if (string.IsNullOrEmpty(_tenantId) || string.IsNullOrEmpty(_clientId) || string.IsNullOrEmpty(_clientSecret))
        {
            throw new InvalidOperationException("OAuth not configured.");
        }

        var credential = new ClientSecretCredential(_tenantId, _clientId, _clientSecret);
        return new GraphServiceClient(credential, _defaultScopes);
    }

    /// <summary>
    /// Checks if user has valid token.
    /// </summary>
    public bool HasValidToken(string userId)
    {
        return _tokenCache.TryGetValue(userId, out var cached) && cached.ExpiresAt > DateTime.UtcNow.AddMinutes(5);
    }

    /// <summary>
    /// Maps roles to Graph API scopes.
    /// </summary>
    public string[] GetScopesForRole(string role)
    {
        return role switch
        {
            "SharePoint.Only" => new[] { "Sites.Read.All" },
            "SharePoint.And.Outlook" => new[] { "Sites.Read.All", "Mail.Read", "Mail.Send" },
            _ => new[] { "Sites.Read.All" }
        };
    }

    /// <summary>
    /// Clears cached tokens.
    /// </summary>
    [McpServerTool]
    [Description("Clears cached authentication tokens.")]
    public string ClearTokens(
        [Description("User ID to clear token for, or 'all' to clear all")] string userId = "all")
    {
        if (userId == "all")
        {
            _tokenCache.Clear();
            return JsonSerializer.Serialize(new { success = true, message = "All tokens cleared" });
        }
        else
        {
            var removed = _tokenCache.Remove(userId);
            return JsonSerializer.Serialize(new { success = removed, message = removed ? $"Token for {userId} cleared" : $"No token found for {userId}" });
        }
    }

    /// <summary>
    /// Lists cached users.
    /// </summary>
    [McpServerTool]
    [Description("Lists authenticated users.")]
    public string ListUsers()
    {
        var users = _tokenCache.Select(kvp => new
        {
            userId = kvp.Key,
            expiresAt = kvp.Value.ExpiresAt.ToString("yyyy-MM-dd HH:mm:ss"),
            isValid = kvp.Value.ExpiresAt > DateTime.UtcNow
        }).ToList();

        return JsonSerializer.Serialize(new
        {
            users = users,
            total = users.Count
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// Checks OAuth configuration status.
    /// </summary>
    [McpServerTool]
    [Description("Checks OAuth configuration and connection status.")]
    public string CheckOAuthStatus()
    {
        if (string.IsNullOrEmpty(_tenantId) || string.IsNullOrEmpty(_clientId) || string.IsNullOrEmpty(_clientSecret))
        {
            return JsonSerializer.Serialize(new
            {
                configured = false,
                message = "OAuth not fully configured. Required: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET"
            }, new JsonSerializerOptions { WriteIndented = true });
        }

        return JsonSerializer.Serialize(new
        {
            configured = true,
            tenantId = _tenantId,
            clientId = _clientId,
            redirectUri = _redirectUri,
            cachedUsers = _tokenCache.Count,
            message = "OAuth is configured."
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    // Private helper methods

    private async Task<(string AccessToken, string? RefreshToken, DateTime ExpiresAt)> ExchangeCodeForToken(string authCode, string roles)
    {
        var scopes = GetScopesForRole(roles);
        var scopeString = string.Join(" ", scopes);

        var tokenUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";
        var content = new Dictionary<string, string>
        {
            ["client_id"] = _clientId!,
            ["scope"] = scopeString,
            ["code"] = authCode,
            ["redirect_uri"] = _redirectUri ?? "http://localhost",
            ["grant_type"] = "authorization_code",
            ["client_secret"] = _clientSecret!
        };

        using var client = new HttpClient();
        var response = await client.PostAsync(tokenUrl, new FormUrlEncodedContent(content));
        var responseJson = await response.Content.ReadFromJsonAsync<JsonElement>();

        if (responseJson.TryGetProperty("access_token", out var accessToken))
        {
            var expiresIn = responseJson.TryGetProperty("expires_in", out var exp) ? exp.GetInt32() : 3600;
            var refreshToken = responseJson.TryGetProperty("refresh_token", out var rt) ? rt.GetString() : null;

            return (accessToken.GetString()!, refreshToken, DateTime.UtcNow.AddSeconds(expiresIn - 60));
        }

        var error = responseJson.TryGetProperty("error_description", out var desc) 
            ? desc.GetString() 
            : "Token exchange failed";
        throw new InvalidOperationException(error ?? "Unknown error");
    }

    private async Task<(string AccessToken, string? RefreshToken, DateTime ExpiresAt)> RefreshAccessToken(string refreshToken)
    {
        var tokenUrl = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";
        var content = new Dictionary<string, string>
        {
            ["client_id"] = _clientId!,
            ["scope"] = "https://graph.microsoft.com/.default",
            ["refresh_token"] = refreshToken,
            ["grant_type"] = "refresh_token",
            ["client_secret"] = _clientSecret!
        };

        using var client = new HttpClient();
        var response = await client.PostAsync(tokenUrl, new FormUrlEncodedContent(content));
        var responseJson = await response.Content.ReadFromJsonAsync<JsonElement>();

        if (responseJson.TryGetProperty("access_token", out var accessToken))
        {
            var expiresIn = responseJson.TryGetProperty("expires_in", out var exp) ? exp.GetInt32() : 3600;
            var newRefreshToken = responseJson.TryGetProperty("refresh_token", out var rt) ? rt.GetString() : refreshToken;

            return (accessToken.GetString()!, newRefreshToken, DateTime.UtcNow.AddSeconds(expiresIn - 60));
        }

        throw new InvalidOperationException("Token refresh failed");
    }
}

/// <summary>
/// Adapter to use Azure Core token with Graph SDK.
/// </summary>
internal class TokenCredentialAdapter : Azure.Core.TokenCredential
{
    private readonly Azure.Core.AccessToken _token;

    public TokenCredentialAdapter(Azure.Core.AccessToken token)
    {
        _token = token;
    }

    public override Azure.Core.AccessToken GetToken(Azure.Core.TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        return _token;
    }

    public override async ValueTask<Azure.Core.AccessToken> GetTokenAsync(Azure.Core.TokenRequestContext requestContext, CancellationToken cancellationToken)
    {
        return _token;
    }
}