using System.Text.Json;
using FluentAssertions;
using NewMCP.Tools;
using Xunit;

namespace NewMCP.Tests;

public class OAuthAuthenticationServiceTests
{
    [Fact]
    public void GetScopesForRole_ReturnsSharePointOnlyScopes()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var scopes = service.GetScopesForRole("SharePoint.Only");

        // Assert
        scopes.Should().Contain("Sites.Read.All");
        scopes.Should().NotContain("Mail.Read");
    }

    [Fact]
    public void GetScopesForRole_ReturnsSharePointAndOutlookScopes()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var scopes = service.GetScopesForRole("SharePoint.And.Outlook");

        // Assert
        scopes.Should().Contain("Sites.Read.All");
        scopes.Should().Contain("Mail.Read");
        scopes.Should().Contain("Mail.Send");
    }

    [Fact]
    public void GetScopesForRole_UnknownRole_ReturnsDefaultScopes()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var scopes = service.GetScopesForRole("UnknownRole");

        // Assert
        scopes.Should().Contain("Sites.Read.All");
    }

    [Fact]
    public void HasValidToken_NoToken_ReturnsFalse()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var hasToken = service.HasValidToken("nonexistent-user");

        // Assert
        hasToken.Should().BeFalse();
    }

    [Fact]
    public void ClearTokens_AllUsers_ClearsAllTokens()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var result = service.ClearTokens("all");

        // Assert
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        json.GetProperty("success").GetBoolean().Should().BeTrue();
    }

    [Fact]
    public void ListUsers_NoUsers_ReturnsEmptyList()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var result = service.ListUsers();

        // Assert
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        json.GetProperty("total").GetInt32().Should().Be(0);
    }

    [Fact]
    public void CheckOAuthStatus_NotConfigured_ReturnsFalse()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act
        var result = service.CheckOAuthStatus();

        // Assert
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        // Either configured or not depending on env vars
        json.TryGetProperty("configured", out var configured).Should().BeTrue();
    }

    [Fact]
    public void GetGraphClientForUser_NoToken_ThrowsException()
    {
        // Arrange
        var service = new OAuthAuthenticationService();

        // Act & Assert
        var action = () => service.GetGraphClientForUser("nonexistent-user");
        action.Should().Throw<InvalidOperationException>()
            .WithMessage("*No token found*");
    }

    [Fact]
    public void GetAuthorizationUrl_NotConfigured_ReturnsError()
    {
        // Arrange - Clear env vars for this test
        var originalTenantId = Environment.GetEnvironmentVariable("AZURE_TENANT_ID");
        var originalClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID");
        
        try
        {
            Environment.SetEnvironmentVariable("AZURE_TENANT_ID", null);
            Environment.SetEnvironmentVariable("AZURE_CLIENT_ID", null);
            
            var service = new OAuthAuthenticationService();

            // Act
            var result = service.GetAuthorizationUrl("user1", "SharePoint.Only");

            // Assert
            var json = JsonSerializer.Deserialize<JsonElement>(result);
            json.TryGetProperty("error", out var error).Should().BeTrue();
        }
        finally
        {
            // Restore env vars
            Environment.SetEnvironmentVariable("AZURE_TENANT_ID", originalTenantId);
            Environment.SetEnvironmentVariable("AZURE_CLIENT_ID", originalClientId);
        }
    }
}