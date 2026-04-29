# HTTP MCP Server Rearchitecture Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rearchitect newMCP from a stdio-based local process to a remote ASP.NET Core HTTP server hosting the MCP SDK Streamable HTTP transport, with per-user Entra ID Bearer token auth, OBO-based delegated Graph API access, and role-based tool visibility.

**Architecture:** `WebApplication` hosts `app.MapMcp("/mcp")` behind JWT Bearer middleware and a custom `McpAuthMiddleware` that runs OBO to exchange the user's Entra token for a delegated Graph token cached per-user in `IDistributedCache`. A `McpToolFilterMiddleware` intercepts `tools/list` responses and strips Email tools for SharePoint-only users.

**Tech Stack:** .NET 10, ASP.NET Core, `ModelContextProtocol.AspNetCore 1.2.0`, `Microsoft.AspNetCore.Authentication.JwtBearer`, `Azure.Identity.OnBehalfOfCredential`, `Microsoft.Graph 5.40.0`, `Microsoft.Extensions.Caching.StackExchangeRedis`, xUnit, FluentAssertions, Moq, `Microsoft.AspNetCore.Mvc.Testing`

---

## File Map

| Action | File | Responsibility |
|--------|------|----------------|
| Modify | `NewMCP/NewMCP.csproj` | Switch to Web SDK, add AspNetCore MCP + Redis packages |
| Modify | `NewMCP.Tests/NewMCP.Tests.csproj` | Add `Microsoft.AspNetCore.Mvc.Testing` |
| Create | `NewMCP/Auth/EntraOptions.cs` | Strongly-typed config bound from `"Entra"` section |
| Create | `NewMCP/Auth/UserSession.cs` | Cached per-user Graph token record |
| Create | `NewMCP/Auth/TokenCacheService.cs` | `IDistributedCache` wrapper keyed by `mcp:session:{userId}` |
| Create | `NewMCP/Auth/GraphTokenService.cs` | OBO flow via `OnBehalfOfCredential`; `IGraphTokenService` interface |
| Create | `NewMCP/Middleware/McpAuthMiddleware.cs` | Post-JWT enrichment: OBO, cache, 5-min refresh |
| Create | `NewMCP/Graph/GraphClientFactory.cs` | Builds `GraphServiceClient` from cached token for current user |
| Create | `NewMCP/Controllers/OAuthDiscoveryController.cs` | GET `/.well-known/oauth-authorization-server` |
| Modify | `NewMCP/Tools/RoleBasedAccessControl.cs` | Remove `[McpServerTool]` attrs, `USER_ROLE_MAPPING`; new `HasAccess(role, category)` API |
| Create | `NewMCP/Graph/GraphRetryHandler.cs` | `DelegatingHandler` for HTTP 429 / Retry-After |
| Create | `NewMCP/Tools/ToolCategoryAttribute.cs` | `[ToolCategory("Email")]` attribute |
| Create | `NewMCP/Tools/ToolCategoryRegistry.cs` | Maps tool name → category |
| Create | `NewMCP/Middleware/McpToolFilterMiddleware.cs` | Buffers POST /mcp, strips tools by role |
| Modify | `NewMCP/Tools/SharePointTools.cs` | Inject `GraphClientFactory`, fix SearchFiles, add CancellationToken, PageIterator |
| Modify | `NewMCP/Tools/EmailTools.cs` | Inject `GraphClientFactory`, add CancellationToken |
| Delete | `NewMCP/Tools/RandomNumberTools.cs` | Remove dev scaffold |
| Delete | `NewMCP/Services/OAuthAuthenticationService.cs` | Replaced by McpAuthMiddleware + GraphTokenService |
| Delete | `NewMCP.Tests/Services/OAuthAuthenticationServiceTests.cs` | Tests for deleted service |
| Rewrite | `NewMCP/Program.cs` | `WebApplication` with full middleware pipeline |
| Create | `NewMCP/appsettings.json` | Placeholder Entra config (no real secrets) |
| Create | `NewMCP.Tests/Integration/McpEndpointTests.cs` | `WebApplicationFactory<Program>` integration tests |

---

### Task 1: Switch to Web SDK and add required packages

**Files:**
- Modify: `NewMCP/NewMCP.csproj`
- Modify: `NewMCP.Tests/NewMCP.Tests.csproj`

- [ ] **Step 1: Read current csproj files to understand baseline**

Read `NewMCP/NewMCP.csproj` and `NewMCP.Tests/NewMCP.Tests.csproj` to confirm current state before editing.

- [ ] **Step 2: Update `NewMCP/NewMCP.csproj`**

Replace the entire file content with:

```xml
<Project Sdk="Microsoft.NET.Sdk.Web">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>

  <ItemGroup>
    <None Include="README.md" Pack="true" PackagePath="/" />
  </ItemGroup>

  <ItemGroup>
    <PackageReference Include="ModelContextProtocol" Version="1.2.0" />
    <PackageReference Include="ModelContextProtocol.AspNetCore" Version="1.2.0" />
    <PackageReference Include="Microsoft.Graph" Version="5.40.0" />
    <PackageReference Include="Azure.Identity" Version="1.13.1" />
    <PackageReference Include="Microsoft.AspNetCore.Authentication.JwtBearer" Version="10.0.0" />
    <PackageReference Include="Microsoft.Extensions.Caching.StackExchangeRedis" Version="9.0.4" />
  </ItemGroup>

  <ItemGroup>
    <Folder Include="docs\" />
  </ItemGroup>

</Project>
```

- [ ] **Step 3: Update `NewMCP.Tests/NewMCP.Tests.csproj`**

Add `Microsoft.AspNetCore.Mvc.Testing` to the test project:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <IsPackable>false</IsPackable>
    <IsTestProject>true</IsTestProject>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2">
      <IncludeAssets>runtime; build; native; contentfiles; analyzers; buildtransitive</IncludeAssets>
      <PrivateAssets>all</PrivateAssets>
    </PackageReference>
    <PackageReference Include="Moq" Version="4.20.72" />
    <PackageReference Include="FluentAssertions" Version="6.12.2" />
    <PackageReference Include="Microsoft.AspNetCore.Mvc.Testing" Version="10.0.0" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\NewMCP\NewMCP.csproj" />
  </ItemGroup>

</Project>
```

- [ ] **Step 4: Restore packages and verify build**

Run: `dotnet restore NewMCP/NewMCP.csproj && dotnet build NewMCP/NewMCP.csproj --no-restore`
Expected: Build succeeds (may warn about missing `Program.cs` entry point changes — that's OK for now)

- [ ] **Step 5: Commit**

```bash
git add NewMCP/NewMCP.csproj NewMCP.Tests/NewMCP.Tests.csproj
git commit -m "chore: switch to Web SDK and add HTTP MCP + Redis packages"
```

---

### Task 2: Create auth record types — `EntraOptions` and `UserSession`

**Files:**
- Create: `NewMCP/Auth/EntraOptions.cs`
- Create: `NewMCP/Auth/UserSession.cs`
- Test: `NewMCP.Tests/Auth/EntraOptionsTests.cs`

- [ ] **Step 1: Write the failing test**

Create `NewMCP.Tests/Auth/EntraOptionsTests.cs`:

```csharp
using FluentAssertions;
using NewMCP.Auth;

namespace NewMCP.Tests.Auth;

public sealed class EntraOptionsTests
{
    [Fact]
    public void EntraOptions_HasSectionName_Constant()
    {
        EntraOptions.SectionName.Should().Be("Entra");
    }

    [Fact]
    public void UserSession_StoresExpectedProperties()
    {
        var session = new UserSession(
            UserId: "user-1",
            Role: "SharePoint.Only",
            GraphAccessToken: "token-abc",
            GraphRefreshToken: "refresh-xyz",
            ExpiresAt: DateTimeOffset.UtcNow.AddHours(1));

        session.UserId.Should().Be("user-1");
        session.Role.Should().Be("SharePoint.Only");
        session.GraphAccessToken.Should().Be("token-abc");
        session.IsExpiringSoon(TimeSpan.FromMinutes(5)).Should().BeFalse();
    }

    [Fact]
    public void UserSession_IsExpiringSoon_ReturnsTrueWithinWindow()
    {
        var session = new UserSession(
            UserId: "user-1",
            Role: "SharePoint.Only",
            GraphAccessToken: "token-abc",
            GraphRefreshToken: "refresh-xyz",
            ExpiresAt: DateTimeOffset.UtcNow.AddMinutes(3));

        session.IsExpiringSoon(TimeSpan.FromMinutes(5)).Should().BeTrue();
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~EntraOptionsTests" -v minimal`
Expected: FAIL — `NewMCP.Auth` namespace does not exist yet

- [ ] **Step 3: Create `NewMCP/Auth/EntraOptions.cs`**

```csharp
namespace NewMCP.Auth;

public sealed class EntraOptions
{
    public const string SectionName = "Entra";
    public required string TenantId { get; init; }
    public required string ClientId { get; init; }
    public required string ClientSecret { get; init; }
    public string Instance { get; init; } = "https://login.microsoftonline.com/";
}
```

- [ ] **Step 4: Create `NewMCP/Auth/UserSession.cs`**

```csharp
namespace NewMCP.Auth;

public sealed record UserSession(
    string UserId,
    string Role,
    string GraphAccessToken,
    string? GraphRefreshToken,
    DateTimeOffset ExpiresAt)
{
    public bool IsExpiringSoon(TimeSpan window) =>
        DateTimeOffset.UtcNow.Add(window) >= ExpiresAt;
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~EntraOptionsTests" -v minimal`
Expected: PASS — 3 tests

- [ ] **Step 6: Commit**

```bash
git add NewMCP/Auth/EntraOptions.cs NewMCP/Auth/UserSession.cs NewMCP.Tests/Auth/EntraOptionsTests.cs
git commit -m "feat: add EntraOptions and UserSession auth record types"
```

---

### Task 3: Create `TokenCacheService`

**Files:**
- Create: `NewMCP/Auth/TokenCacheService.cs`
- Test: `NewMCP.Tests/Auth/TokenCacheServiceTests.cs`

- [ ] **Step 1: Write the failing test**

Create `NewMCP.Tests/Auth/TokenCacheServiceTests.cs`:

```csharp
using FluentAssertions;
using Microsoft.Extensions.Caching.Distributed;
using Moq;
using NewMCP.Auth;
using System.Text.Json;

namespace NewMCP.Tests.Auth;

public sealed class TokenCacheServiceTests
{
    private readonly Mock<IDistributedCache> _cacheMock = new();
    private readonly TokenCacheService _sut;

    public TokenCacheServiceTests()
    {
        _sut = new TokenCacheService(_cacheMock.Object);
    }

    [Fact]
    public async Task GetAsync_ReturnsNull_WhenCacheMisses()
    {
        _cacheMock.Setup(c => c.GetAsync("mcp:session:user-1", default))
                  .ReturnsAsync((byte[]?)null);

        var result = await _sut.GetAsync("user-1");

        result.Should().BeNull();
    }

    [Fact]
    public async Task GetAsync_DeserializesSession_WhenCacheHits()
    {
        var session = new UserSession("user-1", "SharePoint.Only", "tok", null, DateTimeOffset.UtcNow.AddHours(1));
        var json = JsonSerializer.SerializeToUtf8Bytes(session);

        _cacheMock.Setup(c => c.GetAsync("mcp:session:user-1", default))
                  .ReturnsAsync(json);

        var result = await _sut.GetAsync("user-1");

        result.Should().NotBeNull();
        result!.UserId.Should().Be("user-1");
    }

    [Fact]
    public async Task SetAsync_StoresSerializedSession()
    {
        var session = new UserSession("user-1", "SharePoint.Only", "tok", null, DateTimeOffset.UtcNow.AddHours(1));
        byte[]? stored = null;

        _cacheMock.Setup(c => c.SetAsync(
            "mcp:session:user-1",
            It.IsAny<byte[]>(),
            It.IsAny<DistributedCacheEntryOptions>(),
            default))
            .Callback<string, byte[], DistributedCacheEntryOptions, CancellationToken>(
                (_, bytes, _, _) => stored = bytes)
            .Returns(Task.CompletedTask);

        await _sut.SetAsync(session);

        stored.Should().NotBeNull();
        var deserialized = JsonSerializer.Deserialize<UserSession>(stored!);
        deserialized!.UserId.Should().Be("user-1");
    }

    [Fact]
    public async Task RemoveAsync_CallsCacheRemove()
    {
        await _sut.RemoveAsync("user-1");

        _cacheMock.Verify(c => c.RemoveAsync("mcp:session:user-1", default), Times.Once);
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~TokenCacheServiceTests" -v minimal`
Expected: FAIL — `TokenCacheService` does not exist

- [ ] **Step 3: Create `NewMCP/Auth/TokenCacheService.cs`**

```csharp
using Microsoft.Extensions.Caching.Distributed;
using System.Text.Json;

namespace NewMCP.Auth;

public sealed class TokenCacheService(IDistributedCache cache)
{
    private static string Key(string userId) => $"mcp:session:{userId}";

    public async Task<UserSession?> GetAsync(string userId, CancellationToken cancellationToken = default)
    {
        var bytes = await cache.GetAsync(Key(userId), cancellationToken);
        return bytes is null ? null : JsonSerializer.Deserialize<UserSession>(bytes);
    }

    public Task SetAsync(UserSession session, CancellationToken cancellationToken = default)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(session);
        var options = new DistributedCacheEntryOptions
        {
            AbsoluteExpiration = session.ExpiresAt
        };
        return cache.SetAsync(Key(session.UserId), bytes, options, cancellationToken);
    }

    public Task RemoveAsync(string userId, CancellationToken cancellationToken = default) =>
        cache.RemoveAsync(Key(userId), cancellationToken);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~TokenCacheServiceTests" -v minimal`
Expected: PASS — 4 tests

- [ ] **Step 5: Commit**

```bash
git add NewMCP/Auth/TokenCacheService.cs NewMCP.Tests/Auth/TokenCacheServiceTests.cs
git commit -m "feat: add TokenCacheService wrapping IDistributedCache"
```

---

### Task 4: Create `GraphTokenService` (OBO flow)

**Files:**
- Create: `NewMCP/Auth/GraphTokenService.cs`
- Test: `NewMCP.Tests/Auth/GraphTokenServiceTests.cs`

- [ ] **Step 1: Write the failing test**

Create `NewMCP.Tests/Auth/GraphTokenServiceTests.cs`:

```csharp
using FluentAssertions;
using NewMCP.Auth;

namespace NewMCP.Tests.Auth;

public sealed class GraphTokenServiceTests
{
    [Theory]
    [InlineData("SharePoint.Only", new[] { "https://graph.microsoft.com/Sites.Read.All" })]
    [InlineData("SharePoint.And.Outlook", new[] {
        "https://graph.microsoft.com/Sites.Read.All",
        "https://graph.microsoft.com/Mail.Read",
        "https://graph.microsoft.com/Mail.Send"
    })]
    public void GetScopesForRole_ReturnsCorrectScopes(string role, string[] expectedScopes)
    {
        var scopes = GraphTokenService.GetScopesForRole(role);

        scopes.Should().BeEquivalentTo(expectedScopes);
    }

    [Fact]
    public void GetScopesForRole_Throws_ForUnknownRole()
    {
        var act = () => GraphTokenService.GetScopesForRole("Unknown.Role");

        act.Should().Throw<ArgumentException>()
           .WithMessage("*Unknown.Role*");
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~GraphTokenServiceTests" -v minimal`
Expected: FAIL — `GraphTokenService` does not exist

- [ ] **Step 3: Create `NewMCP/Auth/GraphTokenService.cs`**

```csharp
using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Options;

namespace NewMCP.Auth;

public interface IGraphTokenService
{
    Task<UserSession> ExchangeForGraphTokenAsync(
        string userId, string role, string mcpBearerToken,
        CancellationToken cancellationToken = default);
}

public sealed class GraphTokenService(IOptions<EntraOptions> options) : IGraphTokenService
{
    private static readonly Dictionary<string, string[]> RoleScopeMap = new()
    {
        ["SharePoint.Only"] = ["https://graph.microsoft.com/Sites.Read.All"],
        ["SharePoint.And.Outlook"] = [
            "https://graph.microsoft.com/Sites.Read.All",
            "https://graph.microsoft.com/Mail.Read",
            "https://graph.microsoft.com/Mail.Send"
        ]
    };

    public static string[] GetScopesForRole(string role) =>
        RoleScopeMap.TryGetValue(role, out var scopes)
            ? scopes
            : throw new ArgumentException($"Unknown role: {role}", nameof(role));

    public async Task<UserSession> ExchangeForGraphTokenAsync(
        string userId, string role, string mcpBearerToken,
        CancellationToken cancellationToken = default)
    {
        var opts = options.Value;
        var scopes = GetScopesForRole(role);

        var credential = new OnBehalfOfCredential(
            opts.TenantId,
            opts.ClientId,
            opts.ClientSecret,
            mcpBearerToken);

        var tokenRequestContext = new TokenRequestContext(scopes);
        var token = await credential.GetTokenAsync(tokenRequestContext, cancellationToken);

        return new UserSession(
            UserId: userId,
            Role: role,
            GraphAccessToken: token.Token,
            GraphRefreshToken: null,
            ExpiresAt: token.ExpiresOn);
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~GraphTokenServiceTests" -v minimal`
Expected: PASS — 3 tests

- [ ] **Step 5: Commit**

```bash
git add NewMCP/Auth/GraphTokenService.cs NewMCP.Tests/Auth/GraphTokenServiceTests.cs
git commit -m "feat: add GraphTokenService with OBO flow and role-to-scope mapping"
```

---

### Task 5: Create `McpAuthMiddleware`

**Files:**
- Create: `NewMCP/Middleware/McpAuthMiddleware.cs`
- Test: `NewMCP.Tests/Middleware/McpAuthMiddlewareTests.cs`

- [ ] **Step 1: Write the failing test**

Create `NewMCP.Tests/Middleware/McpAuthMiddlewareTests.cs`:

```csharp
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Moq;
using NewMCP.Auth;
using NewMCP.Middleware;
using System.Security.Claims;

namespace NewMCP.Tests.Middleware;

public sealed class McpAuthMiddlewareTests
{
    private readonly Mock<IGraphTokenService> _tokenSvcMock = new();
    private readonly Mock<TokenCacheService> _cacheMock;
    private readonly McpAuthMiddleware _sut;

    public McpAuthMiddlewareTests()
    {
        var distributedCacheMock = new Mock<Microsoft.Extensions.Caching.Distributed.IDistributedCache>();
        _cacheMock = new Mock<TokenCacheService>(distributedCacheMock.Object);

        _sut = new McpAuthMiddleware(
            next: _ => Task.CompletedTask,
            tokenService: _tokenSvcMock.Object,
            cache: _cacheMock.Object);
    }

    [Fact]
    public async Task InvokeAsync_Returns401_WhenUserNotAuthenticated()
    {
        var context = new DefaultHttpContext();
        context.Request.Path = "/mcp";
        context.User = new ClaimsPrincipal();

        await _sut.InvokeAsync(context);

        context.Response.StatusCode.Should().Be(401);
    }

    [Fact]
    public async Task InvokeAsync_UsesCache_WhenSessionExists()
    {
        var cachedSession = new UserSession("user-1", "SharePoint.Only", "cached-tok", null, DateTimeOffset.UtcNow.AddHours(1));
        _cacheMock.Setup(c => c.GetAsync("user-1", default)).ReturnsAsync(cachedSession);

        var context = new DefaultHttpContext();
        context.Request.Path = "/mcp";
        context.User = MakeAuthenticatedUser("user-1", "SharePoint.Only");
        context.Request.Headers.Authorization = "Bearer entra-token";

        bool nextCalled = false;
        var middleware = new McpAuthMiddleware(
            next: ctx => { nextCalled = true; return Task.CompletedTask; },
            tokenService: _tokenSvcMock.Object,
            cache: _cacheMock.Object);

        await middleware.InvokeAsync(context);

        nextCalled.Should().BeTrue();
        _tokenSvcMock.Verify(s => s.ExchangeForGraphTokenAsync(
            It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), default), Times.Never);
    }

    private static ClaimsPrincipal MakeAuthenticatedUser(string sub, string role)
    {
        var identity = new ClaimsIdentity(
        [
            new Claim("sub", sub),
            new Claim("roles", role)
        ], authenticationType: "Bearer");
        return new ClaimsPrincipal(identity);
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~McpAuthMiddlewareTests" -v minimal`
Expected: FAIL — `McpAuthMiddleware` does not exist

- [ ] **Step 3: Create `NewMCP/Middleware/McpAuthMiddleware.cs`**

```csharp
using Microsoft.AspNetCore.Http;
using NewMCP.Auth;
using System.Security.Claims;

namespace NewMCP.Middleware;

public sealed class McpAuthMiddleware(
    RequestDelegate next,
    IGraphTokenService tokenService,
    TokenCacheService cache)
{
    private static readonly TimeSpan RefreshWindow = TimeSpan.FromMinutes(5);

    public async Task InvokeAsync(HttpContext context)
    {
        if (!context.Request.Path.StartsWithSegments("/mcp"))
        {
            await next(context);
            return;
        }

        var userId = context.User.FindFirstValue("sub");
        var role = context.User.FindFirstValue("roles");

        if (userId is null || role is null)
        {
            context.Response.StatusCode = 401;
            return;
        }

        var session = await cache.GetAsync(userId, context.RequestAborted);

        if (session is not null && !session.IsExpiringSoon(RefreshWindow))
        {
            context.Items["UserSession"] = session;
            await next(context);
            return;
        }

        var authHeader = context.Request.Headers.Authorization.ToString();
        if (!authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            context.Response.StatusCode = 401;
            return;
        }

        var mcpToken = authHeader["Bearer ".Length..].Trim();

        try
        {
            session = await tokenService.ExchangeForGraphTokenAsync(
                userId, role, mcpToken, context.RequestAborted);
            await cache.SetAsync(session, context.RequestAborted);
            context.Items["UserSession"] = session;
            await next(context);
        }
        catch
        {
            await cache.RemoveAsync(userId, context.RequestAborted);
            context.Response.StatusCode = 401;
        }
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~McpAuthMiddlewareTests" -v minimal`
Expected: PASS — 3 tests

- [ ] **Step 5: Commit**

```bash
git add NewMCP/Middleware/McpAuthMiddleware.cs NewMCP.Tests/Middleware/McpAuthMiddlewareTests.cs
git commit -m "feat: add McpAuthMiddleware for Bearer token OBO exchange and session caching"
```

---

### Task 6: Create `GraphClientFactory`

**Files:**
- Create: `NewMCP/Graph/GraphClientFactory.cs`
- Test: `NewMCP.Tests/Graph/GraphClientFactoryTests.cs`

- [ ] **Step 1: Write the failing test**

Create `NewMCP.Tests/Graph/GraphClientFactoryTests.cs`:

```csharp
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Moq;
using NewMCP.Auth;
using NewMCP.Graph;

namespace NewMCP.Tests.Graph;

public sealed class GraphClientFactoryTests
{
    [Fact]
    public void GetClientForCurrentUser_ThrowsInvalidOperationException_WhenNoSession()
    {
        var accessorMock = new Mock<IHttpContextAccessor>();
        var cacheMock = new Mock<TokenCacheService>(
            new Mock<Microsoft.Extensions.Caching.Distributed.IDistributedCache>().Object);

        var context = new DefaultHttpContext();
        accessorMock.Setup(a => a.HttpContext).Returns(context);

        var factory = new GraphClientFactory(accessorMock.Object, cacheMock.Object);

        var act = () => factory.GetClientForCurrentUser();

        act.Should().Throw<InvalidOperationException>()
           .WithMessage("*No authenticated user session*");
    }

    [Fact]
    public void GetClientForCurrentUser_ReturnsClient_WhenSessionPresent()
    {
        var session = new UserSession("user-1", "SharePoint.Only", "graph-token", null, DateTimeOffset.UtcNow.AddHours(1));

        var accessorMock = new Mock<IHttpContextAccessor>();
        var context = new DefaultHttpContext();
        context.Items["UserSession"] = session;
        accessorMock.Setup(a => a.HttpContext).Returns(context);

        var cacheMock = new Mock<TokenCacheService>(
            new Mock<Microsoft.Extensions.Caching.Distributed.IDistributedCache>().Object);

        var factory = new GraphClientFactory(accessorMock.Object, cacheMock.Object);

        var client = factory.GetClientForCurrentUser();

        client.Should().NotBeNull();
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~GraphClientFactoryTests" -v minimal`
Expected: FAIL — `GraphClientFactory` does not exist

- [ ] **Step 3: Create `NewMCP/Graph/GraphClientFactory.cs`**

```csharp
using Microsoft.AspNetCore.Http;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using NewMCP.Auth;

namespace NewMCP.Graph;

public sealed class GraphClientFactory(IHttpContextAccessor httpContextAccessor, TokenCacheService cache)
{
    public GraphServiceClient GetClientForCurrentUser()
    {
        var context = httpContextAccessor.HttpContext
            ?? throw new InvalidOperationException("No HTTP context available.");

        var session = context.Items["UserSession"] as UserSession
            ?? throw new InvalidOperationException("No authenticated user session found in HTTP context.");

        var authProvider = new BaseBearerTokenAuthenticationProvider(
            new StaticTokenProvider(session.GraphAccessToken));

        return new GraphServiceClient(authProvider);
    }
}

internal sealed class StaticTokenProvider(string token) : IAccessTokenProvider
{
    public Task<string> GetAuthorizationTokenAsync(
        Uri uri,
        Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = default) => Task.FromResult(token);

    public AllowedHostsValidator AllowedHostsValidator { get; } =
        new(["graph.microsoft.com"]);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~GraphClientFactoryTests" -v minimal`
Expected: PASS — 2 tests

- [ ] **Step 5: Commit**

```bash
git add NewMCP/Graph/GraphClientFactory.cs NewMCP.Tests/Graph/GraphClientFactoryTests.cs
git commit -m "feat: add GraphClientFactory resolving GraphServiceClient from cached user session"
```

---

### Task 7: Create `OAuthDiscoveryController`

**Files:**
- Create: `NewMCP/Controllers/OAuthDiscoveryController.cs`
- Test: `NewMCP.Tests/Controllers/OAuthDiscoveryControllerTests.cs`

- [ ] **Step 1: Create `NewMCP/Controllers/OAuthDiscoveryController.cs`**

```csharp
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using NewMCP.Auth;

namespace NewMCP.Controllers;

[ApiController]
[AllowAnonymous]
public sealed class OAuthDiscoveryController(IOptions<EntraOptions> options) : ControllerBase
{
    [HttpGet("/.well-known/oauth-authorization-server")]
    public IActionResult GetDiscovery()
    {
        var opts = options.Value;
        var baseUrl = $"{opts.Instance}{opts.TenantId}";
        return Ok(new
        {
            issuer = baseUrl,
            authorization_endpoint = $"{baseUrl}/oauth2/v2.0/authorize",
            token_endpoint = $"{baseUrl}/oauth2/v2.0/token",
            response_types_supported = new[] { "code" },
            grant_types_supported = new[] { "authorization_code" },
            code_challenge_methods_supported = new[] { "S256" }
        });
    }
}
```

- [ ] **Step 2: Write the integration test (runs after Task 14)**

Create `NewMCP.Tests/Controllers/OAuthDiscoveryControllerTests.cs`:

```csharp
using FluentAssertions;
using Microsoft.AspNetCore.Mvc.Testing;
using System.Net;

namespace NewMCP.Tests.Controllers;

public sealed class OAuthDiscoveryControllerTests(WebApplicationFactory<Program> factory)
    : IClassFixture<WebApplicationFactory<Program>>
{
    [Fact]
    public async Task GetDiscovery_Returns200_WithExpectedFields()
    {
        var client = factory.CreateClient();

        var response = await client.GetAsync("/.well-known/oauth-authorization-server");

        response.StatusCode.Should().Be(HttpStatusCode.OK);
        var body = await response.Content.ReadAsStringAsync();
        body.Should().Contain("authorization_endpoint");
        body.Should().Contain("token_endpoint");
        body.Should().Contain("issuer");
    }
}
```

Note: This test references `Program` which requires `public partial class Program {}` — added in Task 14. Run after Task 14 completes.

- [ ] **Step 3: Commit**

```bash
git add NewMCP/Controllers/OAuthDiscoveryController.cs NewMCP.Tests/Controllers/OAuthDiscoveryControllerTests.cs
git commit -m "feat: add OAuthDiscoveryController serving MCP auth discovery metadata"
```

---

### Task 8: Refactor `RoleBasedAccessControl`

**Files:**
- Modify: `NewMCP/Tools/RoleBasedAccessControl.cs`
- Modify or create: `NewMCP.Tests/Tools/RoleBasedAccessControlTests.cs`

- [ ] **Step 1: Read the current file**

Read `NewMCP/Tools/RoleBasedAccessControl.cs` to understand what exists.

- [ ] **Step 2: Write the failing test**

Create/replace `NewMCP.Tests/Tools/RoleBasedAccessControlTests.cs`:

```csharp
using FluentAssertions;
using NewMCP.Tools;

namespace NewMCP.Tests.Tools;

public sealed class RoleBasedAccessControlTests
{
    private readonly RoleBasedAccessControl _sut = new();

    [Theory]
    [InlineData("SharePoint.Only", "SharePoint", true)]
    [InlineData("SharePoint.Only", "Email", false)]
    [InlineData("SharePoint.And.Outlook", "SharePoint", true)]
    [InlineData("SharePoint.And.Outlook", "Email", true)]
    public void HasAccess_ReturnsExpectedResult(string role, string category, bool expected)
    {
        _sut.HasAccess(role, category).Should().Be(expected);
    }

    [Fact]
    public void GetAllowedCategories_SharePointOnly_ReturnsOnlySharePoint()
    {
        _sut.GetAllowedCategories("SharePoint.Only")
            .Should().BeEquivalentTo(["SharePoint"]);
    }

    [Fact]
    public void GetAllowedCategories_SharePointAndOutlook_ReturnsBothCategories()
    {
        _sut.GetAllowedCategories("SharePoint.And.Outlook")
            .Should().BeEquivalentTo(["SharePoint", "Email"]);
    }

    [Theory]
    [InlineData("SharePoint.Only", "Email")]
    [InlineData("Unknown.Role", "SharePoint")]
    public void ValidateAccess_Throws_WhenAccessDenied(string role, string category)
    {
        var act = () => _sut.ValidateAccess(role, category);
        act.Should().Throw<UnauthorizedAccessException>();
    }

    [Fact]
    public void ValidateAccess_DoesNotThrow_WhenAccessGranted()
    {
        var act = () => _sut.ValidateAccess("SharePoint.Only", "SharePoint");
        act.Should().NotThrow();
    }
}
```

- [ ] **Step 3: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~RoleBasedAccessControlTests" -v minimal`
Expected: FAIL — current `RoleBasedAccessControl` has different API

- [ ] **Step 4: Rewrite `NewMCP/Tools/RoleBasedAccessControl.cs`**

```csharp
namespace NewMCP.Tools;

public sealed class RoleBasedAccessControl
{
    private static readonly Dictionary<string, HashSet<string>> RolePermissions = new()
    {
        ["SharePoint.Only"] = ["SharePoint"],
        ["SharePoint.And.Outlook"] = ["SharePoint", "Email"]
    };

    public bool HasAccess(string role, string category) =>
        RolePermissions.TryGetValue(role, out var categories) && categories.Contains(category);

    public IReadOnlySet<string> GetAllowedCategories(string role) =>
        RolePermissions.TryGetValue(role, out var categories)
            ? categories
            : new HashSet<string>();

    public void ValidateAccess(string role, string category)
    {
        if (!HasAccess(role, category))
            throw new UnauthorizedAccessException(
                $"Role '{role}' does not have access to category '{category}'.");
    }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~RoleBasedAccessControlTests" -v minimal`
Expected: PASS — all tests

- [ ] **Step 6: Commit**

```bash
git add NewMCP/Tools/RoleBasedAccessControl.cs NewMCP.Tests/Tools/RoleBasedAccessControlTests.cs
git commit -m "refactor: RoleBasedAccessControl — role-based API, remove MCP tool attrs and env var"
```

---

### Task 9: Create `GraphRetryHandler`

**Files:**
- Create: `NewMCP/Graph/GraphRetryHandler.cs`
- Test: `NewMCP.Tests/Graph/GraphRetryHandlerTests.cs`

- [ ] **Step 1: Write the failing test**

Create `NewMCP.Tests/Graph/GraphRetryHandlerTests.cs`:

```csharp
using FluentAssertions;
using NewMCP.Graph;
using System.Net;

namespace NewMCP.Tests.Graph;

public sealed class GraphRetryHandlerTests
{
    [Fact]
    public async Task SendAsync_RetriesOn429_AndSucceedsOnSecondAttempt()
    {
        int callCount = 0;
        var handler = new GraphRetryHandler
        {
            InnerHandler = new DelegatingHandlerStub(request =>
            {
                callCount++;
                if (callCount == 1)
                {
                    var throttled = new HttpResponseMessage(HttpStatusCode.TooManyRequests);
                    throttled.Headers.Add("Retry-After", "0");
                    return Task.FromResult(throttled);
                }
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK));
            })
        };

        var invoker = new HttpMessageInvoker(handler);
        var response = await invoker.SendAsync(
            new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me"), default);

        response.StatusCode.Should().Be(HttpStatusCode.OK);
        callCount.Should().Be(2);
    }

    [Fact]
    public async Task SendAsync_Returns429_AfterMaxRetries()
    {
        var handler = new GraphRetryHandler
        {
            InnerHandler = new DelegatingHandlerStub(_ =>
            {
                var throttled = new HttpResponseMessage(HttpStatusCode.TooManyRequests);
                throttled.Headers.Add("Retry-After", "0");
                return Task.FromResult(throttled);
            })
        };

        var invoker = new HttpMessageInvoker(handler);
        var response = await invoker.SendAsync(
            new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me"), default);

        response.StatusCode.Should().Be(HttpStatusCode.TooManyRequests);
    }
}

internal sealed class DelegatingHandlerStub(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler)
    : HttpMessageHandler
{
    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken) => handler(request);
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~GraphRetryHandlerTests" -v minimal`
Expected: FAIL — `GraphRetryHandler` does not exist

- [ ] **Step 3: Create `NewMCP/Graph/GraphRetryHandler.cs`**

```csharp
using System.Net;

namespace NewMCP.Graph;

public sealed class GraphRetryHandler : DelegatingHandler
{
    private const int MaxRetries = 3;

    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken)
    {
        HttpResponseMessage response = null!;

        for (int attempt = 0; attempt <= MaxRetries; attempt++)
        {
            response = await base.SendAsync(request, cancellationToken);

            if (response.StatusCode != HttpStatusCode.TooManyRequests || attempt == MaxRetries)
                return response;

            var retryAfter = response.Headers.RetryAfter?.Delta ?? TimeSpan.FromSeconds(1);
            if (retryAfter > TimeSpan.Zero)
                await Task.Delay(retryAfter, cancellationToken);
        }

        return response;
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~GraphRetryHandlerTests" -v minimal`
Expected: PASS — 2 tests

- [ ] **Step 5: Commit**

```bash
git add NewMCP/Graph/GraphRetryHandler.cs NewMCP.Tests/Graph/GraphRetryHandlerTests.cs
git commit -m "feat: add GraphRetryHandler for HTTP 429 / Retry-After backoff"
```

---

### Task 10: Create tool category system and `McpToolFilterMiddleware`

**Files:**
- Create: `NewMCP/Tools/ToolCategoryAttribute.cs`
- Create: `NewMCP/Tools/ToolCategoryRegistry.cs`
- Create: `NewMCP/Middleware/McpToolFilterMiddleware.cs`
- Test: `NewMCP.Tests/Middleware/McpToolFilterMiddlewareTests.cs`

- [ ] **Step 1: Create `NewMCP/Tools/ToolCategoryAttribute.cs`**

```csharp
namespace NewMCP.Tools;

[AttributeUsage(AttributeTargets.Method)]
public sealed class ToolCategoryAttribute(string category) : Attribute
{
    public string Category { get; } = category;
}
```

- [ ] **Step 2: Create `NewMCP/Tools/ToolCategoryRegistry.cs`**

```csharp
using ModelContextProtocol.Server;
using System.Reflection;

namespace NewMCP.Tools;

public sealed class ToolCategoryRegistry
{
    private readonly Dictionary<string, string> _toolCategories;

    public ToolCategoryRegistry(Assembly assembly)
    {
        _toolCategories = assembly.GetTypes()
            .SelectMany(t => t.GetMethods())
            .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() is not null
                     && m.GetCustomAttribute<ToolCategoryAttribute>() is not null)
            .ToDictionary(
                m => m.GetCustomAttribute<McpServerToolAttribute>()!.Name ?? m.Name,
                m => m.GetCustomAttribute<ToolCategoryAttribute>()!.Category);
    }

    public string? GetCategory(string toolName) =>
        _toolCategories.TryGetValue(toolName, out var cat) ? cat : null;
}
```

- [ ] **Step 3: Write the failing test for `McpToolFilterMiddleware`**

Create `NewMCP.Tests/Middleware/McpToolFilterMiddlewareTests.cs`:

```csharp
using FluentAssertions;
using Microsoft.AspNetCore.Http;
using Moq;
using NewMCP.Auth;
using NewMCP.Middleware;
using NewMCP.Tools;
using System.Text;
using System.Text.Json;

namespace NewMCP.Tests.Middleware;

public sealed class McpToolFilterMiddlewareTests
{
    private static readonly string ToolsListResponse = JsonSerializer.Serialize(new
    {
        jsonrpc = "2.0",
        id = 1,
        result = new
        {
            tools = new[]
            {
                new { name = "GetSites", description = "Gets SharePoint sites" },
                new { name = "GetEmails", description = "Gets emails" }
            }
        }
    });

    [Fact]
    public async Task InvokeAsync_FiltersEmailTools_ForSharePointOnlyUser()
    {
        var registryMock = new Mock<ToolCategoryRegistry>(typeof(object).Assembly);
        registryMock.Setup(r => r.GetCategory("GetSites")).Returns("SharePoint");
        registryMock.Setup(r => r.GetCategory("GetEmails")).Returns("Email");

        var rbac = new RoleBasedAccessControl();
        var session = new UserSession("user-1", "SharePoint.Only", "tok", null, DateTimeOffset.UtcNow.AddHours(1));

        var middleware = new McpToolFilterMiddleware(
            next: ctx =>
            {
                var responseBytes = Encoding.UTF8.GetBytes(ToolsListResponse);
                return ctx.Response.Body.WriteAsync(responseBytes).AsTask();
            },
            registry: registryMock.Object,
            rbac: rbac);

        var context = new DefaultHttpContext();
        context.Request.Path = "/mcp";
        context.Request.Method = "POST";
        context.Items["UserSession"] = session;
        context.Response.Body = new MemoryStream();

        await middleware.InvokeAsync(context);

        context.Response.Body.Seek(0, SeekOrigin.Begin);
        var body = await new StreamReader(context.Response.Body).ReadToEndAsync();
        var doc = JsonDocument.Parse(body);
        var tools = doc.RootElement.GetProperty("result").GetProperty("tools");

        tools.GetArrayLength().Should().Be(1);
        tools[0].GetProperty("name").GetString().Should().Be("GetSites");
    }
}
```

- [ ] **Step 4: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~McpToolFilterMiddlewareTests" -v minimal`
Expected: FAIL — `McpToolFilterMiddleware` does not exist

- [ ] **Step 5: Create `NewMCP/Middleware/McpToolFilterMiddleware.cs`**

```csharp
using Microsoft.AspNetCore.Http;
using NewMCP.Auth;
using NewMCP.Tools;
using System.Text;
using System.Text.Json.Nodes;

namespace NewMCP.Middleware;

public sealed class McpToolFilterMiddleware(
    RequestDelegate next,
    ToolCategoryRegistry registry,
    RoleBasedAccessControl rbac)
{
    public async Task InvokeAsync(HttpContext context)
    {
        if (!IsToolsListRequest(context))
        {
            await next(context);
            return;
        }

        var session = context.Items["UserSession"] as UserSession;
        if (session is null)
        {
            await next(context);
            return;
        }

        var originalBody = context.Response.Body;
        using var buffer = new MemoryStream();
        context.Response.Body = buffer;

        await next(context);

        buffer.Seek(0, SeekOrigin.Begin);
        var responseText = await new StreamReader(buffer).ReadToEndAsync();
        var filtered = FilterTools(responseText, session.Role);

        context.Response.Body = originalBody;
        var filteredBytes = Encoding.UTF8.GetBytes(filtered);
        context.Response.ContentLength = filteredBytes.Length;
        await context.Response.Body.WriteAsync(filteredBytes);
    }

    private static bool IsToolsListRequest(HttpContext context) =>
        context.Request.Path.StartsWithSegments("/mcp") &&
        context.Request.Method == HttpMethods.Post;

    private string FilterTools(string responseJson, string role)
    {
        try
        {
            var node = JsonNode.Parse(responseJson);
            var tools = node?["result"]?["tools"]?.AsArray();
            if (tools is null) return responseJson;

            var allowed = new JsonArray();
            foreach (var tool in tools)
            {
                var name = tool?["name"]?.GetValue<string>();
                if (name is null) continue;

                var category = registry.GetCategory(name);
                if (category is null || rbac.HasAccess(role, category))
                    allowed.Add(JsonNode.Parse(tool!.ToJsonString()));
            }

            node!["result"]!["tools"] = allowed;
            return node.ToJsonString();
        }
        catch
        {
            return responseJson;
        }
    }
}
```

- [ ] **Step 6: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~McpToolFilterMiddlewareTests" -v minimal`
Expected: PASS — 1 test

- [ ] **Step 7: Commit**

```bash
git add NewMCP/Tools/ToolCategoryAttribute.cs NewMCP/Tools/ToolCategoryRegistry.cs NewMCP/Middleware/McpToolFilterMiddleware.cs NewMCP.Tests/Middleware/McpToolFilterMiddlewareTests.cs
git commit -m "feat: add ToolCategoryRegistry and McpToolFilterMiddleware for role-based tool visibility"
```

---

### Task 11: Refactor `SharePointTools`

**Files:**
- Modify: `NewMCP/Tools/SharePointTools.cs`
- Create or modify: `NewMCP.Tests/Tools/SharePointToolsTests.cs`

- [ ] **Step 1: Read the current file**

Read `NewMCP/Tools/SharePointTools.cs` to confirm constructor signature and existing methods.

- [ ] **Step 2: Write the failing test**

Create `NewMCP.Tests/Tools/SharePointToolsTests.cs`:

```csharp
using FluentAssertions;
using Moq;
using NewMCP.Graph;
using NewMCP.Tools;

namespace NewMCP.Tests.Tools;

public sealed class SharePointToolsTests
{
    [Fact]
    public void Constructor_AcceptsGraphClientFactoryAndRbac()
    {
        var httpAccessorMock = new Mock<Microsoft.AspNetCore.Http.IHttpContextAccessor>();
        var cacheMock = new Mock<NewMCP.Auth.TokenCacheService>(
            new Mock<Microsoft.Extensions.Caching.Distributed.IDistributedCache>().Object);
        var factoryMock = new Mock<GraphClientFactory>(httpAccessorMock.Object, cacheMock.Object);
        var rbac = new RoleBasedAccessControl();

        var tools = new SharePointTools(factoryMock.Object, rbac);
        tools.Should().NotBeNull();
    }
}
```

- [ ] **Step 3: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~SharePointToolsTests" -v minimal`
Expected: FAIL — constructor signature mismatch

- [ ] **Step 4: Rewrite `NewMCP/Tools/SharePointTools.cs`**

Replace the entire file:

```csharp
using System.ComponentModel;
using System.Text.Json;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using ModelContextProtocol.Server;
using NewMCP.Graph;

namespace NewMCP.Tools;

internal sealed class SharePointTools(GraphClientFactory graphClientFactory, RoleBasedAccessControl accessControl)
{
    private GraphServiceClient GetClient() => graphClientFactory.GetClientForCurrentUser();

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Gets the current user's profile information from SharePoint/OneDrive.")]
    public async Task<string> GetMyProfile(CancellationToken cancellationToken = default)
    {
        var me = await GetClient().Me.GetAsync(cancellationToken: cancellationToken);
        return JsonSerializer.Serialize(new
        {
            displayName = me?.DisplayName,
            mail = me?.Mail,
            userPrincipalName = me?.UserPrincipalName,
            id = me?.Id
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Gets list of SharePoint sites the user has access to.")]
    public async Task<string> GetSites(
        [Description("Maximum number of sites to return (default 10, max 200)")] int top = 10,
        CancellationToken cancellationToken = default)
    {
        top = Math.Min(Math.Max(top, 1), 200);
        var client = GetClient();
        var siteList = new List<object>();

        var page = await client.Sites.GetAsync(cancellationToken: cancellationToken);
        await PageIterator<Site, SiteCollectionResponse>.CreatePageIterator(
            client, page!, site =>
            {
                if (siteList.Count >= top) return false;
                siteList.Add(new
                {
                    name = site.Name,
                    description = site.Description,
                    webUrl = site.WebUrl,
                    id = site.Id,
                    displayName = site.DisplayName
                });
                return true;
            }).IterateAsync(cancellationToken);

        return JsonSerializer.Serialize(siteList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Gets a specific SharePoint site by ID or URL.")]
    public async Task<string> GetSiteById(
        [Description("The site ID or URL")] string siteId,
        CancellationToken cancellationToken = default)
    {
        var site = await GetClient().Sites[siteId].GetAsync(cancellationToken: cancellationToken);
        if (site is null) return "Site not found.";

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
    [ToolCategory("SharePoint")]
    [Description("Gets the drives (document libraries) for a SharePoint site.")]
    public async Task<string> GetDrives(
        [Description("The site ID or leave empty for root site")] string? siteId = null,
        CancellationToken cancellationToken = default)
    {
        var client = GetClient();
        var driveList = new List<object>();

        var page = string.IsNullOrEmpty(siteId)
            ? await client.Drives.GetAsync(cancellationToken: cancellationToken)
            : await client.Sites[siteId].Drives.GetAsync(cancellationToken: cancellationToken);

        await PageIterator<Drive, DriveCollectionResponse>.CreatePageIterator(
            client, page!, drive =>
            {
                if (driveList.Count >= 200) return false;
                driveList.Add(new
                {
                    name = drive.Name,
                    description = drive.Description,
                    driveType = drive.DriveType,
                    webUrl = drive.WebUrl,
                    id = drive.Id
                });
                return true;
            }).IterateAsync(cancellationToken);

        return JsonSerializer.Serialize(driveList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Gets the items (files and folders) in a folder.")]
    public async Task<string> GetFolderItems(
        [Description("The drive ID")] string driveId,
        [Description("The folder ID (use 'root' for root folder)")] string folderId = "root",
        [Description("Maximum number of items to return (default 50, max 200)")] int top = 50,
        CancellationToken cancellationToken = default)
    {
        top = Math.Min(Math.Max(top, 1), 200);
        var client = GetClient();
        var itemList = new List<object>();

        var page = await client.Drives[driveId].Items[folderId].Children
            .GetAsync(cancellationToken: cancellationToken);

        await PageIterator<DriveItem, DriveItemCollectionResponse>.CreatePageIterator(
            client, page!, item =>
            {
                if (itemList.Count >= top) return false;
                itemList.Add(new
                {
                    name = item.Name,
                    webUrl = item.WebUrl,
                    id = item.Id,
                    createdDateTime = item.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
                    lastModifiedDateTime = item.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
                    size = item.Size,
                    type = item.Folder != null ? "folder" : "file"
                });
                return true;
            }).IterateAsync(cancellationToken);

        return JsonSerializer.Serialize(itemList, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Searches for files in SharePoint/OneDrive using the Graph Search API.")]
    public async Task<string> SearchFiles(
        [Description("Search query string")] string query,
        [Description("Maximum number of results (default 25, max 200)")] int top = 25,
        CancellationToken cancellationToken = default)
    {
        top = Math.Min(Math.Max(top, 1), 200);
        var client = GetClient();

        var searchBody = new Microsoft.Graph.Search.QueryPostRequestBody
        {
            Requests =
            [
                new Microsoft.Graph.Models.SearchRequest
                {
                    EntityTypes = [Microsoft.Graph.Models.EntityType.DriveItem],
                    Query = new Microsoft.Graph.Models.SearchQuery { QueryString = query },
                    Size = top
                }
            ]
        };

        var result = await client.Search.Query.PostAsQueryPostResponseAsync(
            searchBody, cancellationToken: cancellationToken);

        var hits = result?.Value?
            .SelectMany(r => r.HitsContainers ?? [])
            .SelectMany(c => c.Hits ?? [])
            .Take(top)
            .Select(h => new
            {
                id = h.HitId,
                name = h.Resource?.AdditionalData.TryGetValue("name", out var n) == true ? n?.ToString() : null,
                webUrl = h.Resource?.AdditionalData.TryGetValue("webUrl", out var u) == true ? u?.ToString() : null
            })
            .ToList();

        return JsonSerializer.Serialize(hits ?? [], new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Gets a specific file or folder by ID.")]
    public async Task<string> GetItemById(
        [Description("The drive ID")] string driveId,
        [Description("The item ID")] string itemId,
        CancellationToken cancellationToken = default)
    {
        var item = await GetClient().Drives[driveId].Items[itemId]
            .GetAsync(cancellationToken: cancellationToken);
        if (item is null) return "Item not found.";

        return JsonSerializer.Serialize(new
        {
            name = item.Name,
            webUrl = item.WebUrl,
            id = item.Id,
            createdDateTime = item.CreatedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            lastModifiedDateTime = item.LastModifiedDateTime?.ToString("yyyy-MM-dd HH:mm:ss"),
            size = item.Size,
            type = item.Folder != null ? "folder" : "file"
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    [McpServerTool]
    [ToolCategory("SharePoint")]
    [Description("Downloads file content as base64.")]
    public async Task<string> DownloadFile(
        [Description("The drive ID")] string driveId,
        [Description("The file item ID")] string itemId,
        CancellationToken cancellationToken = default)
    {
        try
        {
            using var stream = await GetClient().Drives[driveId].Items[itemId].Content
                .GetAsync(cancellationToken: cancellationToken);
            using var ms = new MemoryStream();
            await stream!.CopyToAsync(ms, cancellationToken);
            var bytes = ms.ToArray();
            return JsonSerializer.Serialize(new
            {
                success = true,
                content = Convert.ToBase64String(bytes),
                size = bytes.Length
            }, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (ODataError ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Error?.Message ?? ex.Message });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Message });
        }
    }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~SharePointToolsTests" -v minimal`
Expected: PASS

- [ ] **Step 6: Build to check for errors**

Run: `dotnet build NewMCP/NewMCP.csproj`
Expected: No errors

- [ ] **Step 7: Commit**

```bash
git add NewMCP/Tools/SharePointTools.cs NewMCP.Tests/Tools/SharePointToolsTests.cs
git commit -m "refactor: SharePointTools — inject GraphClientFactory, fix SearchFiles, add CancellationToken and PageIterator"
```

---

### Task 12: Refactor `EmailTools`

**Files:**
- Modify: `NewMCP/Tools/EmailTools.cs`
- Create: `NewMCP.Tests/Tools/EmailToolsTests.cs`

- [ ] **Step 1: Read the current `EmailTools.cs`**

Read `NewMCP/Tools/EmailTools.cs` to understand current constructor and methods.

- [ ] **Step 2: Write the failing test**

Create `NewMCP.Tests/Tools/EmailToolsTests.cs`:

```csharp
using FluentAssertions;
using Moq;
using NewMCP.Graph;
using NewMCP.Tools;

namespace NewMCP.Tests.Tools;

public sealed class EmailToolsTests
{
    [Fact]
    public void Constructor_AcceptsGraphClientFactoryAndRbac()
    {
        var httpAccessorMock = new Mock<Microsoft.AspNetCore.Http.IHttpContextAccessor>();
        var cacheMock = new Mock<NewMCP.Auth.TokenCacheService>(
            new Mock<Microsoft.Extensions.Caching.Distributed.IDistributedCache>().Object);
        var factoryMock = new Mock<GraphClientFactory>(httpAccessorMock.Object, cacheMock.Object);
        var rbac = new RoleBasedAccessControl();

        var tools = new EmailTools(factoryMock.Object, rbac);
        tools.Should().NotBeNull();
    }
}
```

- [ ] **Step 3: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~EmailToolsTests" -v minimal`
Expected: FAIL — constructor signature mismatch

- [ ] **Step 4: Refactor `NewMCP/Tools/EmailTools.cs`**

Read the current file, then apply these changes:

1. Replace the constructor with: `internal sealed class EmailTools(GraphClientFactory graphClientFactory, RoleBasedAccessControl accessControl)`
2. Add private helper: `private GraphServiceClient GetClient() => graphClientFactory.GetClientForCurrentUser();`
3. Remove `_oauthService`, `_graphClient`, `_tenantId`, `_clientId`, `_clientSecret`, `_currentUserId` fields and all env var reads
4. Replace every `GetGraphClient()` call with `GetClient()`
5. Add `[ToolCategory("Email")]` to every `[McpServerTool]` method
6. Add `CancellationToken cancellationToken = default` as the final parameter on every public async method
7. Pass `cancellationToken: cancellationToken` to every Graph API async call

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~EmailToolsTests" -v minimal`
Expected: PASS

- [ ] **Step 6: Build to check for errors**

Run: `dotnet build NewMCP/NewMCP.csproj`
Expected: No errors

- [ ] **Step 7: Commit**

```bash
git add NewMCP/Tools/EmailTools.cs NewMCP.Tests/Tools/EmailToolsTests.cs
git commit -m "refactor: EmailTools — inject GraphClientFactory, add CancellationToken and ToolCategory"
```

---

### Task 13: Remove dev scaffold files

**Files:**
- Delete: `NewMCP/Tools/RandomNumberTools.cs`
- Delete: `NewMCP/Services/OAuthAuthenticationService.cs`
- Delete: `NewMCP.Tests/Services/OAuthAuthenticationServiceTests.cs`

- [ ] **Step 1: Verify files exist before deleting**

Run:
```bash
ls NewMCP/Tools/RandomNumberTools.cs NewMCP/Services/OAuthAuthenticationService.cs NewMCP.Tests/Services/OAuthAuthenticationServiceTests.cs
```

- [ ] **Step 2: Delete files**

```bash
rm NewMCP/Tools/RandomNumberTools.cs
rm NewMCP/Services/OAuthAuthenticationService.cs
rm NewMCP.Tests/Services/OAuthAuthenticationServiceTests.cs
```

- [ ] **Step 3: Build to confirm no remaining references**

Run: `dotnet build NewMCP/NewMCP.csproj`
Expected: No errors (fix any compile errors from removed types before proceeding)

- [ ] **Step 4: Run all tests**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj -v minimal`
Expected: All remaining tests pass

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m "chore: remove RandomNumberTools, OAuthAuthenticationService dev scaffolding"
```

---

### Task 14: Rewrite `Program.cs` and integration tests

**Files:**
- Rewrite: `NewMCP/Program.cs`
- Create: `NewMCP/appsettings.json`
- Create: `NewMCP.Tests/Integration/McpEndpointTests.cs`

- [ ] **Step 1: Read current `Program.cs`**

Read `NewMCP/Program.cs` to confirm what exists before rewriting.

- [ ] **Step 2: Write the failing integration test**

Create `NewMCP.Tests/Integration/McpEndpointTests.cs`:

```csharp
using FluentAssertions;
using Microsoft.AspNetCore.Mvc.Testing;
using System.Net;

namespace NewMCP.Tests.Integration;

public sealed class McpEndpointTests(WebApplicationFactory<Program> factory)
    : IClassFixture<WebApplicationFactory<Program>>
{
    [Fact]
    public async Task DiscoveryEndpoint_Returns200()
    {
        var client = factory.CreateClient();
        var response = await client.GetAsync("/.well-known/oauth-authorization-server");
        response.StatusCode.Should().Be(HttpStatusCode.OK);
    }

    [Fact]
    public async Task McpEndpoint_Returns401_WithNoBearer()
    {
        var client = factory.CreateClient();
        var response = await client.PostAsync("/mcp",
            new StringContent("{}", System.Text.Encoding.UTF8, "application/json"));
        response.StatusCode.Should().Be(HttpStatusCode.Unauthorized);
    }

    [Fact]
    public async Task HealthEndpoint_Returns200()
    {
        var client = factory.CreateClient();
        var response = await client.GetAsync("/health");
        response.StatusCode.Should().Be(HttpStatusCode.OK);
    }
}
```

- [ ] **Step 3: Run test to verify it fails**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~McpEndpointTests" -v minimal`
Expected: FAIL — `Program` not accessible as `public partial`

- [ ] **Step 4: Create `NewMCP/appsettings.json`**

```json
{
  "Entra": {
    "TenantId": "YOUR_TENANT_ID",
    "ClientId": "YOUR_CLIENT_ID",
    "ClientSecret": "YOUR_CLIENT_SECRET",
    "Instance": "https://login.microsoftonline.com/"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  }
}
```

- [ ] **Step 5: Rewrite `NewMCP/Program.cs`**

```csharp
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using NewMCP.Auth;
using NewMCP.Graph;
using NewMCP.Middleware;
using NewMCP.Tools;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<EntraOptions>(
    builder.Configuration.GetSection(EntraOptions.SectionName));

var entraSection = builder.Configuration.GetSection(EntraOptions.SectionName);
var tenantId = entraSection["TenantId"] ?? "placeholder";
var clientId = entraSection["ClientId"] ?? "placeholder";
var instance = entraSection["Instance"] ?? "https://login.microsoftonline.com/";

builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.Authority = $"{instance}{tenantId}/v2.0";
        options.Audience = clientId;
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidateLifetime = true,
            ValidateIssuerSigningKey = true
        };
    });

builder.Services.AddAuthorization();

var redisConnection = builder.Configuration.GetConnectionString("Redis");
if (!string.IsNullOrEmpty(redisConnection))
    builder.Services.AddStackExchangeRedisCache(o => o.Configuration = redisConnection);
else
    builder.Services.AddDistributedMemoryCache();

builder.Services.AddSingleton<TokenCacheService>();
builder.Services.AddSingleton<IGraphTokenService, GraphTokenService>();
builder.Services.AddHttpContextAccessor();
builder.Services.AddScoped<GraphClientFactory>();
builder.Services.AddSingleton<RoleBasedAccessControl>();
builder.Services.AddSingleton(new ToolCategoryRegistry(Assembly.GetExecutingAssembly()));
builder.Services.AddTransient<GraphRetryHandler>();

builder.Services.AddMcpServer()
    .WithTools<SharePointTools>()
    .WithTools<EmailTools>()
    .WithStreamableHttpServerTransport();

builder.Services.AddControllers();
builder.Services.AddHealthChecks();

var app = builder.Build();

app.UseAuthentication();
app.UseAuthorization();

app.UseMiddleware<McpAuthMiddleware>();
app.UseMiddleware<McpToolFilterMiddleware>();

app.MapControllers();
app.MapHealthChecks("/health");
app.MapMcp("/mcp");

app.Run();

public partial class Program { }
```

- [ ] **Step 6: Run test to verify it passes**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj --filter "FullyQualifiedName~McpEndpointTests" -v minimal`
Expected: PASS — 3 integration tests pass

- [ ] **Step 7: Run all tests**

Run: `dotnet test NewMCP.Tests/NewMCP.Tests.csproj -v minimal`
Expected: All tests pass

- [ ] **Step 8: Commit**

```bash
git add NewMCP/Program.cs NewMCP/appsettings.json NewMCP.Tests/Integration/McpEndpointTests.cs
git commit -m "feat: rewrite Program.cs as ASP.NET Core WebApplication with Streamable HTTP MCP transport"
```

---

## Self-Review Against Spec

| Spec requirement | Covered by task |
|---|---|
| `WithStreamableHttpServerTransport()` + `app.MapMcp("/mcp")` | Task 14 |
| `/.well-known/oauth-authorization-server` discovery | Task 7 |
| Bearer JWT validation against Entra JWKS | Task 14 (AddJwtBearer) |
| OBO flow for delegated Graph token | Tasks 4, 5 |
| Cached Graph tokens keyed by userId | Tasks 3, 5 |
| 5-minute refresh window | Task 5 |
| Role → Graph scope mapping | Task 4 |
| `tools/list` filtering for SharePoint-only users | Task 10 |
| `tools/call` role enforcement (defence in depth) | Task 8 |
| `GraphClientFactory` with per-user token | Task 6 |
| `PageIterator<T>` on all list calls | Task 11 |
| `GraphRetryHandler` for HTTP 429 | Task 9 |
| Real `SearchFiles` via Graph Search API | Task 11 |
| `CancellationToken` on all tool methods | Tasks 11, 12 |
| `RandomNumberTools` removed | Task 13 |
| `public partial class Program {}` for `WebApplicationFactory` | Task 14 |
| `appsettings.json` with placeholder secrets | Task 14 |
