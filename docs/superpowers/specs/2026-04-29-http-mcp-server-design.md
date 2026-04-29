# Design: HTTP Remote MCP Server Rearchitecture

**Date:** 2026-04-29
**Project:** newMCP — Custom MCP Server: Claude ↔ Microsoft 365 Per-User Access Control
**Status:** Approved

---

## Background

The existing codebase uses `WithStdioServerTransport()` — a transport model where Claude spawns a local process per session. The project spec requires a **remote HTTP server hosted on Azure App Service** (one URL serving all users). This design documents the rearchitecture required to align the implementation with the spec.

---

## Architecture

```
Claude.ai
  │  HTTPS POST /mcp  (Authorization: Bearer <entra_token>)
  ▼
Azure App Service  (:443)
  ├── /.well-known/oauth-authorization-server  → MCP auth discovery (points to Entra)
  ├── /mcp                                     → MCP SDK endpoint (tools)
  │     └── Middleware validates Bearer JWT, extracts userId + roles
  │     └── OBO flow: exchanges MCP token → Graph delegated token (cached)
  └── /auth/callback                           → Receives auth code (Entra redirect)
  │
  ▼
Microsoft Graph API  (per-user delegated token via OBO)
```

### Hosting

- **Runtime:** .NET 10, ASP.NET Core `WebApplication` (replaces generic `Host`)
- **Platform:** Azure App Service, Linux plan, HTTPS Only enforced
- **MCP transport:** Streamable HTTP via `app.MapMcp("/mcp")`
- **Token cache:** `IDistributedCache` — Azure Cache for Redis (production), in-memory (local dev). Stores Graph delegated tokens keyed by userId.

### Structural changes from current code

| Current | New |
|---|---|
| `Host.CreateDefaultBuilder` | `WebApplication.CreateBuilder` |
| `WithStdioServerTransport()` | `WithStreamableHttpServerTransport()` + `app.MapMcp("/mcp")` |
| `CURRENT_USER_ID` env var | User identity from validated Bearer JWT claims (per-request) |
| `USER_ROLE_MAPPING` env var | Roles read from `roles` claim in validated Entra JWT |
| OAuth exposed as MCP tools | MCP auth discovery + Bearer token validation middleware |
| `RandomNumberTools` | Removed (dev scaffold) |

---

## OAuth & Authorization Flow

The MCP HTTP protocol uses `Authorization: Bearer` headers — not session cookies. Claude's MCP client follows the MCP Authorization spec (2025-03-26): it discovers auth requirements, completes the OAuth flow in a browser popup, and sends the resulting token as a Bearer header on every MCP request.

### Flow

```
1. MCP auth discovery
   └── Claude fetches /.well-known/oauth-authorization-server
   └── Server returns Entra ID's OAuth metadata (authorization_endpoint, token_endpoint, etc.)
   └── Claude knows how to start the OAuth flow

2. User authorizes (browser popup, once per connector setup)
   └── Claude opens Entra ID login (PKCE, response_type=code)
   └── User logs in and consents
   └── Entra redirects → /auth/callback?code=...&state=...
   └── Server exchanges code for Entra access token + id token
   └── Server returns access token to Claude

3. Claude calls POST /mcp
   └── Authorization: Bearer <entra_access_token>
   └── Bearer token middleware:
       a. Validates JWT signature against Entra JWKS endpoint
       b. Extracts userId (sub claim) and roles (`roles` claim)
       c. Checks distributed cache for cached Graph delegated token
       d. If not cached: OBO flow — exchanges MCP Bearer token for Graph token
          (only permitted scopes for the user's role are requested)
       e. Caches { graphAccessToken, graphRefreshToken, role, expiry } keyed by userId
   └── MCP tools execute with per-user Graph client

4. Subsequent requests
   └── Same Bearer token sent by Claude
   └── Middleware validates + resolves from cache (fast path — no OBO round-trip)
   └── Graph token refreshed transparently if expiring within 5 minutes
```

### Role → Graph scope mapping

| Entra App Role | Graph API Scopes requested via OBO |
|---|---|
| `SharePoint.Only` | `Sites.Read.All` |
| `SharePoint.And.Outlook` | `Sites.Read.All`, `Mail.Read`, `Mail.Send` |

Roles are assigned in the Entra portal. The server reads them from the `roles` claim in the validated Entra JWT — `USER_ROLE_MAPPING` env var is removed.

### Token refresh

On each MCP request, middleware checks the cached Graph token expiry. If within 5 minutes, it refreshes using the stored refresh token before executing the tool. Refresh failures clear the cache entry and return an auth error, prompting Claude to re-authorize.

---

## Role-Based Tool Visibility

The spec requires Outlook tools to be **absent** (not just blocked) for SharePoint-only users.

### Mechanism

When Claude calls `tools/list`, a `McpToolFilter` middleware intercepts the response and removes tools the user's role does not permit:

- `EmailTools` methods tagged `[ToolCategory("Email")]`
- `SharePointTools` methods tagged `[ToolCategory("SharePoint")]`
- Filter reads `RoleBasedAccessControl.RoleToolPermissions` (existing mapping, unchanged)
- SharePoint-only users receive a `tools/list` response with no Email tools visible

`tools/call` still enforces role access as a second layer (defence in depth).

### Tool registration

All tools registered at startup via `WithTools<>()`. No per-session dynamic registration needed — filtering happens at the protocol response layer.

`RandomNumberTools` removed from registration entirely.

---

## Graph API Integration

### Per-user client resolution

`GraphClientFactory` (replaces ad-hoc `GetGraphClientForUser`) is injected into tools:

```
Tool method invoked
  └── GraphClientFactory.GetClientForCurrentUser()
      └── IHttpContextAccessor → HttpContext → Bearer token claims → userId
      └── Retrieves cached delegated Graph token for userId
      └── Returns GraphServiceClient scoped to that token
      └── Throws McpException("Unauthorized") if no valid token found
```

### Pagination

All list-style Graph calls (`GetSites`, `GetDrives`, `GetFolderItems`, `GetRecentFiles`, `GetSharedFiles`) iterate using `PageIterator<T>` until `NextLink` is null. Results capped at 200 items per call to avoid oversized responses.

### Throttling

A `GraphRetryHandler : DelegatingHandler` registered on the `GraphServiceClient` HTTP pipeline:
- Catches HTTP 429 responses
- Reads `Retry-After` header
- Waits and retries (max 3 attempts)
- Applies to all tools automatically — no per-tool retry logic

### `SearchFiles` fix

Current placeholder iterates drive root. Replace with `graphClient.Search.Query(new SearchRequestBody { Requests = [...] })` — the correct Graph Search API endpoint supporting full-text search across all accessible SharePoint content.

### `CancellationToken`

All public async tool methods gain `CancellationToken cancellationToken = default`. MCP SDK wires the connection cancellation token automatically.

---

## Logging & Security

### Structured audit log

Every event logged via `ILogger<T>` with structured fields (no token values or PII in logs):

| Event type | Fields |
|---|---|
| Auth success | `event`, `userId`, `role`, `scopesGranted`, `timestamp`, `ipAddress` |
| Auth failure | `event`, `reason`, `ipAddress`, `timestamp` |
| Tool invocation | `event`, `userId`, `tool`, `category`, `success`, `durationMs`, `timestamp` |
| Token refresh | `event`, `userId`, `expiresAt`, `timestamp` |

In Azure App Service, logs route to Application Insights or Log Analytics via built-in sink.

### Security controls

| Control | Implementation |
|---|---|
| TLS | Azure App Service `HTTPS Only = true` |
| IP allowlisting | App Service access restrictions: `/mcp` locked to Claude.ai egress IPs |
| Bearer token validation | JWT signature validated against Entra JWKS on every request; no token = 401 |
| Token cache encryption | ASP.NET Data Protection encrypts Graph tokens at rest in distributed cache |
| CSRF on OAuth | Cryptographically random `state` parameter validated at `/auth/callback` |
| Scope enforcement | OBO flow only requests role-mapped scopes; `GraphClientFactory` cannot escalate |
| Secret management | `ClientId`, `ClientSecret`, `TenantId`, Redis connection string from env vars / Key Vault references; never in `appsettings.json` |

---

## What Does Not Change

- `RoleBasedAccessControl` role-to-category mapping logic
- `SharePointTools` and `EmailTools` Graph API call implementations (except client injection + CancellationToken)
- xUnit + FluentAssertions test setup
- Entra App Role definitions (`SharePoint.Only`, `SharePoint.And.Outlook`)
- MCP tool names and input/output schemas (Claude connector config unchanged)

---

## Out of Scope

- Any Graph API scopes beyond SharePoint and Outlook mail (change request required)
- CI/CD pipeline (optional — Phase 4)
- Multi-tenant support
- Admin UI for role management (Entra portal is the admin surface)