# CLAUDE.md — newMCP Project

## Project Overview

**Custom MCP Server: Claude ↔ Microsoft 365 Per-User Access Control**

This project builds a custom MCP (Model Context Protocol) server that enforces per-user access control between Claude and Microsoft 365. It replaces the built-in Claude M365 connector, which applies Microsoft Graph API permission scopes tenant-wide and cannot differentiate access by user or group.

### Problem Being Solved

- All users need SharePoint access through Claude.
- One designated user also needs Outlook access through Claude.
- The pre-built connector cannot enforce this without either granting Outlook access to everyone (unacceptable) or relying on users to self-manage settings (no central enforcement).

### Solution

A custom MCP server hosted on Azure (App Service) acts as middleware between Claude and Microsoft Graph API. Access control is enforced server-side using Microsoft Entra ID App Roles.

---

## Architecture

```
Claude (claude.ai) → Custom MCP Server (Azure App Service) → Microsoft Graph API → M365 Data
```

### Auth Flow (per request)

1. User authenticates via OAuth 2.0 authorization code flow through the MCP server.
2. Server reads the user's assigned Entra App Role at authentication time.
3. Based on the role, the server requests only the Graph API scopes that user is permitted.
4. **SharePoint-only users** → token scoped to `Sites.Read.All`
5. **Outlook-enabled user** → additional token scoped to `Mail.Read` + related mail permissions
6. No user can exceed their assigned permissions — enforcement is entirely server-side.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Language / Runtime | C# / .NET |
| Hosting | Azure App Service |
| Identity | Microsoft Entra ID (Azure AD) |
| Auth Protocol | OAuth 2.0 (authorization code flow) |
| API | Microsoft Graph API |
| MCP Protocol | MCP SDK (C# integration) |
| CI/CD (optional) | Azure DevOps or GitHub Actions |

---

## Entra ID Configuration

- Separate Entra App Registration from Anthropic's pre-built connector apps.
- Graph API scopes defined on the registration:
  - `Sites.Read.All` (SharePoint)
  - `Mail.Read` + related mail permissions (Outlook)
- **App Roles:**
  - `SharePoint.Only` — assigned to all standard users
  - `SharePoint.And.Outlook` — assigned only to the designated Outlook user
- Tenant-wide admin consent granted for required Graph API permissions.

---

## MCP Tools to Implement

| Tool | Scope | Users |
|---|---|---|
| SharePoint file listing | `Sites.Read.All` | All |
| SharePoint search | `Sites.Read.All` | All |
| SharePoint file read | `Sites.Read.All` | All |
| Outlook mail read | `Mail.Read` | Designated user only |
| Outlook mail search | `Mail.Read` | Designated user only |

---

## Project Phases

| Phase | Description | Est. Hours |
|---|---|---|
| 1 | Discovery & Architecture | 7 |
| 2 | Entra ID Configuration | 4.5 |
| 3 | MCP Server Development | 25 |
| 4 | Hosting & Deployment | 8 |
| 5 | Testing & Validation | 8 |
| 6 | Documentation & Handover | 7 |
| **Total** | | **59.5** |

### Phase 3 — MCP Server Development Breakdown

| Task | Est. Hours |
|---|---|
| Project scaffolding (C# .NET, MCP SDK, repo structure) | 3 |
| OAuth 2.0 implementation (auth code flow, token refresh, secure storage) | 6 |
| App Role → Graph API scope mapping logic | 4 |
| MCP tool definitions (SharePoint + Outlook endpoints) | 6 |
| Microsoft Graph API integration (pagination, errors, throttling) | 4 |
| Structured logging (auth events, tool invocations: user, scope, timestamp) | 2 |

---

## Key Development Guidelines

### Security
- All access control logic must live server-side — never rely on client-side enforcement.
- OAuth tokens must be stored securely; never log token values.
- Implement structured audit logging: every auth event and tool invocation must record user identity, scope requested, and timestamp.
- TLS required on all endpoints; configure IP allowlisting to Claude.ai egress ranges.
- No Graph API scope may be issued beyond the role-mapped set for that user.

### Graph API
- Handle pagination for all list-style responses.
- Implement retry/backoff for Graph API throttling (HTTP 429).
- Validate Graph API error responses and surface meaningful errors to MCP callers.

### MCP Protocol
- All tools must be MCP-protocol compliant.
- Tool definitions must clearly declare input schemas and return types.
- Outlook tools must be absent (not just blocked) for SharePoint-only users — do not expose tools the user cannot invoke.

### Code Style
- Follow standard C# / .NET conventions.
- Keep functions small and focused.
- Handle all errors explicitly — no silent failures.
- Use dependency injection throughout.

---

## Infrastructure & Deployment

- **Hosting:** Azure App Service (recommended)
- **Runtime:** C# .NET
- **Environment variables** used for all secrets and configuration (client ID, client secret, tenant ID, etc.) — never hardcode.
- Managed identity preferred where possible to reduce secret surface area.
- The pre-built Claude M365 connector must be decommissioned upon successful go-live.
- Register custom MCP server URL in Claude Team/Enterprise admin settings.

---

## Assumptions & Constraints

- Client has Microsoft 365 tenant with Entra ID and global admin access.
- Azure subscription available for hosting.
- Client has active Claude Team or Enterprise plan with admin access to connector settings.
- Any Graph API scopes beyond SharePoint and Outlook mail are **out of scope** — require a change request.
- CI/CD pipeline (Phase 4) is optional.

---

## Testing Requirements

| Test Scenario | Expected Result |
|---|---|
| SharePoint-only user — SharePoint tools | Access granted |
| SharePoint-only user — Outlook tools | Tools absent / blocked |
| Outlook-enabled user — SharePoint tools | Access granted |
| Outlook-enabled user — Outlook tools | Access granted |
| Scope escalation / role bypass attempt | Server-side enforcement holds |
| Audit log review | All auth events and tool calls logged correctly |