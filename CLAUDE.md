# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

### Development Commands
- `uv run main.py` - Start server in stdio mode (default for MCP clients)
- `uv run main.py --transport streamable-http` - Start server in HTTP mode (for debugging/web interfaces)
- `uv run main.py --single-user` - Run in single-user mode (bypasses session mapping)
- `uv run main.py --tools gmail drive calendar` - Start with specific tools only

### Building DXT (Desktop Extension)
The DXT file is already built and available as `google_workspace_mcp.dxt`. The DXT contains:
- Python server code
- Dependencies via uv
- Manifest with configuration schema
- Ready for one-click Claude Desktop installation

To rebuild the DXT:
1. Ensure all dependencies are in `pyproject.toml`
2. Update version in both `pyproject.toml` and `manifest.json`
3. The DXT file combines the server code with the manifest for distribution

### Testing & Quality
- Development environment requires Python 3.10+
- Uses `ruff>=0.12.4` for code formatting and linting
- OAuth testing requires `OAUTHLIB_INSECURE_TRANSPORT=1` for local development

### Authentication Testing
```bash
python install_claude.py  # Guided Claude Desktop setup
```

## Architecture

### Core Components

**MCP Server Framework** (`core/server.py`)
- Built on FastMCP with custom CORSEnabledFastMCP class
- Supports both stdio and streamable-http transports
- Handles OAuth 2.0 and OAuth 2.1 authentication flows
- Transport-aware OAuth callback handling on port 8000

**Authentication System** (`auth/`)
- `service_decorator.py` - Main decorator for Google service authentication with 30-minute caching
- `google_auth.py` - Core OAuth 2.0 implementation
- `oauth21_integration.py` - OAuth 2.1 support for multi-user sessions
- `scopes.py` - Centralized scope management with predefined scope groups
- Service caching with TTL to reduce authentication overhead

**Service Modules** (g*/`)
- Each Google service has its own module (gmail/, gdrive/, gcalendar/, etc.)
- Tools use `@require_google_service(service_type, scopes)` decorator pattern
- Automatic service injection and authentication handling
- Support for both single and multiple service requirements

**Google Sheets CSV Download** (`gsheets/sheets_tools.py`)
- `download_sheet_as_csv` - Downloads spreadsheets as CSV files to temp directory
- Cross-platform support (macOS/Linux: `/tmp/`, Windows: `%TEMP%`)
- Enables data processing with pandas and other tools
- Uses Google Drive API's export functionality for optimal CSV conversion

### Key Patterns

**Service Decorator Pattern**
```python
@require_google_service("drive", "drive_read")
async def search_drive_files(service, user_google_email: str, query: str):
    # service parameter automatically injected and authenticated
    return service.files().list(q=query).execute()
```

**Scope Groups** - Centralized in `auth/scopes.py`:
- `gmail_read`, `gmail_send`, `drive_read`, `drive_file`
- `calendar_read`, `calendar_events`, `sheets_read`, `sheets_write`
- And more for each Google Workspace service

**Configuration Management** (`core/config.py`)
- Environment variable loading with `.env` file support
- Transport mode detection and OAuth URI generation
- Credential loading priority: env vars → .env file → client_secret.json

**Multi-Transport Support**
- Stdio mode: Starts minimal HTTP server on port 8000 for OAuth callbacks only
- HTTP mode: Uses FastAPI server for both MCP and OAuth handling
- Same OAuth flow works in both modes for consistency

### Service Architecture

**Modular Service Design**
- Each service module is dynamically imported based on `--tools` flag
- Service configurations in `SERVICE_CONFIGS` mapping
- Centralized error handling and token refresh logic
- Thread-safe session management with OAuth 2.1 support

**Authentication Flows**
- Legacy OAuth 2.0: File-based credential storage
- OAuth 2.1: Session-based with bearer tokens for multi-user support
- Transport-aware callbacks: HTTP server handles OAuth redirects in both modes
- Automatic token refresh with graceful error handling

**Caching Strategy**
- Service instances cached for 30 minutes per user/service/scope combination
- Cache keys include user email, service name, version, and sorted scopes
- Automatic cache invalidation on token refresh errors

### Error Handling

**Authentication Errors**
- `GoogleAuthenticationError` for auth-specific issues
- Automatic token refresh with user-friendly error messages
- Clear instructions for reauthentication when tokens expire
- Service cache clearing on token refresh failures

**Transport Isolation**
- OAuth callback server runs independently of main MCP transport
- Graceful degradation when OAuth callback server fails to start
- Proper cleanup of background servers on shutdown

This codebase implements a production-ready MCP server with comprehensive Google Workspace integration, featuring robust authentication, service caching, and multi-transport support for various deployment scenarios.