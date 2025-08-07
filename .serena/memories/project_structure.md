# Project Structure

## Root Directory Files
- `main.py` - Main entry point for the MCP server
- `pyproject.toml` - Python project configuration and dependencies
- `uv.lock` - Locked dependencies for UV package manager
- `manifest.json` - MCP server manifest with tool definitions
- `README.md` - Comprehensive project documentation
- `LICENSE` - MIT license
- `Dockerfile` - Container configuration for deployment
- `.env.oauth21` - OAuth 2.1 configuration template
- `install_claude.py` - Claude Desktop installation helper
- `smithery.yaml` - Smithery registry configuration
- `.python-version` - Python version specification (3.11)

## Core Module (`core/`)
- `server.py` - FastMCP server setup and configuration
- `config.py` - Configuration management and environment variables
- `context.py` - Context management for services
- `utils.py` - Utility functions and error handling
- `api_enablement.py` - Google API enablement utilities
- `comments.py` - Comment handling for Google services

## Authentication Module (`auth/`)
- `google_auth.py` - Google OAuth authentication
- `service_decorator.py` - Authentication decorators for tools
- `scopes.py` - Google API scope definitions
- `oauth21_session_store.py` - OAuth 2.1 session management
- `mcp_session_middleware.py` - Session middleware
- `auth_info_middleware.py` - Authentication info middleware
- `fastmcp_google_auth.py` - FastMCP Google auth provider
- `google_remote_auth_provider.py` - Remote auth provider (FastMCP 2.11.1+)

## Service Modules
Each Google service has its own directory with tool implementations:
- `gmail/` - Gmail tools and operations
- `gdrive/` - Google Drive file operations
- `gcalendar/` - Google Calendar management
- `gdocs/` - Google Docs operations
- `gsheets/` - Google Sheets operations
- `gslides/` - Google Slides operations
- `gforms/` - Google Forms management
- `gchat/` - Google Chat integration
- `gtasks/` - Google Tasks management
- `gsearch/` - Google Custom Search

## Hidden Directories
- `.serena/` - Serena project configuration
- `.claude/` - Claude-specific configuration
- `.venv/` - Python virtual environment (created by UV)

## Key Patterns
1. Each service module contains a `*_tools.py` file with MCP tool implementations
2. Tools use the `@require_google_service` decorator for authentication
3. All tools are registered with the FastMCP server
4. Service initialization is lazy-loaded for performance
5. Configuration is centralized in environment variables