# Google Workspace MCP Server Project Overview

## Project Purpose
This is a production-ready MCP (Model Context Protocol) server that integrates all major Google Workspace services with AI assistants. It enables full natural language control over Google Calendar, Drive, Gmail, Docs, Sheets, Slides, Forms, Tasks, Chat, and Custom Search through MCP clients and AI assistants.

## Tech Stack
- **Language**: Python 3.10+ (uses Python 3.11 specifically)
- **Package Manager**: UV (ultrafast Python package manager)
- **Web Framework**: FastAPI with FastMCP (v2.11.1)
- **Google Integration**: Google API Python Client (google-api-python-client)
- **Authentication**: OAuth 2.0 and OAuth 2.1 with google-auth-oauthlib
- **HTTP Client**: httpx and aiohttp
- **Security**: cryptography, pyjwt
- **Linting**: Ruff (>=0.12.4)

## Architecture
- MCP server implementation with stdio and streamable-http transports
- Service-oriented architecture with separate modules for each Google service
- Centralized authentication handling with OAuth 2.1 support
- Thread-safe sessions and service caching for performance
- Docker support for containerized deployment

## Key Features
- Multi-user authentication via OAuth 2.1
- Support for all major Google Workspace services
- Automatic token refresh and session management
- CORS proxy architecture for browser compatibility
- High performance with service caching and FastMCP integration