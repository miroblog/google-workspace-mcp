# Code Style and Conventions

## Python Code Style
- **Python Version**: 3.10+ (specifically 3.11)
- **Docstrings**: Module-level docstrings at the top of each file
- **Type Hints**: Used extensively for function parameters and return types
- **Imports**: Organized in groups (stdlib, third-party, local imports)
- **Logging**: Comprehensive logging using Python's logging module
- **Error Handling**: Dedicated error handling decorators and utility functions

## Naming Conventions
- **Modules**: Lowercase with underscores (e.g., `gmail_tools.py`, `calendar_tools.py`)
- **Classes**: PascalCase (e.g., `GoogleWorkspaceAuthProvider`)
- **Functions**: Lowercase with underscores (e.g., `get_authenticated_google_service`)
- **Private Functions**: Prefix with underscore (e.g., `_correct_time_format_for_api`)
- **Constants**: UPPERCASE with underscores (e.g., `GMAIL_BATCH_SIZE`, `SCOPES`)

## Project Structure
- **Service Modules**: Each Google service has its own directory (gmail/, gdrive/, gcalendar/, etc.)
- **Core Module**: Core functionality in core/ (server, config, utils, context)
- **Auth Module**: Authentication handling in auth/ (OAuth, session management, scopes)
- **Tool Files**: Each service has a `*_tools.py` file containing MCP tool implementations

## Best Practices
- Use decorators for service authentication (`@require_google_service`)
- Implement proper error handling with HttpError from googleapiclient
- Use async/await for asynchronous operations
- Maintain thread-safe operations for concurrent access
- Document all public functions with proper docstrings
- Use configuration constants from environment variables