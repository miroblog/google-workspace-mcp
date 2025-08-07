# Suggested Commands for Development

## Running the Project
```bash
# Install dependencies and run with UV (recommended)
uv run python main.py

# Run with different transports
uv run python main.py --transport stdio  # Default for Claude Desktop
uv run python main.py --transport streamable-http  # For HTTP/SSE support

# Run with specific tools only
uv run python main.py --tools gmail drive calendar

# Run in single-user mode
uv run python main.py --single-user
```

## Development Commands
```bash
# Install/sync dependencies
uv sync
uv sync --frozen --no-dev  # Production dependencies only

# Run linting
ruff check .
ruff format .

# Check Python version
python --version  # Should be 3.11.x

# List installed packages
uv pip list

# Run with environment variables
export GOOGLE_OAUTH_CLIENT_ID="your-client-id.apps.googleusercontent.com"
export GOOGLE_OAUTH_CLIENT_SECRET="your-secret"
export OAUTHLIB_INSECURE_TRANSPORT=1  # For development only
uv run python main.py
```

## Docker Commands
```bash
# Build Docker image
docker build -t workspace-mcp .

# Run Docker container
docker run -p 8000:8000 \
  -e GOOGLE_OAUTH_CLIENT_ID="your-id" \
  -e GOOGLE_OAUTH_CLIENT_SECRET="your-secret" \
  workspace-mcp
```

## System Commands (Darwin/macOS)
```bash
# Version control
git status
git diff
git add .
git commit -m "message"
git push

# File operations
ls -la
find . -name "*.py"
grep -r "pattern" .

# Process management
ps aux | grep python
kill -9 PID

# Environment
echo $PATH
which python
which uv
```

## Testing and Debugging
```bash
# Check server health
curl http://localhost:8000/health

# View logs
tail -f mcp_server_debug.log

# Test OAuth callback
curl http://localhost:8000/oauth2callback
```