# Task Completion Checklist

When completing development tasks in this project, follow these steps:

## 1. Code Quality Checks
```bash
# Run Ruff linter to check code style
ruff check .

# Format code with Ruff
ruff format .

# Check for any Python syntax errors
python -m py_compile <file.py>
```

## 2. Dependency Management
```bash
# If new dependencies were added, sync them
uv sync

# Verify all dependencies are installed
uv pip list
```

## 3. Manual Testing
```bash
# Test the main entry point
uv run python main.py --help

# For API changes, test with the appropriate transport
uv run python main.py --transport streamable-http

# Check server health endpoint
curl http://localhost:8000/health
```

## 4. Documentation Updates
- Update README.md if functionality changed
- Update manifest.json if new tools or configuration added
- Update docstrings for any modified functions
- Add comments for complex logic

## 5. Git Workflow
```bash
# Check what changed
git status
git diff

# Stage changes
git add <files>

# Commit with descriptive message
git commit -m "type: description"
# Types: feat, fix, docs, style, refactor, test, chore

# Push to repository
git push origin <branch>
```

## 6. Environment Verification
- Ensure no sensitive data (API keys, secrets) in code
- Verify .env files are in .gitignore
- Check that OAuth credentials are properly configured
- Ensure OAUTHLIB_INSECURE_TRANSPORT is only set for development

## 7. Error Handling Verification
- Check that all API calls have proper error handling
- Verify authentication errors are caught and handled
- Ensure logging is appropriate (no sensitive data logged)

## Note on Testing
This project currently does not have automated tests (no pytest, unittest found). Testing is done manually through:
- Running the server with various configurations
- Testing individual tools through the MCP interface
- Checking OAuth flow manually
- Verifying API responses