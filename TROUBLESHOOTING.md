# Google Workspace MCP - Troubleshooting Guide

This guide addresses common issues and their solutions when using the Google Workspace MCP tools.

## Table of Contents
- [Formatting Issues](#formatting-issues)
- [Value Update Issues](#value-update-issues)
- [Range Specification Issues](#range-specification-issues)
- [Authentication Issues](#authentication-issues)
- [Table Styling Issues](#table-styling-issues)
- [General Best Practices](#general-best-practices)

## Formatting Issues

### Issue: `format_cells` font_size validation error
**Error**: `Input validation error: '12' is not valid under any of the given schemas`

**Solution**: 
- Font size must be an integer between 6 and 400
- Common valid values: 8, 10, 11, 12, 14, 16, 18, 24, 36
- Do NOT use strings or include "pt" suffix

**Correct Usage**:
```python
# ✅ CORRECT
format_cells(..., font_size=12)  # Integer value

# ❌ INCORRECT
format_cells(..., font_size="12")  # String
format_cells(..., font_size="12pt")  # With suffix
```

### Issue: Boolean parameters not working
**Error**: Validation errors when passing boolean values

**Solution**: 
The functions accept multiple formats for boolean parameters:
- Python booleans: `True`, `False`
- Integers: `1` (true), `0` (false)
- Strings: `"true"`, `"false"`, `"1"`, `"0"`, `"yes"`, `"no"`

**Examples**:
```python
# All of these work:
format_cells(..., bold=True)
format_cells(..., bold=1)
format_cells(..., bold="true")
format_cells(..., italic=0, underline="false")
```

## Value Update Issues

### Issue: `update_sheet_values` reports "1 cell updated" for large ranges
**Symptom**: Specifying `A1:N10` but getting "1 cell updated" message

**Explanation**: 
- This is Google Sheets API behavior when the range auto-adjusts
- The function may extend or shrink the range based on your data

**Solution**:
Use the right function for your needs:

| Use Case | Function | Behavior |
|----------|----------|----------|
| Single range, simple update | `update_sheet_values` | May auto-extend range |
| Precise range control | `batch_update_values` | Exact range control |
| Multiple ranges | `batch_update_values` | Better performance |
| Appending data | `append_sheet_values` | Finds end of data |

**Example**:
```python
# For precise control over a specific range:
batch_update_values(
    ...,
    updates=[
        {"range": "Sheet1!A1:C3", "values": [[1,2,3],[4,5,6],[7,8,9]]}
    ]
)
```

### Issue: Data dimensions don't match range
**Warning**: `Data dimensions (5x5) don't match range dimensions (3x3)`

**Explanation**: 
- If your data is larger than the specified range, the API will extend it
- This is expected behavior for `update_sheet_values`

**Solution**:
1. Match your data to your range size
2. Use `batch_update_values` for strict range control
3. Accept the auto-extension if it's acceptable for your use case

## Range Specification Issues

### Issue: Inconsistent range specification behavior
**Symptom**: Sometimes `Sheet1!A1:B10` works, sometimes `A1:B10` works, sometimes neither

**Best Practices**:
1. **ALWAYS include the sheet name** for clarity and consistency
2. Use quotes for sheet names with spaces
3. Be consistent with the parameter name (`range` vs `range_name`)

**Correct Format**:
```python
# ✅ RECOMMENDED - Always specify sheet
"Sheet1!A1:D10"

# ✅ Sheet name with spaces - use quotes
"'My Sheet'!A1:D10"

# ⚠️ Works but not recommended - uses first/active sheet
"A1:D10"
```

**Parameter Names**:
- Most functions use `range`
- `modify_sheet_values` uses `range_name` for backward compatibility
- Both work the same way

## Authentication Issues

### Issue: `start_google_auth` throws unexpected keyword argument error
**Error**: `start_auth_flow() got an unexpected keyword argument 'scopes'`

**Solution**:
1. This is a known issue in some environments
2. You can often skip this function and proceed directly
3. Authentication may be handled automatically

**Workaround**:
```python
# If start_google_auth fails, try proceeding directly:
# Just use your functions with your email address
list_spreadsheets(user_google_email="your-email@company.com")
```

## Table Styling Issues

### Issue: `apply_table_style` returns "'str' object is not callable"
**Cause**: Usually means the range is empty or invalid

**Prerequisites Check**:
1. ✅ Range MUST contain actual data (not empty cells)
2. ✅ Spreadsheet must be accessible with write permissions
3. ✅ Range must use valid A1 notation
4. ✅ The sheet specified must exist

**Solution**:
```python
# 1. First, ensure your range has data
update_sheet_values(..., range="Sheet1!A1:D10", values=your_data)

# 2. Then apply the style
apply_table_style(..., range="Sheet1!A1:D10", style="professional")
```

### Issue: No visible changes after applying table style
**Cause**: The range might be empty or the function silently failed

**Debugging Steps**:
1. Verify the range contains data using `read_sheet_values`
2. Check that the sheet name is correct
3. Ensure you have write permissions
4. Try a smaller range first

## General Best Practices

### 1. Range Specification
- **Always include sheet name**: `"Sheet1!A1:D10"` not just `"A1:D10"`
- **Use quotes for special names**: `"'My Sheet'!A1:D10"`
- **Be consistent**: Stick to one format throughout your code

### 2. Data Types
- **Font sizes**: Use integers (12, not "12" or "12pt")
- **Colors**: Use hex with # prefix (`"#FF0000"`)
- **Booleans**: Can use `True/False`, `1/0`, or `"true"/"false"`

### 3. Function Selection Guide

| Task | Best Function | Why |
|------|--------------|-----|
| Update single range | `update_sheet_values` | Simple and straightforward |
| Update multiple ranges | `batch_update_values` | Better performance, atomic |
| Need exact range control | `batch_update_values` | No auto-extension |
| Append to end of data | `append_sheet_values` | Finds last row automatically |
| Apply formatting | `format_cells` | Comprehensive options |
| Professional tables | `apply_table_style` | Quick predefined styles |
| Find data boundaries | `get_data_boundaries` | Dynamic range detection |
| Clear formatting | `reset_to_default_formatting` | Clean slate |

### 4. Error Recovery
If a function fails:
1. Check the prerequisites in the function's docstring
2. Verify your parameter types match the documentation
3. Ensure the sheet/range exists and has appropriate data
4. Try with a simpler example first
5. Check this troubleshooting guide

### 5. Performance Tips
- Use `batch_update_values` for multiple updates (single API call)
- Use `get_data_boundaries` to avoid processing empty cells
- Cache spreadsheet IDs to avoid repeated lookups
- Group related operations together

## Common Error Messages

| Error | Likely Cause | Solution |
|-------|--------------|----------|
| `"Sheet not found"` | Typo in sheet name | Check exact spelling, use quotes for spaces |
| `"Invalid range"` | Malformed A1 notation | Use format: `"SheetName!A1:D10"` |
| `"Permission denied"` | No write access | Check spreadsheet sharing settings |
| `"Values cannot be None"` | Empty values parameter | Provide data or use `clear_values=True` |
| `"'str' object is not callable"` | Empty range or internal error | Ensure range has data |
| `"Input validation error"` | Wrong parameter type | Check data types (int vs string) |

## Need More Help?

If you encounter an issue not covered here:
1. Check the function's docstring for detailed parameter information
2. Try the example code provided in the docstrings
3. Start with a minimal example and build up
4. Verify your authentication is working properly
5. Report persistent issues with full error messages and code samples