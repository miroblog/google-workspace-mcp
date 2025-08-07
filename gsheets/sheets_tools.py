"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import re
from typing import List, Optional, Dict, Union, Literal, Any


from auth.service_decorator import require_google_service
from core.server import server
from core.utils import handle_http_errors
from core.comments import create_comment_tools

# Configure module logger
logger = logging.getLogger(__name__)


@server.tool()
@handle_http_errors("list_spreadsheets", is_read_only=True, service_type="sheets")
@require_google_service("drive", "drive_read")
async def list_spreadsheets(
    service,
    user_google_email: str,
    max_results: int = 25,
) -> str:
    """
    Lists spreadsheets from Google Drive that the user has access to.

    Args:
        user_google_email (str): The user's Google email address. Required.
        max_results (int): Maximum number of spreadsheets to return. Defaults to 25.

    Returns:
        str: A formatted list of spreadsheet files (name, ID, modified time).
    """
    logger.info(f"[list_spreadsheets] Invoked. Email: '{user_google_email}'")

    files_response = await asyncio.to_thread(
        service.files()
        .list(
            q="mimeType='application/vnd.google-apps.spreadsheet'",
            pageSize=max_results,
            fields="files(id,name,modifiedTime,webViewLink)",
            orderBy="modifiedTime desc",
        )
        .execute
    )

    files = files_response.get("files", [])
    if not files:
        return f"No spreadsheets found for {user_google_email}."

    spreadsheets_list = [
        f'- "{file["name"]}" (ID: {file["id"]}) | Modified: {file.get("modifiedTime", "Unknown")} | Link: {file.get("webViewLink", "No link")}'
        for file in files
    ]

    text_output = (
        f"Successfully listed {len(files)} spreadsheets for {user_google_email}:\n"
        + "\n".join(spreadsheets_list)
    )

    logger.info(
        f"Successfully listed {len(files)} spreadsheets for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("get_spreadsheet_info", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def get_spreadsheet_info(
    service,
    user_google_email: str,
    spreadsheet_id: str,
) -> str:
    """
    Gets information about a specific spreadsheet including its sheets.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet to get info for. Required.

    Returns:
        str: Formatted spreadsheet information including title and sheets list.
    """
    logger.info(
        f"[get_spreadsheet_info] Invoked. Email: '{user_google_email}', Spreadsheet ID: {spreadsheet_id}"
    )

    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    title = spreadsheet.get("properties", {}).get("title", "Unknown")
    sheets = spreadsheet.get("sheets", [])

    sheets_info = []
    for sheet in sheets:
        sheet_props = sheet.get("properties", {})
        sheet_name = sheet_props.get("title", "Unknown")
        sheet_id = sheet_props.get("sheetId", "Unknown")
        grid_props = sheet_props.get("gridProperties", {})
        rows = grid_props.get("rowCount", "Unknown")
        cols = grid_props.get("columnCount", "Unknown")

        sheets_info.append(f'  - "{sheet_name}" (ID: {sheet_id}) | Size: {rows}x{cols}')

    text_output = (
        f'Spreadsheet: "{title}" (ID: {spreadsheet_id})\n'
        f"Sheets ({len(sheets)}):\n" + "\n".join(sheets_info)
        if sheets_info
        else "  No sheets found"
    )

    logger.info(
        f"Successfully retrieved info for spreadsheet {spreadsheet_id} for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("read_sheet_values", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str = "A1:Z1000",
) -> str:
    """
    Reads values from a specific range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to read (e.g., "Sheet1!A1:D10", "A1:D10"). Defaults to "A1:Z1000".

    Returns:
        str: The formatted values from the specified range.
    """
    logger.info(
        f"[read_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute
    )

    values = result.get("values", [])
    if not values:
        return f"No data found in range '{range_name}' for {user_google_email}."

    # Format the output as a readable table
    formatted_rows = []
    for i, row in enumerate(values, 1):
        # Pad row with empty strings to show structure
        padded_row = row + [""] * max(0, len(values[0]) - len(row)) if values else row
        formatted_rows.append(f"Row {i:2d}: {padded_row}")

    text_output = (
        f"Successfully read {len(values)} rows from range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}:\n"
        + "\n".join(formatted_rows[:50])  # Limit to first 50 rows for readability
        + (f"\n... and {len(values) - 50} more rows" if len(values) > 50 else "")
    )

    logger.info(f"Successfully read {len(values)} rows for {user_google_email}.")
    return text_output


# ============================================================================
# HELPER FUNCTIONS FOR VALUE VALIDATION
# ============================================================================


def _validate_2d_array(values: Any) -> List[List[Any]]:
    """
    Validate and convert input to 2D array format for Google Sheets API.

    Args:
        values: Input values in various formats (string, 1D list, 2D list)

    Returns:
        List[List[Any]]: Properly formatted 2D array

    Raises:
        ValueError: If values format is invalid
    """
    if values is None:
        raise ValueError("Values cannot be None")

    # Handle single value (string, number, etc.)
    if not isinstance(values, list):
        return [[values]]

    # Empty list check
    if len(values) == 0:
        raise ValueError("Values cannot be an empty list")

    # Check if it's already a 2D array
    if isinstance(values[0], list):
        # Validate that all elements are lists
        for i, row in enumerate(values):
            if not isinstance(row, list):
                raise ValueError(f"Row {i} is not a list: {type(row)}")
        return values

    # Convert 1D array to 2D array (single row)
    return [values]


def _analyze_range(range_str: str) -> Optional[Dict]:
    """
    Analyze an A1 notation range to extract information.
    
    Args:
        range_str: A1 notation range (e.g., "Sheet1!A1:D10")
    
    Returns:
        Dict with range information or None if cannot parse
    """
    if not range_str:
        return None
    
    try:
        # Parse sheet and range
        if "!" in range_str:
            sheet, cell_range = range_str.split("!", 1)
        else:
            sheet = None
            cell_range = range_str
        
        # Check if it has bounds (contains :)
        if ":" not in cell_range:
            return {
                "sheet": sheet,
                "has_bounds": False,
                "start": cell_range,
                "rows": 1,
                "cols": 1
            }
        
        # Parse start and end
        parts = cell_range.split(":")
        if len(parts) != 2:
            return None
            
        start, end = parts
        
        # Extract row and column info using regex
        import re
        start_match = re.match(r"([A-Z]+)(\d+)", start)
        end_match = re.match(r"([A-Z]+)(\d+)", end)
        
        if not start_match or not end_match:
            return {
                "sheet": sheet,
                "has_bounds": True,
                "start": start,
                "end": end,
            }
        
        start_col_letter = start_match.group(1)
        start_row = int(start_match.group(2))
        end_col_letter = end_match.group(1)
        end_row = int(end_match.group(2))
        
        # Calculate dimensions
        start_col = _column_letter_to_index(start_col_letter)
        end_col = _column_letter_to_index(end_col_letter)
        
        rows = end_row - start_row + 1
        cols = end_col - start_col + 1
        
        return {
            "sheet": sheet,
            "has_bounds": True,
            "start": start,
            "end": end,
            "rows": rows,
            "cols": cols,
            "start_row": start_row,
            "start_col": start_col,
            "end_row": end_row,
            "end_col": end_col
        }
    except Exception:
        return None


def _parse_boolean(value: Optional[Union[bool, int, str]]) -> Optional[bool]:
    """
    Parse various boolean representations to actual boolean value.
    
    Args:
        value: Boolean value as bool, int (0/1), or string ("true"/"false", "1"/"0")
    
    Returns:
        Optional[bool]: Parsed boolean value or None if input is None
    
    Examples:
        _parse_boolean(True) -> True
        _parse_boolean(1) -> True
        _parse_boolean("true") -> True
        _parse_boolean("1") -> True
        _parse_boolean(0) -> False
        _parse_boolean("false") -> False
        _parse_boolean(None) -> None
    """
    if value is None:
        return None
    
    # Already a boolean
    if isinstance(value, bool):
        return value
    
    # Integer: 0 = False, non-zero = True
    if isinstance(value, int):
        return bool(value)
    
    # String representations
    if isinstance(value, str):
        lower_val = value.lower().strip()
        if lower_val in ("true", "1", "yes", "on"):
            return True
        elif lower_val in ("false", "0", "no", "off"):
            return False
        else:
            # Try to parse as number
            try:
                return bool(int(value))
            except (ValueError, TypeError):
                pass
    
    # Default to treating truthy values as True
    return bool(value)


@server.tool()
@handle_http_errors("modify_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def modify_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    values: Any = None,  # Changed to Any to avoid FastMCP validation issues
    value_input_option: str = "USER_ENTERED",
    clear_values: bool = False,
) -> str:
    """
    Modifies values in a specific range of a Google Sheet - can write, update, or clear values.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to modify (e.g., "Sheet1!A1:D10", "A1:D10"). Required.
        values (Any): Values to write/update. Can be a single value, 1D array, or 2D array. Required unless clear_values=True.
        value_input_option (str): How to interpret input values ("RAW" or "USER_ENTERED"). Defaults to "USER_ENTERED".
        clear_values (bool): If True, clears the range instead of writing values. Defaults to False.

    Returns:
        str: Confirmation message of the successful modification operation.
    """
    operation = "clear" if clear_values else "write"
    logger.info(
        f"[modify_sheet_values] Invoked. Operation: {operation}, Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    if not clear_values and values is None:
        raise Exception(
            "Either 'values' must be provided or 'clear_values' must be True."
        )

    # Validate and convert values to proper 2D array format
    if not clear_values and values is not None:
        try:
            values = _validate_2d_array(values)
        except ValueError as e:
            raise Exception(f"Invalid values format: {e}")

    if clear_values:
        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=range_name)
            .execute
        )

        cleared_range = result.get("clearedRange", range_name)
        text_output = f"Successfully cleared range '{cleared_range}' in spreadsheet {spreadsheet_id} for {user_google_email}."
        logger.info(
            f"Successfully cleared range '{cleared_range}' for {user_google_email}."
        )
    else:
        body = {"values": values}

        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption=value_input_option,
                body=body,
            )
            .execute
        )

        updated_cells = result.get("updatedCells", 0)
        updated_rows = result.get("updatedRows", 0)
        updated_columns = result.get("updatedColumns", 0)

        text_output = (
            f"Successfully updated range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}. "
            f"Updated: {updated_cells} cells, {updated_rows} rows, {updated_columns} columns."
        )
        logger.info(
            f"Successfully updated {updated_cells} cells for {user_google_email}."
        )

    return text_output


@server.tool()
@handle_http_errors("update_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def update_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range: str,
    values: Any,
    value_input_option: str = "USER_ENTERED",
    include_values_in_response: bool = False,
) -> str:
    """
    Simplified function for updating values in a Google Sheet range.
    
    IMPORTANT: This function will update the exact range specified. The API may extend
    the range if your data exceeds the specified bounds. To ensure exact range updates,
    make sure your data dimensions match the range dimensions.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): The A1 notation range to update (e.g., "Sheet1!A1:D10"). Required.
                    NOTE: The range determines where data starts, but may be extended if values exceed it.
        values (Any): Values to write. Can be:
                     - Single value: "100" or 42
                     - 1D array (single row): ["A", "B", "C"]
                     - 2D array: [["A", "B"], ["C", "D"]]
        value_input_option (str): How to interpret values:
                                 - "RAW": Values are stored as-is
                                 - "USER_ENTERED": Values are parsed (formulas, numbers, dates)
                                 Defaults to "USER_ENTERED".
        include_values_in_response (bool): Include the updated values in response. Defaults to False.

    Returns:
        str: Confirmation message with update statistics.

    Examples:
        # Single cell - updates only A1
        update_sheet_values(..., range="A1", values="Hello")
        
        # Single row - updates A1:C1
        update_sheet_values(..., range="A1:C1", values=["One", "Two", "Three"])
        
        # Multiple rows - updates A1:B2
        update_sheet_values(..., range="A1:B2", values=[["A", "B"], ["C", "D"]])
        
        # Formula entry - use USER_ENTERED
        update_sheet_values(..., range="D1", values="=SUM(A1:C1)", value_input_option="USER_ENTERED")
        
        # Exact text entry - use RAW
        update_sheet_values(..., range="E1", values="=Not a formula", value_input_option="RAW")

    Troubleshooting:
        - Range Mismatch: If your data has more rows/columns than the range, the API extends the range
        - Empty Cells: To clear cells, pass empty strings: [["", "", ""]]
        - Formulas: Use value_input_option="USER_ENTERED" for formulas to be evaluated
        - Exact Text: Use value_input_option="RAW" to store text exactly as provided
        - Data Validation: The range must exist; create sheets first if needed
    """
    logger.info(
        f"[update_sheet_values] Updating range {range} in spreadsheet {spreadsheet_id}"
    )

    # Validate and format values
    try:
        validated_values = _validate_2d_array(values)
    except ValueError as e:
        raise Exception(f"Invalid values format: {e}")
    
    # Analyze the range to provide better feedback
    range_info = _analyze_range(range)
    data_rows = len(validated_values)
    data_cols = len(validated_values[0]) if validated_values else 0
    
    # Log if data dimensions don't match range (when range has defined bounds)
    if range_info and range_info.get("has_bounds"):
        expected_rows = range_info.get("rows", 0)
        expected_cols = range_info.get("cols", 0)
        if data_rows != expected_rows or data_cols != expected_cols:
            logger.warning(
                f"Data dimensions ({data_rows}x{data_cols}) don't match range dimensions "
                f"({expected_rows}x{expected_cols}). API may adjust the range."
            )

    # Prepare the request body
    body = {
        "values": validated_values,
    }
    
    # Add response options
    response_value_render_option = "FORMATTED_VALUE" if include_values_in_response else None
    response_date_time_render_option = "FORMATTED_STRING" if include_values_in_response else None

    # Execute the update with proper parameters
    params = {
        "spreadsheetId": spreadsheet_id,
        "range": range,
        "valueInputOption": value_input_option,
        "body": body,
    }
    
    if include_values_in_response:
        params["responseValueRenderOption"] = response_value_render_option
        params["responseDateTimeRenderOption"] = response_date_time_render_option
    
    result = await asyncio.to_thread(
        lambda: service.spreadsheets().values().update(**params).execute()
    )

    # Extract update statistics
    updated_cells = result.get("updatedCells", 0)
    updated_rows = result.get("updatedRows", 0)
    updated_columns = result.get("updatedColumns", 0)
    updated_range = result.get("updatedRange", range)

    logger.info(f"Successfully updated {updated_cells} cells in range {updated_range}")
    
    response = (
        f"Successfully updated range '{updated_range}' in spreadsheet {spreadsheet_id}.\n"
        f"Statistics: {updated_cells} cells, {updated_rows} rows, {updated_columns} columns updated."
    )
    
    # Add actual vs requested range info if different
    if updated_range != range:
        response += f"\nNote: Actual updated range '{updated_range}' differs from requested '{range}'"
    
    return response


@server.tool()
@handle_http_errors("batch_update_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def batch_update_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    updates: List[Dict[str, Any]],
    value_input_option: str = "USER_ENTERED",
) -> str:
    """
    Update multiple ranges in a spreadsheet with a single API call for better performance.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        updates (List[Dict]): List of update operations, each containing:
                              - "range": The A1 notation range
                              - "values": The values to write (any format)
        value_input_option (str): How to interpret values - "RAW" or "USER_ENTERED". Defaults to "USER_ENTERED".

    Returns:
        str: Summary of all updates performed.

    Example:
        batch_update_values(..., updates=[
            {"range": "Sheet1!A1:B2", "values": [["A", "B"], ["C", "D"]]},
            {"range": "Sheet2!A1", "values": "Single Value"},
            {"range": "Sheet1!D1:F1", "values": [1, 2, 3]}
        ])
    """
    logger.info(
        f"[batch_update_values] Performing {len(updates)} updates in spreadsheet {spreadsheet_id}"
    )

    # Validate and prepare all updates
    data = []
    for update in updates:
        if "range" not in update or "values" not in update:
            raise Exception("Each update must have 'range' and 'values' keys")

        try:
            validated_values = _validate_2d_array(update["values"])
        except ValueError as e:
            raise Exception(f"Invalid values format for range {update['range']}: {e}")

        data.append({"range": update["range"], "values": validated_values})

    # Prepare the batch update request
    body = {"valueInputOption": value_input_option, "data": data}

    # Execute the batch update
    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute
    )

    # Process results
    total_updated_cells = result.get("totalUpdatedCells", 0)
    total_updated_rows = result.get("totalUpdatedRows", 0)
    total_updated_columns = result.get("totalUpdatedColumns", 0)
    total_updated_sheets = result.get("totalUpdatedSheets", 0)

    responses = result.get("responses", [])
    update_details = []
    for i, response in enumerate(responses):
        range_updated = response.get("updatedRange", data[i]["range"])
        cells = response.get("updatedCells", 0)
        update_details.append(f"  - {range_updated}: {cells} cells")

    logger.info(
        f"Successfully updated {total_updated_cells} cells across {len(responses)} ranges"
    )

    details_str = (
        "\n".join(update_details) if update_details else "No details available"
    )

    return (
        f"Successfully performed {len(responses)} batch updates in spreadsheet {spreadsheet_id}.\n"
        f"Total statistics: {total_updated_cells} cells, {total_updated_rows} rows, "
        f"{total_updated_columns} columns across {total_updated_sheets} sheets.\n"
        f"Updates:\n{details_str}"
    )


@server.tool()
@handle_http_errors("append_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def append_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range: str,
    values: Any,
    value_input_option: str = "USER_ENTERED",
    insert_data_option: str = "INSERT_ROWS",
) -> str:
    """
    Append values to the end of existing data in a spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): The A1 notation range to search for existing data (e.g., "Sheet1!A:D"). Required.
        values (Any): Values to append. Can be single value, 1D array, or 2D array.
        value_input_option (str): How to interpret values - "RAW" or "USER_ENTERED". Defaults to "USER_ENTERED".
        insert_data_option (str): How to insert data - "INSERT_ROWS" or "OVERWRITE". Defaults to "INSERT_ROWS".

    Returns:
        str: Confirmation message with the range where data was appended.

    Examples:
        # Append a single row
        append_sheet_values(..., range="Sheet1!A:C", values=["New", "Row", "Data"])

        # Append multiple rows
        append_sheet_values(..., range="Sheet1!A:B", values=[["Row1A", "Row1B"], ["Row2A", "Row2B"]])
    """
    logger.info(
        f"[append_sheet_values] Appending values to range {range} in spreadsheet {spreadsheet_id}"
    )

    # Validate and format values
    try:
        validated_values = _validate_2d_array(values)
    except ValueError as e:
        raise Exception(f"Invalid values format: {e}")

    # Prepare the request body
    body = {"values": validated_values}

    # Execute the append
    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=range,
            valueInputOption=value_input_option,
            insertDataOption=insert_data_option,
            body=body,
        )
        .execute
    )

    # Extract results
    updates = result.get("updates", {})
    updated_range = updates.get("updatedRange", "Unknown")
    updated_rows = updates.get("updatedRows", 0)
    updated_columns = updates.get("updatedColumns", 0)
    updated_cells = updates.get("updatedCells", 0)

    logger.info(f"Successfully appended {updated_rows} rows to {updated_range}")

    return (
        f"Successfully appended data to spreadsheet {spreadsheet_id}.\n"
        f"Appended to range: {updated_range}\n"
        f"Statistics: {updated_cells} cells, {updated_rows} rows, {updated_columns} columns added."
    )


@server.tool()
@handle_http_errors("create_spreadsheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def create_spreadsheet(
    service,
    user_google_email: str,
    title: str,
    sheet_names: Optional[List[str]] = None,
) -> str:
    """
    Creates a new Google Spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title of the new spreadsheet. Required.
        sheet_names (Optional[List[str]]): List of sheet names to create. If not provided, creates one sheet with default name.

    Returns:
        str: Information about the newly created spreadsheet including ID and URL.
    """
    logger.info(
        f"[create_spreadsheet] Invoked. Email: '{user_google_email}', Title: {title}"
    )

    spreadsheet_body = {"properties": {"title": title}}

    if sheet_names:
        spreadsheet_body["sheets"] = [
            {"properties": {"title": sheet_name}} for sheet_name in sheet_names
        ]

    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().create(body=spreadsheet_body).execute
    )

    spreadsheet_id = spreadsheet.get("spreadsheetId")
    spreadsheet_url = spreadsheet.get("spreadsheetUrl")

    text_output = (
        f"Successfully created spreadsheet '{title}' for {user_google_email}. "
        f"ID: {spreadsheet_id} | URL: {spreadsheet_url}"
    )

    logger.info(
        f"Successfully created spreadsheet for {user_google_email}. ID: {spreadsheet_id}"
    )
    return text_output


@server.tool()
@handle_http_errors("create_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def create_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: str,
) -> str:
    """
    Creates a new sheet within an existing spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (str): The name of the new sheet. Required.

    Returns:
        str: Confirmation message of the successful sheet creation.
    """
    logger.info(
        f"[create_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet: {sheet_name}"
    )

    request_body = {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}

    response = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    sheet_id = response["replies"][0]["addSheet"]["properties"]["sheetId"]

    text_output = f"Successfully created sheet '{sheet_name}' (ID: {sheet_id}) in spreadsheet {spreadsheet_id} for {user_google_email}."

    logger.info(
        f"Successfully created sheet for {user_google_email}. Sheet ID: {sheet_id}"
    )
    return text_output


# Create comment management tools for sheets
_comment_tools = create_comment_tools("spreadsheet", "spreadsheet_id")

# Extract and register the functions
read_sheet_comments = _comment_tools["read_comments"]
create_sheet_comment = _comment_tools["create_comment"]
reply_to_sheet_comment = _comment_tools["reply_to_comment"]
resolve_sheet_comment = _comment_tools["resolve_comment"]

# ============================================================================
# ENHANCED FEATURES: CELL FORMATTING & STYLING
# ============================================================================


@server.tool()
@handle_http_errors("format_cells", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def format_cells(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range: str,
    background_color: Optional[str] = None,
    font_color: Optional[str] = None,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    bold: Optional[Union[bool, int, str]] = None,  # Accept bool, int (0/1), or str ("true"/"false")
    italic: Optional[Union[bool, int, str]] = None,  # Accept bool, int (0/1), or str ("true"/"false")
    underline: Optional[Union[bool, int, str]] = None,  # Accept bool, int (0/1), or str ("true"/"false")
    strikethrough: Optional[Union[bool, int, str]] = None,  # Accept bool, int (0/1), or str ("true"/"false")
    horizontal_alignment: Optional[Literal["LEFT", "CENTER", "RIGHT"]] = None,
    vertical_alignment: Optional[Literal["TOP", "MIDDLE", "BOTTOM"]] = None,
    text_wrap: Optional[Literal["OVERFLOW_CELL", "WRAP", "CLIP"]] = None,
    number_format: Optional[
        Literal[
            "NUMBER", "CURRENCY", "PERCENT", "DATE", "TIME", "DATETIME", "SCIENTIFIC"
        ]
    ] = None,
    number_format_pattern: Optional[str] = None,
    border_style: Optional[
        Literal["DOTTED", "DASHED", "SOLID", "SOLID_MEDIUM", "SOLID_THICK", "DOUBLE"]
    ] = None,
    border_color: Optional[str] = None,
) -> str:
    """
    Formats cells in a Google Sheet with comprehensive styling options.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): The range to format (e.g., "Sheet1!A1:D10"). Required.
        background_color (Optional[str]): Hex color for background (e.g., "#FF0000").
        font_color (Optional[str]): Hex color for text (e.g., "#000000").
        font_size (Optional[int]): Font size in points (e.g., 10, 12, 14).
        font_family (Optional[str]): Font family name (e.g., "Arial", "Times New Roman", "Roboto").
        bold (Optional[Union[bool, int, str]]): Make text bold. Accepts: True/False, 1/0, "true"/"false".
        italic (Optional[Union[bool, int, str]]): Make text italic. Accepts: True/False, 1/0, "true"/"false".
        underline (Optional[Union[bool, int, str]]): Underline text. Accepts: True/False, 1/0, "true"/"false".
        strikethrough (Optional[Union[bool, int, str]]): Strikethrough text. Accepts: True/False, 1/0, "true"/"false".
        horizontal_alignment (Optional[str]): "LEFT", "CENTER", or "RIGHT".
        vertical_alignment (Optional[str]): "TOP", "MIDDLE", or "BOTTOM".
        text_wrap (Optional[str]): "OVERFLOW_CELL", "WRAP", or "CLIP".
        number_format (Optional[str]): "NUMBER", "CURRENCY", "PERCENT", "DATE", "TIME", "DATETIME", or "SCIENTIFIC".
        number_format_pattern (Optional[str]): Custom format pattern (e.g., "#,##0.00", "MM/DD/YYYY").
        border_style (Optional[str]): "DOTTED", "DASHED", "SOLID", "SOLID_MEDIUM", "SOLID_THICK", or "DOUBLE".
        border_color (Optional[str]): Hex color for borders (e.g., "#000000").

    Returns:
        str: Confirmation message of successful formatting.

    Examples:
        # Basic text formatting
        format_cells(..., range="A1:C3", bold=True, font_size=14, font_color="#0000FF")
        
        # Apply background color and borders
        format_cells(..., range="A1:D10", background_color="#FFFF00", 
                    border_style="SOLID", border_color="#000000")
        
        # Format numbers as currency
        format_cells(..., range="B2:B10", number_format="CURRENCY")
        
        # Custom number format with thousands separator
        format_cells(..., range="C1:C20", number_format_pattern="#,##0.00")
        
        # Center align with text wrapping
        format_cells(..., range="A1:Z1", horizontal_alignment="CENTER", 
                    text_wrap="WRAP", bold=True)
        
        # Boolean parameters work with multiple formats
        format_cells(..., range="A1", bold=1, italic="true", underline=0)  # bold=True, italic=True, underline=False

    Troubleshooting:
        - Boolean values: Can be passed as bool (True/False), int (1/0), or string ("true"/"false", "1"/"0")
        - Colors: Must be in hex format with # prefix (e.g., "#FF0000" for red)
        - Range: Must include sheet name if not the first sheet (e.g., "Sheet2!A1:B10")
        - Number formats: Use predefined types or custom patterns following Google Sheets format syntax
    """
    logger.info(
        f"[format_cells] Invoked for {user_google_email}, spreadsheet: {spreadsheet_id}, range: {range}"
    )

    # Parse the range to get sheet ID and grid range
    sheet_name, cell_range = _parse_range(range)

    # Get spreadsheet metadata to find sheet ID
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    sheet_id = _get_sheet_id_by_name(spreadsheet, sheet_name)
    grid_range = _convert_a1_to_grid_range(cell_range, sheet_id)

    # Build the cell format
    cell_format = {}

    # Background color
    if background_color:
        cell_format["backgroundColor"] = _hex_to_rgb_dict(background_color)

    # Text format
    text_format = {}
    if font_color:
        text_format["foregroundColor"] = _hex_to_rgb_dict(font_color)
    if font_size:
        text_format["fontSize"] = font_size
    if font_family:
        text_format["fontFamily"] = font_family
    
    # Parse boolean values with flexible input handling
    parsed_bold = _parse_boolean(bold)
    if parsed_bold is not None:
        text_format["bold"] = parsed_bold
    
    parsed_italic = _parse_boolean(italic)
    if parsed_italic is not None:
        text_format["italic"] = parsed_italic
    
    parsed_underline = _parse_boolean(underline)
    if parsed_underline is not None:
        text_format["underline"] = parsed_underline
    
    parsed_strikethrough = _parse_boolean(strikethrough)
    if parsed_strikethrough is not None:
        text_format["strikethrough"] = parsed_strikethrough

    if text_format:
        cell_format["textFormat"] = text_format

    # Alignment
    if horizontal_alignment:
        cell_format["horizontalAlignment"] = horizontal_alignment
    if vertical_alignment:
        cell_format["verticalAlignment"] = vertical_alignment

    # Text wrapping
    if text_wrap:
        cell_format["wrapStrategy"] = text_wrap

    # Number format
    if number_format or number_format_pattern:
        number_format_obj = {}
        if number_format:
            number_format_obj["type"] = number_format
        if number_format_pattern:
            number_format_obj["pattern"] = number_format_pattern
        cell_format["numberFormat"] = number_format_obj

    # Borders
    if border_style or border_color:
        border = {
            "style": border_style or "SOLID",
            "color": _hex_to_rgb_dict(border_color)
            if border_color
            else {"red": 0, "green": 0, "blue": 0},
        }
        cell_format["borders"] = {
            "top": border,
            "bottom": border,
            "left": border,
            "right": border,
        }

    # Build the request
    request = {
        "repeatCell": {
            "range": grid_range,
            "cell": {"userEnteredFormat": cell_format},
            "fields": _get_update_fields_from_format(cell_format),
        }
    }

    # Execute the batch update
    body = {"requests": [request]}
    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute
    )

    logger.info(
        f"Successfully formatted cells in range {range} for {user_google_email}"
    )
    return f"Successfully formatted cells in range '{range}' with specified styling for {user_google_email}."


# ============================================================================
# CONDITIONAL FORMATTING
# ============================================================================


@server.tool()
@handle_http_errors("add_conditional_format_rule", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def add_conditional_format_rule(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    ranges: List[str],
    rule_type: Literal[
        "custom_formula",
        "number_greater",
        "number_less",
        "number_between",
        "text_contains",
        "text_starts_with",
        "text_ends_with",
        "date_before",
        "date_after",
        "gradient",
    ],
    formula: Optional[str] = None,
    value: Optional[Union[str, float]] = None,
    min_value: Optional[Union[str, float]] = None,
    max_value: Optional[Union[str, float]] = None,
    background_color: Optional[str] = None,
    font_color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    strikethrough: Optional[bool] = None,
    gradient_min_color: Optional[str] = None,
    gradient_mid_color: Optional[str] = None,
    gradient_max_color: Optional[str] = None,
    gradient_min_value: Optional[float] = None,
    gradient_mid_value: Optional[float] = None,
    gradient_max_value: Optional[float] = None,
) -> str:
    """
    Adds a conditional formatting rule to specified ranges in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        ranges (List[str]): List of ranges to apply the rule to (e.g., ["A1:D10", "F1:F20"]). Required.
        rule_type (str): Type of conditional formatting rule. Required.
        formula (Optional[str]): Custom formula for "custom_formula" type (e.g., "=$B1>100").
        value (Optional[Union[str, float]]): Value for comparison rules.
        min_value (Optional[Union[str, float]]): Minimum value for "number_between" rule.
        max_value (Optional[Union[str, float]]): Maximum value for "number_between" rule.
        background_color (Optional[str]): Hex color for cell background when condition is met.
        font_color (Optional[str]): Hex color for text when condition is met.
        bold (Optional[bool]): Make text bold when condition is met.
        italic (Optional[bool]): Make text italic when condition is met.
        underline (Optional[bool]): Underline text when condition is met.
        strikethrough (Optional[bool]): Strikethrough text when condition is met.
        gradient_min_color (Optional[str]): Color for minimum value in gradient rule.
        gradient_mid_color (Optional[str]): Color for midpoint value in gradient rule.
        gradient_max_color (Optional[str]): Color for maximum value in gradient rule.
        gradient_min_value (Optional[float]): Minimum value for gradient scale.
        gradient_mid_value (Optional[float]): Midpoint value for gradient scale.
        gradient_max_value (Optional[float]): Maximum value for gradient scale.

    Returns:
        str: Confirmation message with the rule ID.
    """
    logger.info(
        f"[add_conditional_format_rule] Invoked for {user_google_email}, spreadsheet: {spreadsheet_id}"
    )

    # Get spreadsheet metadata
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    # Convert ranges to grid ranges
    grid_ranges = []
    for range_str in ranges:
        sheet_name, cell_range = _parse_range(range_str)
        sheet_id = _get_sheet_id_by_name(spreadsheet, sheet_name)
        grid_ranges.append(_convert_a1_to_grid_range(cell_range, sheet_id))

    # Build the conditional format rule
    rule = {"ranges": grid_ranges}

    if rule_type == "gradient":
        # Gradient rule
        gradient_rule = {}

        if gradient_min_value is not None and gradient_max_value is not None:
            gradient_rule["minpoint"] = {
                "type": "NUMBER",
                "value": str(gradient_min_value),
                "color": _hex_to_rgb_dict(gradient_min_color or "#FFFFFF"),
            }
            gradient_rule["maxpoint"] = {
                "type": "NUMBER",
                "value": str(gradient_max_value),
                "color": _hex_to_rgb_dict(gradient_max_color or "#FF0000"),
            }

            if gradient_mid_value is not None:
                gradient_rule["midpoint"] = {
                    "type": "NUMBER",
                    "value": str(gradient_mid_value),
                    "color": _hex_to_rgb_dict(gradient_mid_color or "#FFFF00"),
                }
        else:
            # Use percentile-based gradient
            gradient_rule["minpoint"] = {
                "type": "MIN",
                "color": _hex_to_rgb_dict(gradient_min_color or "#FFFFFF"),
            }
            gradient_rule["maxpoint"] = {
                "type": "MAX",
                "color": _hex_to_rgb_dict(gradient_max_color or "#FF0000"),
            }

        rule["gradientRule"] = gradient_rule
    else:
        # Boolean rule
        boolean_rule = {}
        condition = {
            "type": "CUSTOM_FORMULA"
            if rule_type == "custom_formula"
            else rule_type.upper()
        }

        # Set condition values based on rule type
        if rule_type == "custom_formula":
            if not formula:
                raise ValueError("Formula is required for custom_formula rule type")
            condition["values"] = [{"userEnteredValue": formula}]
        elif rule_type in ["number_greater", "number_less"]:
            if value is None:
                raise ValueError(f"Value is required for {rule_type} rule type")
            condition["values"] = [{"userEnteredValue": str(value)}]
        elif rule_type == "number_between":
            if min_value is None or max_value is None:
                raise ValueError(
                    "Both min_value and max_value are required for number_between rule type"
                )
            condition["values"] = [
                {"userEnteredValue": str(min_value)},
                {"userEnteredValue": str(max_value)},
            ]
        elif rule_type in ["text_contains", "text_starts_with", "text_ends_with"]:
            if not value:
                raise ValueError(f"Value is required for {rule_type} rule type")
            condition["values"] = [{"userEnteredValue": str(value)}]
        elif rule_type in ["date_before", "date_after"]:
            if not value:
                raise ValueError(f"Value is required for {rule_type} rule type")
            condition["values"] = [{"relativeDate": value}]

        boolean_rule["condition"] = condition

        # Set format when condition is true
        format_obj = {}
        if background_color:
            format_obj["backgroundColor"] = _hex_to_rgb_dict(background_color)

        text_format = {}
        if font_color:
            text_format["foregroundColor"] = _hex_to_rgb_dict(font_color)
        if bold is not None:
            text_format["bold"] = bold
        if italic is not None:
            text_format["italic"] = italic
        if underline is not None:
            text_format["underline"] = underline
        if strikethrough is not None:
            text_format["strikethrough"] = strikethrough

        if text_format:
            format_obj["textFormat"] = text_format

        if format_obj:
            boolean_rule["format"] = format_obj

        rule["booleanRule"] = boolean_rule

    # Build and execute the request
    request = {"addConditionalFormatRule": {"rule": rule}}
    body = {"requests": [request]}

    result = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute
    )

    # Get the rule ID from the response
    replies = result.get("replies", [])
    rule_id = None
    if replies and "addConditionalFormatRule" in replies[0]:
        rule_id = replies[0]["addConditionalFormatRule"]["rule"]["index"]

    logger.info(
        f"Successfully added conditional format rule (ID: {rule_id}) for {user_google_email}"
    )
    return f"Successfully added conditional formatting rule (ID: {rule_id}) to ranges {ranges} for {user_google_email}."


@server.tool()
@handle_http_errors(
    "list_conditional_format_rules", is_read_only=True, service_type="sheets"
)
@require_google_service("sheets", "sheets_read")
async def list_conditional_format_rules(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: Optional[str] = None,
) -> str:
    """
    Lists all conditional formatting rules in a spreadsheet or specific sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (Optional[str]): Name of specific sheet to list rules for. If None, lists all.

    Returns:
        str: Formatted list of conditional formatting rules.
    """
    logger.info(
        f"[list_conditional_format_rules] Invoked for {user_google_email}, spreadsheet: {spreadsheet_id}"
    )

    # Get spreadsheet with conditional format rules
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, includeGridData=False)
        .execute
    )

    rules_output = []
    sheets = spreadsheet.get("sheets", [])

    for sheet in sheets:
        sheet_props = sheet.get("properties", {})
        current_sheet_name = sheet_props.get("title", "Unknown")

        # Skip if specific sheet requested and this isn't it
        if sheet_name and current_sheet_name != sheet_name:
            continue

        conditional_formats = sheet.get("conditionalFormats", [])

        if conditional_formats:
            rules_output.append(f"\n**Sheet: {current_sheet_name}**")
            for i, rule in enumerate(conditional_formats):
                rule_index = rule.get("index", i)
                ranges = rule.get("ranges", [])
                range_strs = [_grid_range_to_a1(r, current_sheet_name) for r in ranges]

                if "booleanRule" in rule:
                    rule_type = "Boolean Rule"
                    condition = rule["booleanRule"].get("condition", {})
                    condition_type = condition.get("type", "Unknown")
                    values = condition.get("values", [])
                    # format_info = rule["booleanRule"].get("format", {})
                elif "gradientRule" in rule:
                    rule_type = "Gradient Rule"
                    condition_type = "Gradient"
                    values = []
                    # format_info = rule["gradientRule"]
                else:
                    rule_type = "Unknown Rule"
                    condition_type = "Unknown"
                    values = []
                    # format_info = {}

                rules_output.append(f"  Rule {rule_index}: {rule_type}")
                rules_output.append(f"    Ranges: {', '.join(range_strs)}")
                rules_output.append(f"    Condition: {condition_type}")
                if values:
                    value_strs = [
                        v.get("userEnteredValue", v.get("relativeDate", ""))
                        for v in values
                    ]
                    rules_output.append(f"    Values: {', '.join(value_strs)}")

    if not rules_output:
        return f"No conditional formatting rules found in spreadsheet {spreadsheet_id} for {user_google_email}."

    return f"Conditional formatting rules for {user_google_email}:\n" + "\n".join(
        rules_output
    )


@server.tool()
@handle_http_errors("delete_conditional_format_rule", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def delete_conditional_format_rule(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    rule_index: int,
    sheet_name: str,
) -> str:
    """
    Deletes a conditional formatting rule from a sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        rule_index (int): The index of the rule to delete. Required.
        sheet_name (str): Name of the sheet containing the rule. Required.

    Returns:
        str: Confirmation message of deletion.
    """
    logger.info(
        f"[delete_conditional_format_rule] Deleting rule {rule_index} from sheet {sheet_name}"
    )

    # Get spreadsheet metadata
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    sheet_id = _get_sheet_id_by_name(spreadsheet, sheet_name)

    # Build the delete request
    request = {
        "deleteConditionalFormatRule": {"sheetId": sheet_id, "index": rule_index}
    }

    body = {"requests": [request]}

    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute
    )

    logger.info(
        f"Successfully deleted conditional format rule {rule_index} for {user_google_email}"
    )
    return f"Successfully deleted conditional formatting rule {rule_index} from sheet '{sheet_name}' for {user_google_email}."


# ============================================================================
# READING CELL FORMATTING METADATA
# ============================================================================


@server.tool()
@handle_http_errors("read_sheet_formatting", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_sheet_formatting(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    ranges: Optional[List[str]] = None,
    include_values: bool = True,
    include_formulas: bool = False,
    summary_only: bool = False,
) -> str:
    """
    Reads comprehensive formatting information for specified ranges in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        ranges (Optional[List[str]]): List of A1 notation ranges (e.g., ["Sheet1!A1:D10"]).
                                      If None, reads first 100 cells of first sheet.
        include_values (bool): Whether to include cell values. Defaults to True.
        include_formulas (bool): Whether to include formulas. Defaults to False.
        summary_only (bool): If True, returns a summary instead of detailed formatting. Defaults to False.

    Returns:
        str: Detailed or summarized formatting information for the specified ranges.
    """
    logger.info(
        f"[read_sheet_formatting] Reading formatting for spreadsheet {spreadsheet_id}"
    )

    # If no ranges specified, use a default range
    if not ranges:
        # Get basic spreadsheet info first to get sheet names
        basic_info = await asyncio.to_thread(
            service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
        )
        first_sheet = basic_info.get("sheets", [{}])[0]
        sheet_name = first_sheet.get("properties", {}).get("title", "Sheet1")
        ranges = [f"{sheet_name}!A1:Z100"]

    # Build the fields parameter for selective retrieval
    fields = ["sheets.properties", "sheets.data.startRow", "sheets.data.startColumn"]

    if include_values:
        fields.append("sheets.data.rowData.values.userEnteredValue")
        fields.append("sheets.data.rowData.values.formattedValue")

    if include_formulas:
        fields.append("sheets.data.rowData.values.userEnteredValue.formulaValue")

    # Always include formatting fields
    fields.extend(
        [
            "sheets.data.rowData.values.effectiveFormat",
            "sheets.data.rowData.values.userEnteredFormat",
        ]
    )

    # Make the API call with includeGridData
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=ranges,
            includeGridData=True,
            fields=",".join(fields),
        )
        .execute
    )

    # Process the response
    if summary_only:
        return _summarize_formatting_data(spreadsheet, ranges, user_google_email)
    else:
        return _format_detailed_formatting(
            spreadsheet, ranges, include_values, user_google_email
        )


@server.tool()
@handle_http_errors(
    "get_spreadsheet_metadata", is_read_only=True, service_type="sheets"
)
@require_google_service("sheets", "sheets_read")
async def get_spreadsheet_metadata(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    include_grid_data: bool = False,
    ranges: Optional[List[str]] = None,
) -> str:
    """
    Gets comprehensive spreadsheet metadata including sheet properties, conditional formatting,
    protected ranges, and named ranges.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        include_grid_data (bool): Whether to include cell data. Defaults to False for performance.
        ranges (Optional[List[str]]): Specific ranges to include if include_grid_data is True.

    Returns:
        str: Comprehensive spreadsheet metadata including all structural information.
    """
    logger.info(
        f"[get_spreadsheet_metadata] Getting metadata for spreadsheet {spreadsheet_id}"
    )

    # Build request parameters
    params = {"spreadsheetId": spreadsheet_id}

    if include_grid_data:
        params["includeGridData"] = True
        if ranges:
            params["ranges"] = ranges

    # Get comprehensive spreadsheet data
    spreadsheet = await asyncio.to_thread(service.spreadsheets().get(**params).execute)

    # Extract metadata
    output = []

    # Basic spreadsheet properties
    props = spreadsheet.get("properties", {})
    output.append(f"**Spreadsheet: {props.get('title', 'Unknown')}**")
    output.append(f"ID: {spreadsheet_id}")
    output.append(f"Locale: {props.get('locale', 'Unknown')}")
    output.append(f"Time Zone: {props.get('timeZone', 'Unknown')}")

    # Default format info
    default_format = props.get("defaultFormat", {})
    if default_format:
        output.append("\n**Default Format:**")
        if "backgroundColor" in default_format:
            output.append(
                f"  Background: {_format_color(default_format['backgroundColor'])}"
            )
        if "textFormat" in default_format:
            tf = default_format["textFormat"]
            output.append(
                f"  Font: {tf.get('fontFamily', 'Default')} {tf.get('fontSize', 'Default')}pt"
            )

    # Sheets information
    sheets = spreadsheet.get("sheets", [])
    output.append(f"\n**Sheets ({len(sheets)}):**")

    for sheet in sheets:
        sheet_props = sheet.get("properties", {})
        sheet_name = sheet_props.get("title", "Unknown")
        sheet_id = sheet_props.get("sheetId", "Unknown")

        output.append(f"\n  **{sheet_name}** (ID: {sheet_id})")

        # Grid properties
        grid = sheet_props.get("gridProperties", {})
        output.append(
            f"    Size: {grid.get('rowCount', 0)}x{grid.get('columnCount', 0)}"
        )
        output.append(
            f"    Frozen: {grid.get('frozenRowCount', 0)} rows, {grid.get('frozenColumnCount', 0)} cols"
        )

        # Tab color if present
        if "tabColor" in sheet_props:
            output.append(f"    Tab Color: {_format_color(sheet_props['tabColor'])}")

        # Conditional formats
        cond_formats = sheet.get("conditionalFormats", [])
        if cond_formats:
            output.append(f"    Conditional Format Rules: {len(cond_formats)}")

        # Protected ranges
        protected = sheet.get("protectedRanges", [])
        if protected:
            output.append(f"    Protected Ranges: {len(protected)}")
            for pr in protected[:3]:  # Show first 3
                desc = pr.get("description", "No description")
                output.append(f"      - {desc}")

        # Basic filter
        if "basicFilter" in sheet:
            output.append("    Basic Filter: Active")

        # Filter views
        filter_views = sheet.get("filterViews", [])
        if filter_views:
            output.append(f"    Filter Views: {len(filter_views)}")

    # Named ranges
    named_ranges = spreadsheet.get("namedRanges", [])
    if named_ranges:
        output.append(f"\n**Named Ranges ({len(named_ranges)}):**")
        for nr in named_ranges[:5]:  # Show first 5
            output.append(f"  - {nr.get('name', 'Unknown')}: {nr.get('range', {})}")

    # Developer metadata if present
    dev_metadata = spreadsheet.get("developerMetadata", [])
    if dev_metadata:
        output.append(f"\n**Developer Metadata: {len(dev_metadata)} items**")

    return "\n".join(output) + f"\n\nMetadata retrieved for {user_google_email}."


@server.tool()
@handle_http_errors("read_cell_properties", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_cell_properties(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range: str,
    properties: Optional[List[str]] = None,
) -> str:
    """
    Reads specific cell properties for a single range with performance optimization.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): Single A1 notation range to analyze (e.g., "Sheet1!A1:D10"). Required.
        properties (Optional[List[str]]): Specific properties to retrieve.
                                          Options: "background", "text_format", "number_format",
                                                  "borders", "alignment", "padding", "wrap"
                                          If None, retrieves all properties.

    Returns:
        str: Filtered cell property data in a readable format.
    """
    logger.info(f"[read_cell_properties] Reading properties for range {range}")

    # Map property names to API field paths
    property_map = {
        "background": "sheets.data.rowData.values.effectiveFormat.backgroundColor",
        "text_format": "sheets.data.rowData.values.effectiveFormat.textFormat",
        "number_format": "sheets.data.rowData.values.effectiveFormat.numberFormat",
        "borders": "sheets.data.rowData.values.effectiveFormat.borders",
        "alignment": "sheets.data.rowData.values.effectiveFormat.horizontalAlignment,sheets.data.rowData.values.effectiveFormat.verticalAlignment",
        "padding": "sheets.data.rowData.values.effectiveFormat.padding",
        "wrap": "sheets.data.rowData.values.effectiveFormat.wrapStrategy",
    }

    # Build fields list
    fields = ["sheets.properties", "sheets.data.startRow", "sheets.data.startColumn"]
    fields.append(
        "sheets.data.rowData.values.formattedValue"
    )  # Always include formatted value

    if properties:
        for prop in properties:
            if prop in property_map:
                fields.extend(property_map[prop].split(","))
    else:
        # Get all formatting properties
        fields.append("sheets.data.rowData.values.effectiveFormat")
        fields.append("sheets.data.rowData.values.userEnteredFormat")

    # Make the API call
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=[range],
            includeGridData=True,
            fields=",".join(fields),
        )
        .execute
    )

    # Process and format the response
    output = [f"**Cell Properties for Range: {range}**\n"]

    sheets = spreadsheet.get("sheets", [])
    if not sheets:
        return f"No data found for range {range}"

    sheet = sheets[0]
    data = sheet.get("data", [])
    if not data:
        return f"No cell data found for range {range}"

    grid_data = data[0]
    row_data = grid_data.get("rowData", [])

    # Analyze formatting patterns
    format_summary = _analyze_formatting_patterns(row_data, properties)
    output.append(format_summary)

    # Show sample cell details (first few cells)
    output.append("\n**Sample Cell Details:**")
    sample_count = 0
    for row_idx, row in enumerate(row_data[:5]):  # First 5 rows
        values = row.get("values", [])
        for col_idx, cell in enumerate(values[:5]):  # First 5 columns
            if sample_count >= 10:  # Limit to 10 cells total
                break

            cell_ref = f"{_column_index_to_letter(grid_data.get('startColumn', 0) + col_idx)}{grid_data.get('startRow', 0) + row_idx + 1}"
            cell_desc = _describe_cell_properties(cell, properties)
            if cell_desc:
                output.append(f"\n  Cell {cell_ref}:")
                output.append(cell_desc)
                sample_count += 1

    return "\n".join(output) + f"\n\nProperties retrieved for {user_google_email}."


# ============================================================================
# HELPER FUNCTIONS FOR FORMATTING METADATA
# ============================================================================


def _format_color(color_dict: Dict) -> str:
    """Convert color dictionary to readable format."""
    if not color_dict:
        return "Default"

    red = int(color_dict.get("red", 0) * 255)
    green = int(color_dict.get("green", 0) * 255)
    blue = int(color_dict.get("blue", 0) * 255)

    # Convert to hex
    hex_color = f"#{red:02x}{green:02x}{blue:02x}"

    # Add alpha if present
    if "alpha" in color_dict:
        alpha = int(color_dict["alpha"] * 255)
        hex_color += f" (alpha: {alpha})"

    return hex_color


def _describe_cell_properties(
    cell: Dict, properties: Optional[List[str]] = None
) -> str:
    """Describe the properties of a single cell."""
    descriptions = []

    formatted_value = cell.get("formattedValue", "")
    if formatted_value:
        descriptions.append(f"    Value: {formatted_value}")

    # Get effective format (what the user sees)
    eff_format = cell.get("effectiveFormat", {})

    # Background color
    if not properties or "background" in properties:
        bg_color = eff_format.get("backgroundColor")
        if bg_color:
            descriptions.append(f"    Background: {_format_color(bg_color)}")

    # Text format
    if not properties or "text_format" in properties:
        text_format = eff_format.get("textFormat", {})
        if text_format:
            text_desc = []
            if "foregroundColor" in text_format:
                text_desc.append(
                    f"color: {_format_color(text_format['foregroundColor'])}"
                )
            if text_format.get("bold"):
                text_desc.append("bold")
            if text_format.get("italic"):
                text_desc.append("italic")
            if text_format.get("underline"):
                text_desc.append("underline")
            if text_format.get("strikethrough"):
                text_desc.append("strikethrough")
            if "fontSize" in text_format:
                text_desc.append(f"{text_format['fontSize']}pt")
            if "fontFamily" in text_format:
                text_desc.append(text_format["fontFamily"])

            if text_desc:
                descriptions.append(f"    Text: {', '.join(text_desc)}")

    # Number format
    if not properties or "number_format" in properties:
        num_format = eff_format.get("numberFormat", {})
        if num_format:
            format_type = num_format.get("type", "UNKNOWN")
            pattern = num_format.get("pattern", "")
            if pattern:
                descriptions.append(f"    Number Format: {format_type} ({pattern})")
            else:
                descriptions.append(f"    Number Format: {format_type}")

    # Alignment
    if not properties or "alignment" in properties:
        h_align = eff_format.get("horizontalAlignment")
        v_align = eff_format.get("verticalAlignment")
        if h_align or v_align:
            align_desc = []
            if h_align:
                align_desc.append(f"H: {h_align}")
            if v_align:
                align_desc.append(f"V: {v_align}")
            descriptions.append(f"    Alignment: {', '.join(align_desc)}")

    # Wrap strategy
    if not properties or "wrap" in properties:
        wrap = eff_format.get("wrapStrategy")
        if wrap:
            descriptions.append(f"    Text Wrap: {wrap}")

    # Borders
    if not properties or "borders" in properties:
        borders = eff_format.get("borders", {})
        if borders:
            border_desc = []
            for side in ["top", "bottom", "left", "right"]:
                if side in borders:
                    border = borders[side]
                    style = border.get("style", "NONE")
                    if style != "NONE":
                        border_desc.append(f"{side}: {style}")
            if border_desc:
                descriptions.append(f"    Borders: {', '.join(border_desc)}")

    return "\n".join(descriptions) if descriptions else ""


def _analyze_formatting_patterns(
    row_data: List[Dict], properties: Optional[List[str]] = None
) -> str:
    """Analyze and summarize formatting patterns in the data."""
    patterns = {
        "backgrounds": {},
        "text_colors": {},
        "fonts": {},
        "number_formats": {},
        "alignments": {},
        "borders": 0,
    }

    total_cells = 0
    formatted_cells = 0

    for row in row_data:
        for cell in row.get("values", []):
            total_cells += 1
            eff_format = cell.get("effectiveFormat", {})

            if eff_format:
                formatted_cells += 1

                # Track background colors
                bg = eff_format.get("backgroundColor")
                if bg:
                    color_key = _format_color(bg)
                    patterns["backgrounds"][color_key] = (
                        patterns["backgrounds"].get(color_key, 0) + 1
                    )

                # Track text colors
                text_format = eff_format.get("textFormat", {})
                if text_format:
                    fg = text_format.get("foregroundColor")
                    if fg:
                        color_key = _format_color(fg)
                        patterns["text_colors"][color_key] = (
                            patterns["text_colors"].get(color_key, 0) + 1
                        )

                    # Track fonts
                    font = text_format.get("fontFamily")
                    if font:
                        patterns["fonts"][font] = patterns["fonts"].get(font, 0) + 1

                # Track number formats
                num_fmt = eff_format.get("numberFormat", {})
                if num_fmt:
                    fmt_type = num_fmt.get("type", "UNKNOWN")
                    patterns["number_formats"][fmt_type] = (
                        patterns["number_formats"].get(fmt_type, 0) + 1
                    )

                # Track alignment
                h_align = eff_format.get("horizontalAlignment")
                if h_align:
                    patterns["alignments"][h_align] = (
                        patterns["alignments"].get(h_align, 0) + 1
                    )

                # Count cells with borders
                if eff_format.get("borders"):
                    patterns["borders"] += 1

    # Build summary
    summary = ["**Formatting Analysis:**"]
    summary.append(f"  Total Cells: {total_cells}")
    summary.append(
        f"  Formatted Cells: {formatted_cells} ({formatted_cells * 100 // max(total_cells, 1)}%)"
    )

    if patterns["backgrounds"]:
        summary.append(f"\n  Background Colors: {len(patterns['backgrounds'])} unique")
        for color, count in list(patterns["backgrounds"].items())[:3]:
            summary.append(f"    - {color}: {count} cells")

    if patterns["text_colors"]:
        summary.append(f"\n  Text Colors: {len(patterns['text_colors'])} unique")
        for color, count in list(patterns["text_colors"].items())[:3]:
            summary.append(f"    - {color}: {count} cells")

    if patterns["fonts"]:
        summary.append("\n  Fonts Used:")
        for font, count in patterns["fonts"].items():
            summary.append(f"    - {font}: {count} cells")

    if patterns["number_formats"]:
        summary.append("\n  Number Formats:")
        for fmt, count in patterns["number_formats"].items():
            summary.append(f"    - {fmt}: {count} cells")

    if patterns["alignments"]:
        summary.append("\n  Alignments:")
        for align, count in patterns["alignments"].items():
            summary.append(f"    - {align}: {count} cells")

    if patterns["borders"] > 0:
        summary.append(f"\n  Cells with Borders: {patterns['borders']}")

    return "\n".join(summary)


def _summarize_formatting_data(
    spreadsheet: Dict, ranges: List[str], user_email: str
) -> str:
    """Create a summary of formatting data."""
    output = [f"**Formatting Summary for Ranges: {', '.join(ranges)}**\n"]

    for sheet in spreadsheet.get("sheets", []):
        sheet_name = sheet.get("properties", {}).get("title", "Unknown")
        data_ranges = sheet.get("data", [])

        if not data_ranges:
            continue

        output.append(f"\n**Sheet: {sheet_name}**")

        for data_range in data_ranges:
            row_data = data_range.get("rowData", [])
            if row_data:
                summary = _analyze_formatting_patterns(row_data)
                output.append(summary)

    return "\n".join(output) + f"\n\nFormatting summary retrieved for {user_email}."


def _format_detailed_formatting(
    spreadsheet: Dict, ranges: List[str], include_values: bool, user_email: str
) -> str:
    """Format detailed formatting information."""
    output = [f"**Detailed Formatting for Ranges: {', '.join(ranges)}**\n"]

    for sheet in spreadsheet.get("sheets", []):
        sheet_name = sheet.get("properties", {}).get("title", "Unknown")
        data_ranges = sheet.get("data", [])

        if not data_ranges:
            continue

        output.append(f"\n**Sheet: {sheet_name}**")

        for data_range in data_ranges:
            start_row = data_range.get("startRow", 0)
            start_col = data_range.get("startColumn", 0)
            row_data = data_range.get("rowData", [])

            # Show first 10 rows of detailed data
            for row_idx, row in enumerate(row_data[:10]):
                current_row = start_row + row_idx + 1
                values = row.get("values", [])

                if not values:
                    continue

                output.append(f"\n  Row {current_row}:")

                for col_idx, cell in enumerate(values[:10]):  # First 10 columns
                    current_col = _column_index_to_letter(start_col + col_idx)
                    cell_ref = f"{current_col}{current_row}"

                    # Get cell value if requested
                    if include_values:
                        formatted_value = cell.get("formattedValue", "")
                        if formatted_value:
                            output.append(f"    {cell_ref}: {formatted_value}")
                    else:
                        output.append(f"    {cell_ref}:")

                    # Describe formatting
                    cell_desc = _describe_cell_properties(cell)
                    if cell_desc:
                        output.append(cell_desc)

            if len(row_data) > 10:
                output.append(f"\n  ... and {len(row_data) - 10} more rows")

    return "\n".join(output) + f"\n\nDetailed formatting retrieved for {user_email}."


# ============================================================================
# NEW HELPER FUNCTIONS: DATA BOUNDARIES, TABLE STYLES, RESET FORMATTING
# ============================================================================


@server.tool()
@handle_http_errors("get_data_boundaries", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def get_data_boundaries(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: Optional[str] = None,
    include_empty_cells: bool = False,
) -> str:
    """
    Detect the boundaries of data in a sheet (the actual used range).
    
    This function helps identify where data actually exists in a sheet,
    useful for dynamic range operations and avoiding processing empty cells.
    
    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (Optional[str]): Name of the sheet to analyze. If None, uses the first sheet.
        include_empty_cells (bool): If True, includes cells with formatting but no values. Defaults to False.
    
    Returns:
        str: Information about data boundaries including range in A1 notation.
    
    Examples:
        # Get boundaries of the first sheet
        get_data_boundaries(..., spreadsheet_id="abc123")
        
        # Get boundaries of a specific sheet
        get_data_boundaries(..., spreadsheet_id="abc123", sheet_name="Sales Data")
        
        # Include formatted but empty cells
        get_data_boundaries(..., spreadsheet_id="abc123", include_empty_cells=True)
    
    Output Example:
        Data boundaries for sheet 'Sales':
        - Data Range: A1:F150
        - Rows with data: 150
        - Columns with data: 6 (A to F)
        - Total cells with data: 900
        - First data cell: A1
        - Last data cell: F150
    """
    logger.info(
        f"[get_data_boundaries] Getting data boundaries for {spreadsheet_id}, sheet: {sheet_name}"
    )
    
    # Get spreadsheet metadata
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            includeGridData=False
        ).execute
    )
    
    # Find the target sheet
    sheets = spreadsheet.get("sheets", [])
    if not sheets:
        return "No sheets found in the spreadsheet."
    
    target_sheet = None
    if sheet_name:
        for sheet in sheets:
            if sheet["properties"]["title"] == sheet_name:
                target_sheet = sheet
                break
        if not target_sheet:
            return f"Sheet '{sheet_name}' not found in spreadsheet."
    else:
        target_sheet = sheets[0]
    
    sheet_title = target_sheet["properties"]["title"]
    # sheet_id is not needed for this function
    
    # Get the sheet data with values
    range_to_check = f"{sheet_title}!A1:ZZ10000"  # Check a large range
    
    try:
        if include_empty_cells:
            # Get data including formatting
            result = await asyncio.to_thread(
                service.spreadsheets().get(
                    spreadsheetId=spreadsheet_id,
                    ranges=[range_to_check],
                    includeGridData=True,
                    fields="sheets.data.rowData.values(formattedValue,effectiveFormat)"
                ).execute
            )
            
            # Analyze grid data for boundaries
            sheet_data = result.get("sheets", [{}])[0].get("data", [{}])[0]
            row_data = sheet_data.get("rowData", [])
            
            max_row = 0
            max_col = 0
            min_row = float('inf')
            min_col = float('inf')
            cell_count = 0
            
            for row_idx, row in enumerate(row_data):
                values = row.get("values", [])
                for col_idx, cell in enumerate(values):
                    # Check if cell has value or formatting
                    has_value = "formattedValue" in cell
                    has_format = "effectiveFormat" in cell and cell["effectiveFormat"]
                    
                    if has_value or (include_empty_cells and has_format):
                        max_row = max(max_row, row_idx)
                        max_col = max(max_col, col_idx)
                        min_row = min(min_row, row_idx)
                        min_col = min(min_col, col_idx)
                        if has_value:
                            cell_count += 1
        else:
            # Get only cells with values
            result = await asyncio.to_thread(
                service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=range_to_check
                ).execute
            )
            
            values = result.get("values", [])
            if not values:
                return f"No data found in sheet '{sheet_title}'."
            
            max_row = len(values) - 1
            max_col = 0
            min_row = 0
            min_col = float('inf')
            cell_count = 0
            
            # Find the maximum column with data
            for row_idx, row in enumerate(values):
                if row:  # Row has data
                    # Find first non-empty cell in row
                    for col_idx, cell in enumerate(row):
                        if cell:  # Cell has value
                            max_col = max(max_col, col_idx)
                            min_col = min(min_col, col_idx)
                            cell_count += 1
            
            if min_col == float('inf'):
                min_col = 0
    
    except Exception as e:
        logger.error(f"Error getting data boundaries: {e}")
        return f"Error analyzing sheet data: {str(e)}"
    
    # Convert to A1 notation
    if cell_count == 0:
        return f"No data found in sheet '{sheet_title}'."
    
    start_col_letter = _column_index_to_letter(min_col)
    end_col_letter = _column_index_to_letter(max_col)
    start_row_num = min_row + 1
    end_row_num = max_row + 1
    
    data_range = f"{start_col_letter}{start_row_num}:{end_col_letter}{end_row_num}"
    full_range = f"{sheet_title}!{data_range}"
    
    rows_with_data = end_row_num - start_row_num + 1
    cols_with_data = max_col - min_col + 1
    
    output = f"""Data boundaries for sheet '{sheet_title}':
- Data Range: {data_range}
- Full Range: {full_range}
- Rows with data: {rows_with_data}
- Columns with data: {cols_with_data} ({start_col_letter} to {end_col_letter})
- Total cells with values: {cell_count}
- First data cell: {start_col_letter}{start_row_num}
- Last data cell: {end_col_letter}{end_row_num}"""
    
    if include_empty_cells:
        output += "\n- Note: Boundaries include formatted but empty cells"
    
    logger.info(f"Successfully analyzed data boundaries for sheet '{sheet_title}'")
    return output


@server.tool()
@handle_http_errors("apply_table_style", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def apply_table_style(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range: str,
    style: Literal["professional", "colorful", "minimal", "dark", "striped"] = "professional",
    has_header: bool = True,
    auto_resize_columns: bool = True,
) -> str:
    """
    Apply a predefined table style to a range for professional-looking data presentation.
    
    This function applies comprehensive formatting including headers, alternating rows,
    borders, and color schemes to make data tables more readable and visually appealing.
    
    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): The range to format as a table (e.g., "Sheet1!A1:F20"). Required.
        style (str): Predefined style to apply:
                    - "professional": Blue header, light gray alternating rows
                    - "colorful": Teal header, light blue alternating rows
                    - "minimal": Light gray header, subtle borders
                    - "dark": Dark theme with white text
                    - "striped": Bold stripes with no borders
                    Defaults to "professional".
        has_header (bool): If True, formats the first row as a header. Defaults to True.
        auto_resize_columns (bool): If True, auto-resizes columns to fit content. Defaults to True.
    
    Returns:
        str: Confirmation message of successful styling.
    
    Examples:
        # Apply professional style to data table
        apply_table_style(..., range="Sheet1!A1:F20", style="professional")
        
        # Apply dark theme without headers
        apply_table_style(..., range="A1:D100", style="dark", has_header=False)
        
        # Minimal style with auto-resize
        apply_table_style(..., range="Data!B2:G50", style="minimal", auto_resize_columns=True)
    
    Style Descriptions:
        - Professional: Clean business look with blue headers and subtle gray stripes
        - Colorful: Vibrant design with teal headers and light blue accents
        - Minimal: Understated elegance with light formatting
        - Dark: Modern dark theme ideal for presentations
        - Striped: Bold alternating rows without borders for easy scanning
    """
    logger.info(
        f"[apply_table_style] Applying {style} style to {range} in spreadsheet {spreadsheet_id}"
    )
    
    # Parse the range to get sheet ID and grid range
    sheet_name, cell_range = _parse_range(range)
    
    # Get spreadsheet metadata to find sheet ID
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )
    
    sheet_id = _get_sheet_id_by_name(spreadsheet, sheet_name)
    grid_range = _convert_a1_to_grid_range(cell_range, sheet_id)
    
    # Define style configurations
    styles = {
        "professional": {
            "header_bg": "#1E40AF",  # Blue
            "header_fg": "#FFFFFF",  # White
            "odd_row_bg": "#FFFFFF",  # White
            "even_row_bg": "#F3F4F6",  # Light gray
            "border_color": "#D1D5DB",  # Gray
            "border_style": "SOLID",
        },
        "colorful": {
            "header_bg": "#0D9488",  # Teal
            "header_fg": "#FFFFFF",  # White
            "odd_row_bg": "#FFFFFF",  # White
            "even_row_bg": "#DBEAFE",  # Light blue
            "border_color": "#60A5FA",  # Blue
            "border_style": "SOLID",
        },
        "minimal": {
            "header_bg": "#F9FAFB",  # Very light gray
            "header_fg": "#111827",  # Dark gray
            "odd_row_bg": "#FFFFFF",  # White
            "even_row_bg": "#FAFAFA",  # Off white
            "border_color": "#E5E7EB",  # Light gray
            "border_style": "SOLID",
        },
        "dark": {
            "header_bg": "#1F2937",  # Dark gray
            "header_fg": "#F9FAFB",  # Light gray
            "odd_row_bg": "#374151",  # Medium gray
            "even_row_bg": "#4B5563",  # Lighter gray
            "text_color": "#F9FAFB",  # Light gray text
            "border_color": "#6B7280",  # Gray
            "border_style": "SOLID",
        },
        "striped": {
            "header_bg": "#7C3AED",  # Purple
            "header_fg": "#FFFFFF",  # White
            "odd_row_bg": "#FFFFFF",  # White
            "even_row_bg": "#E9D5FF",  # Light purple
            "border_style": None,  # No borders
        }
    }
    
    selected_style = styles.get(style, styles["professional"])
    
    requests = []
    
    # 1. Apply header formatting if has_header
    if has_header:
        header_range = {
            "sheetId": sheet_id,
            "startRowIndex": grid_range.get("startRowIndex", 0),
            "endRowIndex": grid_range.get("startRowIndex", 0) + 1,
            "startColumnIndex": grid_range.get("startColumnIndex", 0),
            "endColumnIndex": grid_range.get("endColumnIndex", 1),
        }
        
        header_format = {
            "backgroundColor": _hex_to_rgb_dict(selected_style["header_bg"]),
            "textFormat": {
                "foregroundColor": _hex_to_rgb_dict(selected_style["header_fg"]),
                "bold": True,
                "fontSize": 11,
            },
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        }
        
        # Add borders to header if style has borders
        if selected_style.get("border_style"):
            border = {
                "style": selected_style["border_style"],
                "color": _hex_to_rgb_dict(selected_style.get("border_color", "#000000")),
            }
            header_format["borders"] = {
                "top": border,
                "bottom": {"style": "SOLID_THICK", "color": border["color"]},
                "left": border,
                "right": border,
            }
        
        requests.append({
            "repeatCell": {
                "range": header_range,
                "cell": {"userEnteredFormat": header_format},
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)",
            }
        })
    
    # 2. Apply alternating row colors for data rows
    data_start_row = (grid_range.get("startRowIndex", 0) + 1) if has_header else grid_range.get("startRowIndex", 0)
    data_end_row = grid_range.get("endRowIndex", 1)
    
    # Apply odd row formatting
    for row_idx in range(data_start_row, data_end_row):
        is_even = (row_idx - data_start_row) % 2 == 1
        bg_color = selected_style["even_row_bg"] if is_even else selected_style["odd_row_bg"]
        
        row_range = {
            "sheetId": sheet_id,
            "startRowIndex": row_idx,
            "endRowIndex": row_idx + 1,
            "startColumnIndex": grid_range.get("startColumnIndex", 0),
            "endColumnIndex": grid_range.get("endColumnIndex", 1),
        }
        
        row_format = {
            "backgroundColor": _hex_to_rgb_dict(bg_color),
        }
        
        # Add text color for dark theme
        if style == "dark":
            row_format["textFormat"] = {
                "foregroundColor": _hex_to_rgb_dict(selected_style.get("text_color", "#FFFFFF"))
            }
        
        # Add borders if style has them
        if selected_style.get("border_style"):
            border = {
                "style": selected_style["border_style"],
                "color": _hex_to_rgb_dict(selected_style.get("border_color", "#000000")),
            }
            row_format["borders"] = {
                "top": border if row_idx == data_start_row else None,
                "bottom": border,
                "left": border,
                "right": border,
            }
        
        fields = "userEnteredFormat(backgroundColor"
        if "textFormat" in row_format:
            fields += ",textFormat"
        if "borders" in row_format:
            fields += ",borders"
        fields += ")"
        
        requests.append({
            "repeatCell": {
                "range": row_range,
                "cell": {"userEnteredFormat": row_format},
                "fields": fields,
            }
        })
    
    # 3. Auto-resize columns if requested
    if auto_resize_columns:
        requests.append({
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": grid_range.get("startColumnIndex", 0),
                    "endIndex": grid_range.get("endColumnIndex", 1),
                }
            }
        })
    
    # Execute the batch update
    body = {"requests": requests}
    await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
        .execute
    )
    
    logger.info(f"Successfully applied {style} table style to range {range}")
    
    return (
        f"Successfully applied '{style}' table style to range '{range}' in spreadsheet {spreadsheet_id}.\n"
        f"Style features:\n"
        f"- Header row: {'Yes (bold with colored background)' if has_header else 'No'}\n"
        f"- Alternating rows: Yes\n"
        f"- Borders: {'Yes' if selected_style.get('border_style') else 'No'}\n"
        f"- Auto-resize columns: {'Yes' if auto_resize_columns else 'No'}\n"
        f"Table styling complete for {user_google_email}."
    )


@server.tool()
@handle_http_errors("reset_to_default_formatting", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def reset_to_default_formatting(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range: str,
    preserve_values: bool = True,
    clear_conditional_formatting: bool = True,
) -> str:
    """
    Reset cell formatting to Google Sheets defaults while preserving data.
    
    This function removes all custom formatting from the specified range,
    returning cells to the default appearance. Useful for cleaning up
    over-formatted sheets or starting fresh with formatting.
    
    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): The range to reset (e.g., "Sheet1!A1:D10"). Required.
        preserve_values (bool): If True, keeps cell values intact. If False, clears values too. Defaults to True.
        clear_conditional_formatting (bool): If True, also removes conditional formatting rules. Defaults to True.
    
    Returns:
        str: Confirmation message of successful reset.
    
    Examples:
        # Reset formatting only, keep data
        reset_to_default_formatting(..., range="Sheet1!A1:Z100")
        
        # Clear everything (formatting and values)
        reset_to_default_formatting(..., range="A1:D50", preserve_values=False)
        
        # Reset format but keep conditional formatting rules
        reset_to_default_formatting(..., range="Data!B2:F20", clear_conditional_formatting=False)
    
    What Gets Reset:
        - Background colors → White
        - Text colors → Black
        - Font styles → Default (Arial, 10pt, no bold/italic/underline)
        - Borders → None
        - Number formats → Automatic
        - Alignment → Default (left for text, right for numbers)
        - Text wrapping → Overflow
        - Conditional formatting → Removed (if clear_conditional_formatting=True)
    
    What Gets Preserved:
        - Cell values and formulas (if preserve_values=True)
        - Cell comments
        - Data validation rules
        - Protected ranges
    """
    logger.info(
        f"[reset_to_default_formatting] Resetting formatting for range {range} in spreadsheet {spreadsheet_id}"
    )
    
    # Parse the range to get sheet ID and grid range
    sheet_name, cell_range = _parse_range(range)
    
    # Get spreadsheet metadata to find sheet ID
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )
    
    sheet_id = _get_sheet_id_by_name(spreadsheet, sheet_name)
    grid_range = _convert_a1_to_grid_range(cell_range, sheet_id)
    
    requests = []
    
    # 1. Clear all formatting by applying empty format
    # This effectively resets to defaults
    default_format = {
        # Google Sheets defaults
        "backgroundColor": {"red": 1, "green": 1, "blue": 1},  # White
        "textFormat": {
            "foregroundColor": {"red": 0, "green": 0, "blue": 0},  # Black
            "fontFamily": "Arial",
            "fontSize": 10,
            "bold": False,
            "italic": False,
            "underline": False,
            "strikethrough": False,
        },
        "horizontalAlignment": "LEFT",  # Default alignment
        "verticalAlignment": "BOTTOM",
        "wrapStrategy": "OVERFLOW_CELL",
        "numberFormat": {
            "type": "NUMBER",
            "pattern": ""  # Automatic formatting
        },
        # Clear borders by setting them to none
        "borders": {
            "top": {"style": "NONE"},
            "bottom": {"style": "NONE"},
            "left": {"style": "NONE"},
            "right": {"style": "NONE"},
        }
    }
    
    requests.append({
        "repeatCell": {
            "range": grid_range,
            "cell": {"userEnteredFormat": default_format},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy,numberFormat,borders)",
        }
    })
    
    # 2. Clear conditional formatting if requested
    if clear_conditional_formatting:
        # First, get existing conditional formatting rules for the sheet
        sheet_props = None
        for sheet in spreadsheet.get("sheets", []):
            if sheet["properties"]["sheetId"] == sheet_id:
                sheet_props = sheet
                break
        
        if sheet_props and "conditionalFormats" in sheet_props:
            # Get rule indices that affect our range
            rules_to_delete = []
            for idx, rule in enumerate(sheet_props["conditionalFormats"]):
                for rule_range in rule.get("ranges", []):
                    # Check if this rule affects our range (simplified check)
                    if rule_range.get("sheetId") == sheet_id:
                        # For simplicity, delete all rules on the sheet when range overlaps
                        # A more sophisticated approach would check actual range overlap
                        rules_to_delete.append(idx)
                        break
            
            # Delete rules in reverse order to maintain indices
            for idx in sorted(rules_to_delete, reverse=True):
                requests.append({
                    "deleteConditionalFormatRule": {
                        "sheetId": sheet_id,
                        "index": idx
                    }
                })
    
    # 3. Clear values if requested
    if not preserve_values:
        # We'll clear values after the formatting reset
        pass  # Will handle separately with clear API
    
    # Execute the batch update for formatting
    if requests:
        body = {"requests": requests}
        await asyncio.to_thread(
            service.spreadsheets()
            .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
            .execute
        )
    
    # Clear values if requested (separate API call)
    if not preserve_values:
        full_range = f"{sheet_name}!{cell_range}"
        await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=full_range)
            .execute
        )
    
    logger.info(f"Successfully reset formatting for range {range}")
    
    result_message = (
        f"Successfully reset formatting for range '{range}' in spreadsheet {spreadsheet_id}.\n"
        f"Actions taken:\n"
        f"- Cell formatting: Reset to defaults\n"
        f"- Cell values: {'Preserved' if preserve_values else 'Cleared'}\n"
        f"- Conditional formatting: {'Removed' if clear_conditional_formatting else 'Preserved'}\n"
        f"- Borders: Removed\n"
        f"- Colors: Reset to default (white background, black text)\n"
        f"- Font: Reset to Arial 10pt\n"
        f"Formatting reset complete for {user_google_email}."
    )
    
    return result_message


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================


def _parse_range(range_str: str) -> tuple:
    """Parse a range string like 'Sheet1!A1:D10' into sheet name and cell range."""
    if "!" in range_str:
        parts = range_str.split("!", 1)
        return parts[0], parts[1]
    else:
        # Assume first sheet if no sheet specified
        return "Sheet1", range_str


def _get_sheet_id_by_name(spreadsheet: Dict, sheet_name: str) -> int:
    """Get sheet ID from sheet name."""
    sheets = spreadsheet.get("sheets", [])
    for sheet in sheets:
        if sheet["properties"]["title"] == sheet_name:
            return sheet["properties"]["sheetId"]
    raise ValueError(f"Sheet '{sheet_name}' not found")


def _convert_a1_to_grid_range(a1_notation: str, sheet_id: int) -> Dict:
    """Convert A1 notation to GridRange object."""

    grid_range = {"sheetId": sheet_id}

    # Parse A1 notation (simplified - handles basic cases)
    match = re.match(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", a1_notation)
    if match:
        start_col = _column_letter_to_index(match.group(1))
        start_row = int(match.group(2)) - 1

        grid_range["startRowIndex"] = start_row
        grid_range["startColumnIndex"] = start_col

        if match.group(3):  # Has end range
            end_col = _column_letter_to_index(match.group(3))
            end_row = int(match.group(4))
            grid_range["endRowIndex"] = end_row
            grid_range["endColumnIndex"] = end_col + 1
        else:
            grid_range["endRowIndex"] = start_row + 1
            grid_range["endColumnIndex"] = start_col + 1

    return grid_range


def _grid_range_to_a1(grid_range: Dict, sheet_name: str) -> str:
    """Convert GridRange to A1 notation."""
    start_col = _column_index_to_letter(grid_range.get("startColumnIndex", 0))
    start_row = grid_range.get("startRowIndex", 0) + 1

    if "endColumnIndex" in grid_range and "endRowIndex" in grid_range:
        end_col = _column_index_to_letter(grid_range["endColumnIndex"] - 1)
        end_row = grid_range["endRowIndex"]
        return f"{sheet_name}!{start_col}{start_row}:{end_col}{end_row}"
    else:
        return f"{sheet_name}!{start_col}{start_row}"


def _column_letter_to_index(letter: str) -> int:
    """Convert column letter to 0-based index."""
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def _column_index_to_letter(index: int) -> str:
    """Convert 0-based column index to letter."""
    letter = ""
    index += 1
    while index > 0:
        index -= 1
        letter = chr(index % 26 + ord("A")) + letter
        index //= 26
    return letter


def _hex_to_rgb_dict(hex_color: str) -> Dict:
    """Convert hex color to RGB dictionary for Sheets API."""
    hex_color = hex_color.lstrip("#")
    return {
        "red": int(hex_color[0:2], 16) / 255.0,
        "green": int(hex_color[2:4], 16) / 255.0,
        "blue": int(hex_color[4:6], 16) / 255.0,
    }


def _get_update_fields_from_format(cell_format: Dict) -> str:
    """Generate the fields parameter for update requests."""
    fields = []

    if "backgroundColor" in cell_format:
        fields.append("userEnteredFormat.backgroundColor")
    if "textFormat" in cell_format:
        fields.append("userEnteredFormat.textFormat")
    if "horizontalAlignment" in cell_format:
        fields.append("userEnteredFormat.horizontalAlignment")
    if "verticalAlignment" in cell_format:
        fields.append("userEnteredFormat.verticalAlignment")
    if "wrapStrategy" in cell_format:
        fields.append("userEnteredFormat.wrapStrategy")
    if "numberFormat" in cell_format:
        fields.append("userEnteredFormat.numberFormat")
    if "borders" in cell_format:
        fields.append("userEnteredFormat.borders")

    return ",".join(fields)
