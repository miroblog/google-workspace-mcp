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
) -> str:
    """
    Simplified function for updating values in a Google Sheet range.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range (str): The A1 notation range to update (e.g., "Sheet1!A1:D10"). Required.
        values (Any): Values to write. Can be:
                     - Single value: "100" or 42
                     - 1D array (single row): ["A", "B", "C"]
                     - 2D array: [["A", "B"], ["C", "D"]]
        value_input_option (str): How to interpret values - "RAW" or "USER_ENTERED". Defaults to "USER_ENTERED".

    Returns:
        str: Confirmation message with update statistics.

    Examples:
        # Single cell
        update_sheet_values(..., range="A1", values="Hello")

        # Single row
        update_sheet_values(..., range="A1:C1", values=["One", "Two", "Three"])

        # Multiple rows
        update_sheet_values(..., range="A1:B2", values=[["A", "B"], ["C", "D"]])
    """
    logger.info(
        f"[update_sheet_values] Updating range {range} in spreadsheet {spreadsheet_id}"
    )

    # Validate and format values
    try:
        validated_values = _validate_2d_array(values)
    except ValueError as e:
        raise Exception(f"Invalid values format: {e}")

    # Prepare the request body
    body = {"values": validated_values}

    # Execute the update
    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range=range,
            valueInputOption=value_input_option,
            body=body,
        )
        .execute
    )

    # Extract update statistics
    updated_cells = result.get("updatedCells", 0)
    updated_rows = result.get("updatedRows", 0)
    updated_columns = result.get("updatedColumns", 0)
    updated_range = result.get("updatedRange", range)

    logger.info(f"Successfully updated {updated_cells} cells in range {updated_range}")

    return (
        f"Successfully updated range '{updated_range}' in spreadsheet {spreadsheet_id}.\n"
        f"Statistics: {updated_cells} cells, {updated_rows} rows, {updated_columns} columns updated."
    )


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
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    strikethrough: Optional[bool] = None,
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
        font_size (Optional[int]): Font size in points.
        font_family (Optional[str]): Font family name (e.g., "Arial", "Times New Roman").
        bold (Optional[bool]): Make text bold.
        italic (Optional[bool]): Make text italic.
        underline (Optional[bool]): Underline text.
        strikethrough (Optional[bool]): Strikethrough text.
        horizontal_alignment (Optional[str]): Horizontal text alignment.
        vertical_alignment (Optional[str]): Vertical text alignment.
        text_wrap (Optional[str]): Text wrapping strategy.
        number_format (Optional[str]): Predefined number format type.
        number_format_pattern (Optional[str]): Custom number format pattern (e.g., "#,##0.00").
        border_style (Optional[str]): Style for all borders.
        border_color (Optional[str]): Hex color for borders.

    Returns:
        str: Confirmation message of successful formatting.
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
    if bold is not None:
        text_format["bold"] = bold
    if italic is not None:
        text_format["italic"] = italic
    if underline is not None:
        text_format["underline"] = underline
    if strikethrough is not None:
        text_format["strikethrough"] = strikethrough

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
