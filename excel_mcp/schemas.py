"""Pydantic models and type definitions for the Excel MCP server."""

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field


# Type alias for cell values
CellValue = Union[str, int, float, bool, None]
CellValueOrList = Union[CellValue, list[CellValue], list[list[CellValue]]]


# =============================================================================
# Style Models
# =============================================================================


class StyleOptions(BaseModel):
    """Style options for cells."""

    bold: Annotated[bool | None, Field(description="Set bold text")] = None
    italic: Annotated[bool | None, Field(description="Set italic text")] = None
    font_size: Annotated[int | None, Field(description="Font size in points")] = None
    font_color: Annotated[
        str | None, Field(description="Font color as hex, e.g., 'FF0000' for red")
    ] = None
    bg_color: Annotated[
        str | None,
        Field(description="Background color as hex, e.g., '00FF00' for green"),
    ] = None
    h_align: Annotated[
        Literal["left", "center", "right"] | None,
        Field(description="Horizontal alignment"),
    ] = None
    v_align: Annotated[
        Literal["top", "center", "bottom"] | None,
        Field(description="Vertical alignment"),
    ] = None


# =============================================================================
# Sheet Operations
# =============================================================================


class RenameSheetOperation(BaseModel):
    """Operation to rename a sheet."""

    old_name: Annotated[str, Field(description="Current name of the sheet")]
    new_name: Annotated[str, Field(description="New name for the sheet")]


class CreateSheetOperation(BaseModel):
    """Operation to create new worksheets."""

    sheet_names: Annotated[
        list[str], Field(description="Names for the new sheets to create")
    ]


class DeleteSheetOperation(BaseModel):
    """Operation to delete worksheets."""

    sheet_names: Annotated[list[str], Field(description="Names of sheets to delete")]


class DisplaySheetOperation(BaseModel):
    """Operation to display sheet contents."""

    sheet_name: Annotated[
        str | None,
        Field(description="Sheet name to read (optional, uses active sheet)"),
    ] = None
    max_rows: Annotated[int, Field(description="Maximum rows to display")] = 50
    width: Annotated[
        int,
        Field(description="Maximum width for the table output (cells wrap if needed)"),
    ] = 300


# =============================================================================
# Row/Column Operations
# =============================================================================


class InsertRowsOperation(BaseModel):
    """Operation to insert rows."""

    position: Annotated[int, Field(description="Row number (1-based) where to insert")]
    count: Annotated[int, Field(description="Number of rows to insert")] = 1
    sheet_name: Annotated[
        str | None, Field(description="Sheet name (optional, uses active sheet)")
    ] = None


class InsertColumnsOperation(BaseModel):
    """Operation to insert columns."""

    position: Annotated[
        int, Field(description="Column number (1-based) where to insert")
    ]
    count: Annotated[int, Field(description="Number of columns to insert")] = 1
    sheet_name: Annotated[
        str | None, Field(description="Sheet name (optional, uses active sheet)")
    ] = None


class DeleteRowsOperation(BaseModel):
    """Operation to delete rows."""

    position: Annotated[
        int, Field(description="Row number (1-based) to start deletion")
    ]
    count: Annotated[int, Field(description="Number of rows to delete")] = 1
    sheet_name: Annotated[
        str | None, Field(description="Sheet name (optional, uses active sheet)")
    ] = None


class DeleteColumnsOperation(BaseModel):
    """Operation to delete columns."""

    position: Annotated[
        int, Field(description="Column number (1-based) to start deletion")
    ]
    count: Annotated[int, Field(description="Number of columns to delete")] = 1
    sheet_name: Annotated[
        str | None, Field(description="Sheet name (optional, uses active sheet)")
    ] = None


# =============================================================================
# Range Data Operations
# =============================================================================


class GetRangeDataOperation(BaseModel):
    """Operation to get cell data for a range."""

    range_a1: Annotated[
        str | list[str],
        Field(
            description="Range in A1 notation (e.g., 'A1', 'B2:C3', 'Sheet1!A1:B2'), or comma-separated ranges"
        ),
    ]
    return_style_info: Annotated[
        bool,
        Field(
            description="If True, returns JSON with style info. Default is formatted table."
        ),
    ] = False
    width: Annotated[
        int,
        Field(description="Maximum width for the table output (cells wrap if needed)"),
    ] = 300
    max_rows: Annotated[
        int,
        Field(description="Maximum rows to display (default 100, prevents huge outputs)"),
    ] = 100


class SetRangeDataOperation(BaseModel):
    """Operation to set value(s) for cells in a range."""

    range_a1: Annotated[
        str,
        Field(
            description="Range in A1 notation (e.g., 'A1', 'Sheet1!A1:B2'). Supports comma-separated ranges."
        ),
    ]
    value: Annotated[
        CellValueOrList,
        Field(
            description="Value(s) to set. Can be: single value, 1D list (fills row/column), 2D list (fills grid), or formula starting with '='"
        ),
    ]


class SetRangeOperation(BaseModel):
    """Operation to set data in a range."""

    range_a1: Annotated[
        str, Field(description="Range in A1 notation (e.g., 'A1', 'A1:B2')")
    ]
    value: Annotated[
        CellValueOrList,
        Field(description="Value(s) to set - single value, 1D list, or 2D list"),
    ]


class AutoFillOperation(BaseModel):
    """Operation to auto fill data from source range to target range."""

    source_range: Annotated[
        str, Field(description="Source range in A1 notation (e.g., 'A1:A2', 'B1:C1')")
    ]
    target_range: Annotated[
        str,
        Field(
            description="Target range in A1 notation to fill (e.g., 'A1:A10', 'B1:F1')"
        ),
    ]


# =============================================================================
# Style Operations
# =============================================================================


class SetRangeStyleOperation(BaseModel):
    """Operation to set style properties for cells in a range."""

    range_a1: Annotated[
        str, Field(description="Range in A1 notation (supports comma-separated ranges)")
    ]
    bold: Annotated[bool | None, Field(description="Set bold text")] = None
    italic: Annotated[bool | None, Field(description="Set italic text")] = None
    font_size: Annotated[int | None, Field(description="Font size in points")] = None
    font_color: Annotated[
        str | None, Field(description="Font color as hex (e.g., 'FF0000')")
    ] = None
    bg_color: Annotated[
        str | None, Field(description="Background color as hex (e.g., '00FF00')")
    ] = None
    h_align: Annotated[
        Literal["left", "center", "right"] | None,
        Field(description="Horizontal alignment"),
    ] = None
    v_align: Annotated[
        Literal["top", "center", "bottom"] | None,
        Field(description="Vertical alignment"),
    ] = None


class BatchStyleOperation(BaseModel):
    """Batch operation to apply style to a range."""

    range_a1: Annotated[
        str, Field(description="Range in A1 notation (e.g., 'A1:B2', 'Sheet1!A1:C3')")
    ]
    style: Annotated[StyleOptions, Field(description="Style options to apply")]


class MergeCellsOperation(BaseModel):
    """Operation to merge cells in a range."""

    range_a1: Annotated[
        str, Field(description="Range to merge in A1 notation (e.g., 'A1:B2')")
    ]


class UnmergeCellsOperation(BaseModel):
    """Operation to unmerge cells in a range."""

    range_a1: Annotated[
        str, Field(description="Range to unmerge in A1 notation (e.g., 'A1:B2')")
    ]


# =============================================================================
# Search Operations
# =============================================================================


class SearchCellsOperation(BaseModel):
    """Operation to search for cells containing a keyword."""

    keyword: Annotated[str, Field(description="Search keyword to find in cells")]
    find_by: Annotated[
        Literal["value", "formula"],
        Field(description="Search in cell values or formulas"),
    ] = "value"
    sheet_name: Annotated[
        str | None, Field(description="Sheet to search (optional, uses active sheet)")
    ] = None
    max_results: Annotated[
        int, Field(description="Maximum number of results to return")
    ] = 50


# =============================================================================
# Conditional Formatting
# =============================================================================


class ConditionalFormatRule(BaseModel):
    """Conditional formatting rule parameters.

    Common fields:
    - rule_id: Optional[str] - Rule identifier (auto-generated if not provided)
    - range: Optional[str] - Target range (A1 notation)
    - ranges: Optional[list[str]] - Multiple target ranges

    Rule type:
    - rule_type: str - 'highlightCell' | 'colorScale' | 'expression'

    HighlightCell fields:
    - operator: str - 'greaterThan', 'lessThan', 'equal', 'between', etc.
    - value: number | str | list - Value(s) for comparison
    - formula: str - Formula for expression-based rules
    - style: dict - Style to apply (fgColor, bgColor, bold, italic)

    ColorScale fields:
    - points: list[dict] - Color scale points with 'color' and optional 'value_type'/'value'

    Style dict:
    - fgColor: str (font color, e.g., "#FF0000")
    - bgColor: str (background color, e.g., "#00FF00")
    - bold: bool
    - italic: bool
    """

    model_config = ConfigDict(extra="allow")


class AddConditionalFormattingOperation(BaseModel):
    """Operation to add conditional formatting rules."""

    sheet_name: Annotated[str, Field(description="Target sheet name")]
    rules: Annotated[
        list[ConditionalFormatRule],
        Field(description="List of conditional formatting rules to add"),
    ]


class DeleteConditionalFormattingOperation(BaseModel):
    """Operation to delete conditional formatting rules."""

    sheet_name: Annotated[str, Field(description="Target sheet name")]
    rule_ids: Annotated[list[str], Field(description="List of rule IDs to delete")]


class GetConditionalFormattingOperation(BaseModel):
    """Operation to get conditional formatting rules."""

    sheet_name: Annotated[str, Field(description="Target sheet name")]


# =============================================================================
# Data Validation
# =============================================================================


class DataValidationRule(BaseModel):
    """Data validation rule parameters.

    Basic fields:
    - rule_id: Optional[str] - Rule identifier (auto-generated if not provided)
    - range_a1: str - Target range (A1 notation), comma-separated for multiple ranges
    - ranges: list[str] - Alternative: list of target ranges

    Validation type:
    - validation_type: str - Type of validation:
        'list': Dropdown list selection
        'whole': Integer validation
        'decimal': Decimal number validation
        'date': Date validation
        'textLength': Text length validation
        'custom': Custom formula validation
        'checkbox': Checkbox (uses list with TRUE,FALSE)

    Operator (for whole, decimal, date, textLength):
    - operator: str - 'between', 'notBetween', 'equal', 'notEqual',
                      'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual'

    Value parameters:
    - value1/formula1: First value or formula
    - value2/formula2: Second value (for 'between'/'notBetween')
    - source: str - For list validation: comma-separated values or cell range

    Custom formula:
    - custom_formula: str - Custom validation formula

    Options:
    - ignore_blank: bool (default True)
    - show_error_message: bool (default True)
    - show_input_message: bool (default True)
    - input_title/title: str - Input prompt title
    - input_message/prompt: str - Input prompt message
    """

    model_config = ConfigDict(extra="allow")


class AddDataValidationOperation(BaseModel):
    """Operation to add data validation rules."""

    sheet_name: Annotated[str, Field(description="Target sheet name")]
    rules: Annotated[
        list[DataValidationRule],
        Field(description="List of data validation rules to add"),
    ]


class DeleteDataValidationOperation(BaseModel):
    """Operation to delete data validation rules."""

    sheet_name: Annotated[str, Field(description="Target sheet name")]
    rule_ids: Annotated[list[str], Field(description="List of rule IDs to delete")]


class GetDataValidationOperation(BaseModel):
    """Operation to get data validation rules."""

    sheet_name: Annotated[str, Field(description="Target sheet name")]


# =============================================================================
# Misc Operations
# =============================================================================


class RequestNewToolOperation(BaseModel):
    """Operation to request a new tool."""

    tool_name: Annotated[str, Field(description="Suggested name for the new tool")]
    description: Annotated[str, Field(description="What the tool should do")]
    parameters: Annotated[
        str, Field(description="What parameters the tool should accept")
    ]
    reason: Annotated[
        str, Field(description="Why existing tools cannot accomplish this task")
    ]
