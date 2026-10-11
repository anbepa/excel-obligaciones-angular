from dataclasses import dataclass, field
from typing import Any

from openpyxl import Workbook


@dataclass
class ManagedWorkbook:
    """Workbook wrapper with Excel agent metadata.

    This class wraps an openpyxl Workbook and adds the metadata needed
    for the Excel agent to manage formulas, caching, and rules.
    """

    wb: Workbook
    filepath: str
    is_dirty: bool = False
    cached_values_wb: Workbook | None = None
    cf_rules_store: dict[str, dict[str, Any]] = field(default_factory=dict)
    dv_rules_store: dict[str, dict[str, Any]] = field(default_factory=dict)

    def close(self) -> None:
        """Close the workbook and release resources."""
        if self.cached_values_wb is not None:
            try:
                self.cached_values_wb.close()
            except Exception:
                pass
        try:
            self.wb.close()
        except Exception:
            pass


@dataclass
class UserState:
    """Holds the state for a single user session."""

    user_id: str
    workbook: ManagedWorkbook | None = None
