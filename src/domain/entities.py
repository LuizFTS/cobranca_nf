from dataclasses import dataclass
from datetime import datetime

@dataclass
class BranchReport:
    branch: str
    period_initial: datetime
    period_final: datetime
    excel_path: str
    quantity: int
    total_value: float
    table: str