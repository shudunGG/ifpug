"""CFP aggregation logic."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict

from .models import DataMovementType, FunctionalProcess, SystemMeasurement


@dataclass
class FunctionalProcessSummary:
    """Aggregated information for a functional process."""

    name: str
    entry_count: int
    exit_count: int
    read_count: int
    write_count: int
    total_cfp: int
    trigger: str | None
    object_of_interest: str | None
    description: str | None


class CosmicCalculator:
    """Compute COSMIC Function Point totals."""

    def __init__(self, measurement: SystemMeasurement) -> None:
        self.measurement = measurement

    def summarize_functional_process(self, process: FunctionalProcess) -> FunctionalProcessSummary:
        """Summarize the CFP metrics for a single functional process."""
        entry_count = process.count_movements(DataMovementType.ENTRY)
        exit_count = process.count_movements(DataMovementType.EXIT)
        read_count = process.count_movements(DataMovementType.READ)
        write_count = process.count_movements(DataMovementType.WRITE)

        return FunctionalProcessSummary(
            name=process.name,
            entry_count=entry_count,
            exit_count=exit_count,
            read_count=read_count,
            write_count=write_count,
            total_cfp=process.total_cfp,
            trigger=process.trigger,
            object_of_interest=process.object_of_interest,
            description=process.description,
        )

    def summarize(self) -> Dict[str, FunctionalProcessSummary]:
        """Return summaries for all functional processes."""
        return {
            process.name: self.summarize_functional_process(process)
            for process in self.measurement.functional_processes
        }

    def total_cfp(self) -> int:
        """Return the total CFP for the measurement."""
        return self.measurement.total_cfp
