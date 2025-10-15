"""Domain models for COSMIC Function Point measurement."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class DataMovementType(str, Enum):
    """Enumeration of COSMIC data movement types."""

    ENTRY = "E"
    EXIT = "X"
    READ = "R"
    WRITE = "W"

    @classmethod
    def from_string(cls, value: str) -> "DataMovementType":
        """Create an enum member from a case-insensitive string."""
        normalized = value.strip().upper()
        try:
            return cls(normalized)
        except ValueError as exc:  # pragma: no cover - defensive branch
            valid = ", ".join(member.value for member in cls)
            raise ValueError(f"Unsupported data movement type '{value}'. Expected one of: {valid}.") from exc


@dataclass
class DataMovement:
    """Represents a single COSMIC data movement."""

    movement_type: DataMovementType
    description: str
    object_of_interest: Optional[str] = None
    trigger: Optional[str] = None
    code_reference: Optional[str] = None
    additional_notes: Optional[str] = None

    @classmethod
    def from_dict(cls, payload: dict) -> "DataMovement":
        """Construct a :class:`DataMovement` from a mapping."""
        if "type" not in payload:
            raise ValueError("Data movement definition must include a 'type' field.")
        if "description" not in payload:
            raise ValueError("Data movement definition must include a 'description' field.")

        movement_type = DataMovementType.from_string(str(payload["type"]))
        description = str(payload["description"]).strip()

        return cls(
            movement_type=movement_type,
            description=description,
            object_of_interest=payload.get("object_of_interest") or payload.get("ooi"),
            trigger=payload.get("trigger"),
            code_reference=payload.get("code_reference"),
            additional_notes=payload.get("notes") or payload.get("additional_notes"),
        )


@dataclass
class FunctionalProcess:
    """Represents a COSMIC functional process."""

    name: str
    description: Optional[str] = None
    trigger: Optional[str] = None
    object_of_interest: Optional[str] = None
    data_movements: List[DataMovement] = field(default_factory=list)

    @classmethod
    def from_dict(cls, payload: dict) -> "FunctionalProcess":
        """Construct a :class:`FunctionalProcess` from a mapping."""
        if "name" not in payload:
            raise ValueError("Functional process definition must include a 'name' field.")

        movements_payload = payload.get("data_movements", []) or []
        movements = [DataMovement.from_dict(item) for item in movements_payload]

        return cls(
            name=str(payload["name"]).strip(),
            description=(payload.get("description") or payload.get("purpose")),
            trigger=payload.get("trigger"),
            object_of_interest=payload.get("object_of_interest") or payload.get("ooi"),
            data_movements=movements,
        )

    def count_movements(self, movement_type: DataMovementType) -> int:
        """Return the number of data movements of the specified type."""
        return sum(1 for movement in self.data_movements if movement.movement_type is movement_type)

    @property
    def total_cfp(self) -> int:
        """Total CFP for this functional process."""
        return len(self.data_movements)


@dataclass
class SystemMeasurement:
    """Represents a full COSMIC measurement for a system boundary."""

    name: str
    boundary: Optional[str]
    description: Optional[str]
    persistence_resources: List[str] = field(default_factory=list)
    external_actors: List[str] = field(default_factory=list)
    objects_of_interest: List[str] = field(default_factory=list)
    functional_processes: List[FunctionalProcess] = field(default_factory=list)

    @classmethod
    def from_dict(cls, payload: dict) -> "SystemMeasurement":
        """Construct a :class:`SystemMeasurement` from a mapping."""
        system_info = payload.get("system") or {}
        objects_payload = payload.get("objects") or payload.get("objects_of_interest") or []
        processes_payload = payload.get("functional_processes") or []

        functional_processes = [FunctionalProcess.from_dict(item) for item in processes_payload]

        return cls(
            name=str(system_info.get("name", "Unnamed System")),
            boundary=system_info.get("boundary"),
            description=system_info.get("description"),
            persistence_resources=list(system_info.get("persistence_resources", [])),
            external_actors=list(system_info.get("external_actors", [])),
            objects_of_interest=[str(obj) for obj in objects_payload],
            functional_processes=functional_processes,
        )

    @property
    def total_cfp(self) -> int:
        """Total CFP for the entire system measurement."""
        return sum(process.total_cfp for process in self.functional_processes)
