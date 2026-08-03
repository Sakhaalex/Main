from enum import Enum

class LXLockState(Enum):
    FREE = 0
    BUSY = 1
    SKIP = 2
    ERROR = 3

class LX_BaseFrame:
    """Universal Lock Interface for all Frame modules."""
    def __init__(self, alias: str):
        self.alias = alias
        self._lock_status = LXLockState.FREE
        self._internal_locks = {}

    def get_lock_status(self) -> LXLockState:
        return self._lock_status

    def set_lock_status(self, status: LXLockState):
        self._lock_status = status

    def set_internal_lock(self, resource_id: str, is_locked: bool):
        self._internal_locks[resource_id] = is_locked

    def is_any_internal_locked(self) -> bool:
        return any(self._internal_locks.values())

    def execute(self, row_data: dict) -> bool:
        """To be overridden by individual Frame modules."""
        raise NotImplementedError("Subclasses must implement execute()")
