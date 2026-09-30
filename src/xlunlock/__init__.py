"""xlunlock: strip passwords and protection from Excel files, in place."""

__version__ = "0.1.0"

from .unlock import unlock, UnlockResult, UnlockError

__all__ = ["unlock", "UnlockResult", "UnlockError", "__version__"]
