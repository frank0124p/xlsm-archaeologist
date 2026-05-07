"""Rich progress bar wrapper for consistent CLI progress display."""

from __future__ import annotations

from types import TracebackType
from typing import Self

from rich.progress import BarColumn, Progress, SpinnerColumn, TaskID, TextColumn, TimeElapsedColumn


class ProgressBar:
    """Context manager wrapping rich.progress.Progress.

    Usage::

        with ProgressBar(quiet=False) as bar:
            task = bar.add_task("Extracting sheets", total=100)
            for i in range(100):
                bar.advance(task)
    """

    def __init__(self, quiet: bool = False) -> None:
        """Initialize progress bar.

        Args:
            quiet: When True, suppresses all output (CI mode).
        """
        self._quiet = quiet
        self._progress: Progress | None = None

    def __enter__(self) -> Self:
        if not self._quiet:
            self._progress = Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeElapsedColumn(),
            )
            self._progress.__enter__()
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: TracebackType | None,
    ) -> None:
        if self._progress is not None:
            self._progress.__exit__(exc_type, exc_val, exc_tb)

    def add_task(self, description: str, total: float = 100) -> TaskID:
        """Add a new task to the progress bar.

        Args:
            description: Task label shown next to the bar.
            total: Total work units (default 100).

        Returns:
            TaskID for use with advance().
        """
        if self._progress is not None:
            return self._progress.add_task(description, total=total)
        return TaskID(0)

    def advance(self, task_id: TaskID, advance: float = 1) -> None:
        """Advance a task by the given amount.

        Args:
            task_id: ID returned by add_task().
            advance: Work units to advance (default 1).
        """
        if self._progress is not None:
            self._progress.advance(task_id, advance)
