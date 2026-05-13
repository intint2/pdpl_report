"""powd_pl_report — pick data root via numbered menu, then batch. No file dialogs."""

from __future__ import annotations

import sys
from pathlib import Path

from batch_plot_pptx import run_powd_batch_report, run_powd_batch_report_editable


def _configure_stdio_utf8() -> None:
    if sys.platform != "win32":
        return
    for stream in (sys.stdin, sys.stdout):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:
            pass


def project_root() -> Path:
    """Directory that contains main.py (browse starts here on request)."""
    return Path(__file__).resolve().parent


def powd_data_root() -> Path:
    """Default powd data root: project_root/datas/powd."""
    return project_root() / "datas" / "powd"


def prompt_choice(prompt: str, lo: int, hi: int) -> int:
    while True:
        raw = input(prompt).strip()
        if not raw.isdigit():
            print(f"Enter a number from {lo} to {hi}.")
            continue
        n = int(raw)
        if lo <= n <= hi:
            return n
        print(f"Enter a number from {lo} to {hi}.")


def browse_data_root(start: Path) -> Path:
    """
    [1] parent (..)
    [2 .. k+1] subfolders (sorted)
    [k+2] use this folder as data root (returns to main menu after setting selection in caller)
    """
    current = start.resolve()
    while True:
        subdirs = sorted(
            (p for p in current.iterdir() if p.is_dir()),
            key=lambda p: p.name.casefold(),
        )
        can_go_up = current != current.parent
        k = len(subdirs)
        confirm_idx = k + 2

        print()
        print(f"Current folder: {current}")
        print("  [1] Parent folder (..)" if can_go_up else "  [1] Parent folder (not available)")
        for j, p in enumerate(subdirs, start=2):
            print(f"  [{j}] {p.name}")
        print(f"  [{confirm_idx}] Use this folder as data root")

        choice = prompt_choice(f"Choice (1-{confirm_idx}): ", 1, confirm_idx)

        if choice == confirm_idx:
            return current
        if choice == 1:
            if not can_go_up:
                print("Already at the top of this path.")
                continue
            current = current.parent
            continue
        current = subdirs[choice - 2]


def condition_subdirs(powd_root: Path) -> list[Path]:
    return sorted(
        (p for p in powd_root.iterdir() if p.is_dir()),
        key=lambda p: p.name.casefold(),
    )


def run_batch(powd_root: Path) -> None:
    """Batch per condition: overlay spectra with matplotlib, export one PPTX."""
    folders = condition_subdirs(powd_root)
    total_csv = 0
    for cond in folders:
        total_csv += sum(1 for _ in cond.glob("*.csv"))
    run_powd_batch_report(
        powd_root,
        project_root=project_root(),
        output_dir=project_root() / "output",
    )
    print(f"Data root: {powd_root}")
    print(f"Condition folders: {len(folders)} | CSV files total: {total_csv}")


def run_batch_editable_pptx(powd_root: Path) -> None:
    """Batch per condition: native PowerPoint XY line charts (editable in Office)."""
    folders = condition_subdirs(powd_root)
    total_csv = 0
    for cond in folders:
        total_csv += sum(1 for _ in cond.glob("*.csv"))
    run_powd_batch_report_editable(
        powd_root,
        project_root=project_root(),
        output_dir=project_root() / "output",
    )
    print(f"Data root: {powd_root}")
    print(f"Condition folders: {len(folders)} | CSV files total: {total_csv}")


def main() -> None:
    _configure_stdio_utf8()
    selected: Path | None = None

    while True:
        print()
        active = selected if selected is not None else powd_data_root()
        print(f"Data root for run if not changed: {active}")
        print("  [1] Select folder path")
        print("  [2] Run batch — matplotlib PNG slides → PPTX")
        print("  [3] Run batch plot — native Office charts (editable data / formatting in PowerPoint)")
        choice = prompt_choice("Choice (1-3): ", 1, 3)
        if choice == 1:
            selected = browse_data_root(project_root())
            continue

        root = selected if selected is not None else powd_data_root()
        if not root.is_dir():
            print(f"Not a folder: {root}", file=sys.stderr)
            sys.exit(1)
        if choice == 2:
            run_batch(root)
        else:
            run_batch_editable_pptx(root)
        break


if __name__ == "__main__":
    main()
