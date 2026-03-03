import os
import sys
from typing import Iterable, List, Optional


def _clear_screen() -> None:
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
    except Exception:
        print("\n" * 50)


def select_icgs(
    icgs: Iterable[str],
    title: str = "Pilih ICG yang akan digenerate",
) -> Optional[List[str]]:
    """Interactive ICG picker.

    Controls (Windows terminal):
    - Up/Down arrows: move cursor
    - Space: toggle selection
    - A: select all
    - N: select none
    - Enter: confirm
    - Esc / Q: cancel (returns None)

    Returns:
        - list of selected ICGs (can be empty if user selects none)
        - None if user cancels

    Notes:
        Uses `msvcrt` on Windows. On non-Windows / non-interactive terminals,
        falls back to a simple text prompt.
    """

    icg_list = [str(x).strip() for x in icgs if str(x).strip()]
    # keep order, drop duplicates
    seen = set()
    icg_list = [x for x in icg_list if not (x in seen or seen.add(x))]

    if not icg_list:
        return []

    if os.name != 'nt':
        return _select_icgs_fallback(icg_list, title)

    try:
        import msvcrt  # type: ignore
    except Exception:
        return _select_icgs_fallback(icg_list, title)

    selected = [False] * len(icg_list)
    cursor = 0
    top = 0
    page_size = 20

    def render() -> None:
        nonlocal top
        _clear_screen()
        print(title)
        print("(↑/↓) move | (Space) toggle | (A) all | (N) none | (Enter) OK | (Q/Esc) cancel")
        print("-")

        # keep cursor visible
        if cursor < top:
            top = cursor
        elif cursor >= top + page_size:
            top = cursor - page_size + 1

        end = min(len(icg_list), top + page_size)
        for i in range(top, end):
            mark = "[x]" if selected[i] else "[ ]"
            pointer = ">" if i == cursor else " "
            print(f"{pointer} {mark} {icg_list[i]}")

        if end < len(icg_list):
            print(f"... ({len(icg_list) - end} item lagi)")

        chosen = sum(1 for v in selected if v)
        print("-")
        print(f"Selected: {chosen}/{len(icg_list)}")

    render()

    while True:
        ch = msvcrt.getwch()

        # arrows are two-key sequences on Windows
        if ch in ('\x00', '\xe0'):
            ch2 = msvcrt.getwch()
            if ch2 == 'H':  # up
                cursor = max(0, cursor - 1)
                render()
                continue
            if ch2 == 'P':  # down
                cursor = min(len(icg_list) - 1, cursor + 1)
                render()
                continue
            continue

        if ch in ('q', 'Q', '\x1b'):
            return None

        if ch == ' ':  # toggle
            selected[cursor] = not selected[cursor]
            render()
            continue

        if ch in ('a', 'A'):
            selected = [True] * len(icg_list)
            render()
            continue

        if ch in ('n', 'N'):
            selected = [False] * len(icg_list)
            render()
            continue

        if ch in ('\r', '\n'):
            return [icg_list[i] for i, v in enumerate(selected) if v]


def _select_icgs_fallback(icg_list: List[str], title: str) -> Optional[List[str]]:
    print(title)
    print("Ketik 'all' untuk semua, atau masukkan nomor dipisah koma (contoh: 1,3,5). Kosong = all.")
    for i, icg in enumerate(icg_list, start=1):
        print(f"{i}. {icg}")

    raw = input("Pilih ICG: ").strip()
    if raw == "":
        return icg_list
    if raw.lower() in {"all", "*"}:
        return icg_list
    if raw.lower() in {"q", "quit", "exit"}:
        return None

    try:
        picks = []
        for part in raw.split(','):
            part = part.strip()
            if not part:
                continue
            idx = int(part)
            if idx < 1 or idx > len(icg_list):
                raise ValueError
            picks.append(icg_list[idx - 1])
        # keep order, drop duplicates
        seen = set()
        picks = [x for x in picks if not (x in seen or seen.add(x))]
        return picks
    except Exception:
        raise ValueError("Input tidak valid. Contoh yang valid: all atau 1,3,5")
