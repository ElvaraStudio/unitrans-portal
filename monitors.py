"""Multi-monitor geometry.

Tk only knows about the primary display: ``winfo_screenwidth`` reports the
primary monitor's width and ``state("zoomed")`` always maximises there. On a
two-monitor desk that means the window can only ever open on the primary
screen. This module asks Windows directly where the monitors are, so the app
can start on a chosen display.

Coordinates are in the virtual-desktop space, where a second monitor to the
right of the primary has a positive x offset (e.g. x=2944 for a 1682-wide
primary plus a gap). That is exactly the space Tk's ``geometry("WxH+X+Y")``
expects, so the rectangles can be used directly.
"""

from __future__ import annotations

from dataclasses import dataclass

#: Sentinel for "whichever monitor the user last used".
MONITOR_PRIMARY = 0
MONITOR_SECONDARY = 1
MONITOR_DEFAULT = MONITOR_SECONDARY


@dataclass(frozen=True)
class Monitor:
    """A display's rectangle in virtual-desktop coordinates."""

    index: int
    left: int
    top: int
    right: int
    bottom: int
    primary: bool
    #: Work area, i.e. the rectangle minus the taskbar. Equal to the full
    #: rectangle when the taskbar is hidden or on another display.
    work_left: int = 0
    work_top: int = 0
    work_right: int = 0
    work_bottom: int = 0

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top

    @property
    def work_width(self) -> int:
        return max(1, self.work_right - self.work_left) if self.work_right else self.width

    @property
    def work_height(self) -> int:
        return max(1, self.work_bottom - self.work_top) if self.work_bottom else self.height

    def describe(self) -> str:
        """Human-readable summary, naming the usable area when a taskbar exists."""
        reserved = self.height - self.work_height
        if reserved > 0:
            return (
                f"{self.width}x{self.height} at ({self.left},{self.top}), "
                f"{self.work_width}x{self.work_height} usable"
            )
        return f"{self.width}x{self.height} at ({self.left},{self.top})"


def _windows_monitors() -> list[Monitor]:
    """Enumerate displays through the Win32 API, including work areas."""
    import ctypes
    import ctypes.wintypes as wintypes

    user32 = ctypes.windll.user32

    class MONITORINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", wintypes.RECT),
            ("rcWork", wintypes.RECT),
            ("dwFlags", wintypes.DWORD),
        ]

    found: list[Monitor] = []

    monitor_enum_proc = ctypes.WINFUNCTYPE(
        ctypes.c_int,
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.POINTER(wintypes.RECT),
        ctypes.c_void_p,
    )

    def callback(handle, _hdc, _rect_ptr, _data):
        info = MONITORINFO()
        info.cbSize = ctypes.sizeof(MONITORINFO)
        if user32.GetMonitorInfoW(handle, ctypes.byref(info)):
            screen = info.rcMonitor
            work = info.rcWork
            found.append(
                Monitor(
                    index=0,  # assigned after sorting
                    left=screen.left,
                    top=screen.top,
                    right=screen.right,
                    bottom=screen.bottom,
                    # MONITORINFOF_PRIMARY is 1; comparing against the virtual
                    # origin also works and is what the old code used.
                    primary=bool(info.dwFlags & 1),
                    work_left=work.left,
                    work_top=work.top,
                    work_right=work.right,
                    work_bottom=work.bottom,
                )
            )
        return 1

    user32.EnumDisplayMonitors(0, None, monitor_enum_proc(callback), 0)

    # Windows reports monitors in an unspecified order; sort left-to-right so
    # "second monitor" means the same thing on every run.
    found.sort(key=lambda monitor: (monitor.left, monitor.top))
    return [
        Monitor(
            index=index,
            left=m.left,
            top=m.top,
            right=m.right,
            bottom=m.bottom,
            primary=m.primary,
            work_left=m.work_left,
            work_top=m.work_top,
            work_right=m.work_right,
            work_bottom=m.work_bottom,
        )
        for index, m in enumerate(found)
    ]


def _single_monitor_fallback() -> list[Monitor]:
    """One monitor, sized from Tk when it is available."""
    width, height = 1920, 1080
    try:
        import tkinter as tk

        root = tk.Tk()
        root.withdraw()
        width = root.winfo_screenwidth()
        height = root.winfo_screenheight()
        root.destroy()
    except Exception:
        pass
    return [
        Monitor(
            index=0,
            left=0,
            top=0,
            right=width,
            bottom=height,
            primary=True,
            work_left=0,
            work_top=0,
            work_right=width,
            work_bottom=height,
        )
    ]


def list_monitors() -> list[Monitor]:
    """Every attached display, ordered left to right. Never empty."""
    try:
        monitors = _windows_monitors()
    except Exception:
        monitors = []
    if not monitors:
        monitors = _single_monitor_fallback()
    return monitors


def resolve_monitor(preference: int, monitors: list[Monitor] | None = None) -> Monitor:
    """Pick a monitor, falling back safely.

    ``preference`` is an index, but a user with one screen must not end up with
    a window placed off-screen, so an unavailable index degrades to the primary
    display.
    """
    available = monitors if monitors is not None else list_monitors()
    if not available:
        return _single_monitor_fallback()[0]

    if 0 <= preference < len(available):
        return available[preference]

    for monitor in available:
        if monitor.primary:
            return monitor
    return available[0]


def placement_for(
    preference: int,
    window_width: int,
    window_height: int,
    monitors: list[Monitor] | None = None,
    fill: bool = False,
) -> tuple[int, int, int, int, Monitor]:
    """Return ``(width, height, x, y, monitor)`` for opening the window.

    With ``fill=True`` the window occupies the whole work area of the chosen
    monitor, which is what "maximised" means here. Passing explicit dimensions
    is unreliable for that purpose because Tk's ``winfo_screenwidth`` reports
    only the *primary* display, so a window opened on a wider secondary screen
    would come up short.

    The result is clamped to the monitor's work area - the rectangle minus the
    taskbar - so the window never opens underneath the taskbar or spills past
    the screen edge.
    """
    monitor = resolve_monitor(preference, monitors)

    if fill:
        width = monitor.work_width
        height = monitor.work_height
    else:
        width = min(window_width, monitor.work_width)
        height = min(window_height, monitor.work_height)

    # Centre within the work area. A filled window is already exactly the work
    # area, so it lands at the work area's origin.
    x = monitor.work_left + max(0, (monitor.work_width - width) // 2)
    y = monitor.work_top + max(0, (monitor.work_height - height) // 2)
    return width, height, x, y, monitor


def describe_monitors() -> str:
    """One line per display, for the Settings window."""
    monitors = list_monitors()
    lines = []
    for monitor in monitors:
        tag = "primary" if monitor.primary else "secondary"
        lines.append(f"{monitor.index}: {monitor.describe()} ({tag})")
    return "\n".join(lines)


def toplevel_handle(window) -> int:
    """Win32 handle of a Tk toplevel's outer window frame.

    ``winfo_id()`` returns Tk's *inner* client window. That handle is not the
    frame: ``ShowWindow`` and ``SetWindowPos`` on it are ignored or affect only
    the client area, which is why maximising through it silently did nothing.
    The frame is its parent.
    """
    try:
        import ctypes

        inner = int(window.winfo_id())
        if not inner:
            return 0
        GA_PARENT = 1
        parent = ctypes.windll.user32.GetAncestor(inner, GA_PARENT)
        return int(parent) if parent else inner
    except Exception:
        return 0


def move_window(window, x: int, y: int, width: int, height: int) -> bool:
    """Move and size a Tk window on the virtual desktop, frame-accurately.

    Tk's own ``geometry()`` positions the client area and applies the window
    manager's frame adjustments on top, so a window cannot be placed precisely
    on a secondary monitor through Tk alone. ``SetWindowPos`` takes the outer
    frame rectangle, which is unambiguous.
    """
    try:
        import ctypes

        handle = toplevel_handle(window)
        if not handle:
            return False
        SWP_NOZORDER = 0x0004
        SWP_NOACTIVATE = 0x0010  # do not steal focus from whatever has it
        ctypes.windll.user32.SetWindowPos(
            handle, 0, int(x), int(y), int(width), int(height),
            SWP_NOZORDER | SWP_NOACTIVATE,
        )
        return True
    except Exception:
        return False


def maximise_window(window) -> bool:
    """Maximise a Tk window natively, on whichever monitor it currently sits on.

    Prefer :func:`apply_startup_placement`, which moves *and* maximises in one
    step. This exists for callers that only need the maximise.

    ``Tk.state("zoomed")`` delegates to the window manager asynchronously, so a
    request made while the window is being moved can be resolved against the
    window's previous bounds - which puts it back on the monitor it came from.
    ``ShowWindow(SW_MAXIMIZE)`` acts immediately on the window's current
    monitor and leaves the title bar's own maximise button working normally.
    """
    try:
        import ctypes

        handle = toplevel_handle(window)
        if not handle:
            return False
        SW_MAXIMIZE = 3
        ctypes.windll.user32.ShowWindow(handle, SW_MAXIMIZE)
        return True
    except Exception:
        return False


def apply_startup_placement(
    window,
    preference: int,
    maximise: bool = True,
    monitors: list[Monitor] | None = None,
    window_width: int = 1400,
    window_height: int = 900,
) -> Monitor | None:
    """Open ``window`` on the chosen display, optionally maximised.

    ``maximise=False`` gives a sensible windowed size instead, which is what
    you want when the app is meant to come up on a screen the user is *not*
    working on: it should be there when they look, without covering anything
    they are using.

    Returns the monitor the window was placed on, or None if placement failed.
    """
    available = monitors if monitors is not None else list_monitors()
    if not available:
        if maximise:
            maximise_window(window)
        return None

    target_width = max(200, min(window_width, 1400))
    target_height = max(200, min(window_height, 900))

    monitor = resolve_monitor(preference, available)
    # Centre horizontally, and sit a third of the way down rather than dead
    # centre - that reads as intentional rather than as a misplaced dialog.
    x = monitor.left + max(0, (monitor.width - target_width) // 2)
    y = monitor.top + max(0, (monitor.height - target_height) // 3)

    moved = move_window(window, x, y, target_width, target_height)
    if not moved:
        # Fall back to Tk's own geometry, which is less precise but works.
        try:
            window.geometry(f"{target_width}x{target_height}+{x}+{y}")
            moved = True
        except Exception:
            moved = False

    if maximise and moved:
        try:
            window.update_idletasks()
        except Exception:
            pass
        maximise_window(window)

    return monitor if moved else None


def monitor_of_point(x: int, y: int, monitors: list[Monitor] | None = None) -> Monitor | None:
    """Which display contains the point, or None if it is on none of them."""
    for monitor in monitors if monitors is not None else list_monitors():
        if monitor.left <= x < monitor.right and monitor.top <= y < monitor.bottom:
            return monitor
    return None
