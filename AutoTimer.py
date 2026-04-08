from __future__ import annotations

import ctypes
import json
import signal
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import psutil


# Display name -> executable name aliases (lowercase)
APP_RULES = {
	"Word": {"winword.exe"},
	"MATLAB": {"matlab.exe"},
	"VSCode": {"code.exe", "code - insiders.exe"},
	"Zotero": {"zotero.exe"},
	"PowerPoint": {"powerpnt.exe"},
	"Excel": {"excel.exe"},
	"TaskManager": {"taskmgr.exe"},
	"Explorer": {"explorer.exe"},
}


POLL_INTERVAL_SEC = 1.0
IDLE_TIMEOUT_SEC = 60.0
SUMMARY_EVERY_SEC = 600.0


@dataclass
class LASTINPUTINFO(ctypes.Structure):
	_fields_ = [
		("cbSize", ctypes.c_uint),
		("dwTime", ctypes.c_uint),
	]


def get_idle_seconds() -> float:
	user32 = ctypes.windll.user32
	kernel32 = ctypes.windll.kernel32
	info = LASTINPUTINFO()
	info.cbSize = ctypes.sizeof(LASTINPUTINFO)
	if not user32.GetLastInputInfo(ctypes.byref(info)):
		return 0.0
	tick_count = kernel32.GetTickCount()
	elapsed_ms = tick_count - info.dwTime
	return max(0.0, elapsed_ms / 1000.0)


def get_foreground_pid() -> int | None:
	user32 = ctypes.windll.user32
	hwnd = user32.GetForegroundWindow()
	if hwnd == 0:
		return None

	pid = ctypes.c_ulong()
	user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
	if pid.value == 0:
		return None
	return int(pid.value)


def get_foreground_exe_name() -> str | None:
	pid = get_foreground_pid()
	if pid is None:
		return None

	try:
		return psutil.Process(pid).name().lower()
	except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
		return None


def resolve_app_name(exe_name: str | None) -> str | None:
	if exe_name is None:
		return None
	for app_name, aliases in APP_RULES.items():
		if exe_name in aliases:
			return app_name
	return None


def format_seconds(seconds: float) -> str:
	total = int(round(seconds))
	hours = total // 3600
	minutes = (total % 3600) // 60
	sec = total % 60
	return f"{hours:02d}:{minutes:02d}:{sec:02d}"


class AutoTimer:
	def __init__(self) -> None:
		self.start_time = datetime.now()
		self.app_seconds = {name: 0.0 for name in APP_RULES}
		self.total_seconds = 0.0
		self.running = True

		self._last_loop_time = time.perf_counter()
		self._last_status_key = None
		self._last_summary_time = self._last_loop_time

	def print_config(self) -> None:
		print("\n=== AutoTimer Configuration ===")
		print(f"Idle timeout: {int(IDLE_TIMEOUT_SEC)} seconds")
		print("Tracked applications:")
		for app_name, aliases in APP_RULES.items():
			alias_text = ", ".join(sorted(aliases))
			print(f"  - {app_name}: {alias_text}")
		print("===============================\n")

	def _print_status_if_changed(self, app_name: str | None, idle_seconds: float) -> None:
		is_idle = idle_seconds >= IDLE_TIMEOUT_SEC
		status_key = (app_name, is_idle)
		if status_key == self._last_status_key:
			return

		self._last_status_key = status_key
		now_text = datetime.now().strftime("%H:%M:%S")

		if app_name is None:
			print(f"[{now_text}] Paused | foreground app not in whitelist")
			return

		if is_idle:
			print(
				f"[{now_text}] Paused | {app_name} in foreground but idle {idle_seconds:.0f}s >= {int(IDLE_TIMEOUT_SEC)}s"
			)
			return

		print(f"[{now_text}] Running | app={app_name}, idle={idle_seconds:.0f}s")

	def _print_periodic_summary(self, now_perf: float) -> None:
		if now_perf - self._last_summary_time < SUMMARY_EVERY_SEC:
			return

		self._last_summary_time = now_perf
		print("\n--- Timer Summary ---")
		print(f"Total tracked time: {format_seconds(self.total_seconds)}")
		for app_name, sec in self.app_seconds.items():
			print(f"  {app_name:<12} {format_seconds(sec)}")
		print("---------------------\n")

	def update(self) -> None:
		now_perf = time.perf_counter()
		delta = max(0.0, now_perf - self._last_loop_time)
		self._last_loop_time = now_perf

		idle_seconds = get_idle_seconds()
		exe_name = get_foreground_exe_name()
		app_name = resolve_app_name(exe_name)

		self._print_status_if_changed(app_name, idle_seconds)

		if app_name is not None and idle_seconds < IDLE_TIMEOUT_SEC:
			self.app_seconds[app_name] += delta
			self.total_seconds += delta

		self._print_periodic_summary(now_perf)

	def stop(self) -> None:
		self.running = False

	def dump_report(self) -> Path:
		end_time = datetime.now()
		payload = {
			"start_time": self.start_time.isoformat(timespec="seconds"),
			"end_time": end_time.isoformat(timespec="seconds"),
			"elapsed_wall_clock_seconds": (end_time - self.start_time).total_seconds(),
			"tracked_total_seconds": self.total_seconds,
			"tracked_total_hms": format_seconds(self.total_seconds),
			"per_app_seconds": self.app_seconds,
			"per_app_hms": {k: format_seconds(v) for k, v in self.app_seconds.items()},
			"idle_timeout_seconds": IDLE_TIMEOUT_SEC,
			"poll_interval_seconds": POLL_INTERVAL_SEC,
			"app_rules": {k: sorted(v) for k, v in APP_RULES.items()},
		}

		out_name = f"autotimer_report_{end_time.strftime('%Y%m%d-%H%M%S')}.json"
		out_path = Path(out_name)
		out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
		return out_path

	def print_final_summary(self, report_path: Path) -> None:
		print("\n=== Final Timer Summary ===")
		print(f"Total tracked time: {format_seconds(self.total_seconds)}")
		for app_name, sec in self.app_seconds.items():
			print(f"  {app_name:<12} {format_seconds(sec)}")
		print(f"Report saved to: {report_path}")
		print("===========================")


def main() -> None:
	timer = AutoTimer()
	timer.print_config()

	def _handle_signal(sig_num, _frame) -> None:
		print(f"\nReceived stop signal: {sig_num}")
		timer.stop()

	signal.signal(signal.SIGINT, _handle_signal)
	signal.signal(signal.SIGTERM, _handle_signal)
	if hasattr(signal, "SIGBREAK"):
		signal.signal(signal.SIGBREAK, _handle_signal)

	print("AutoTimer started. Press Ctrl+C to stop.\n")

	try:
		while timer.running:
			timer.update()
			try:
				time.sleep(POLL_INTERVAL_SEC)
			except KeyboardInterrupt:
				timer.stop()
	except KeyboardInterrupt:
		timer.stop()
	finally:
		report_path = timer.dump_report()
		timer.print_final_summary(report_path)


if __name__ == "__main__":
	main()
