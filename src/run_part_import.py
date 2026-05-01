from __future__ import annotations

import argparse
import contextlib
import io
import traceback
from datetime import datetime
from pathlib import Path

import import_excel_to_sqlite


class Tee(io.TextIOBase):
    def __init__(self, *streams: io.TextIOBase) -> None:
        self.streams = streams

    def write(self, text: str) -> int:
        for stream in self.streams:
            try:
                stream.write(text)
                stream.flush()
            except ValueError:
                pass
        return len(text)

    def flush(self) -> None:
        for stream in self.streams:
            try:
                stream.flush()
            except ValueError:
                pass


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="파트 설정 기반 SQLite 적재 실행기")
    parser.add_argument("--config", required=True, type=Path, help="파트 설정 JSON")
    parser.add_argument("--year", default=str(datetime.now().year), help="연도 폴더명")
    parser.add_argument("--source-dir", type=Path)
    parser.add_argument("--target-dir", type=Path)
    parser.add_argument("--single-db-name")
    parser.add_argument("--replace-db", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args(argv)


def build_log_path(config_path: Path, year: str) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_dir = config_path.resolve().parents[2] / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    stem = config_path.stem
    return log_dir / f"{stem}_{year}_{timestamp}.log"


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    log_path = build_log_path(args.config, args.year)

    import_args = [
        "--config",
        str(args.config),
        "--year",
        args.year,
    ]
    if args.source_dir:
        import_args.extend(["--source-dir", str(args.source_dir)])
    if args.target_dir:
        import_args.extend(["--target-dir", str(args.target_dir)])
    if args.single_db_name:
        import_args.extend(["--single-db-name", args.single_db_name])
    if args.replace_db:
        import_args.append("--replace-db")
    if args.dry_run:
        import_args.append("--dry-run")

    with log_path.open("w", encoding="utf-8") as log_file:
        tee = Tee(log_file)
        with contextlib.redirect_stdout(tee), contextlib.redirect_stderr(tee):
            print(f"Log file: {log_path}")
            print(f"Started at: {datetime.now().isoformat(timespec='seconds')}")
            print(f"Config file: {args.config}")
            print(f"Year: {args.year}")
            try:
                exit_code = import_excel_to_sqlite.main(import_args)
                print(f"Finished with exit code: {exit_code}")
                return exit_code
            except Exception:
                print("Import failed with an exception:")
                traceback.print_exc()
                return 1


if __name__ == "__main__":
    raise SystemExit(main())
