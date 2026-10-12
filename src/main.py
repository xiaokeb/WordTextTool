import sys


if __name__ == "__main__":
    if len(sys.argv) == 4 and sys.argv[1] == "--office-job":
        from pathlib import Path
        from word_text_tool.office_backend import execute_job
        raise SystemExit(execute_job(Path(sys.argv[2]), Path(sys.argv[3])))
    from word_text_tool.app import run
    raise SystemExit(run())
