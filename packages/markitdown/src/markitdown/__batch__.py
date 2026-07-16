import argparse
import os
import sys
import tempfile
import time
from pathlib import Path

from .__about__ import __version__
from ._markitdown import MarkItDown


class WordComServer:
    """Manages a single reusable Microsoft Word COM instance for .doc -> .docx conversion."""

    def __init__(self):
        self._word = None
        self._initialized = False

    def _ensure_com(self):
        """Initialize COM and launch Word if not already running."""
        if self._word is not None:
            return True
        try:
            import pythoncom
            import win32com.client

            if not self._initialized:
                pythoncom.CoInitialize()
                self._initialized = True

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = 0
            word.DisplayAlerts = 0
            self._word = word
            return True
        except ImportError:
            print(
                "Error: 'pywin32' is not installed. Required for .doc conversion."
            )
            return False
        except Exception as e:
            print(f"Error: Could not start Microsoft Word: {e}")
            return False

    def convert(self, input_path: str, output_path: str) -> bool:
        """Convert a single .doc file to .docx using the running Word instance."""
        if not self._ensure_com():
            return False

        wdFormatXMLDocument = 16
        try:
            doc = self._word.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=wdFormatXMLDocument)
            doc.Close(0)  # 0 = wdDoNotSaveChanges
            return True
        except Exception as e:
            print(f"  COM error converting file: {e}")
            # Try to recover by restarting Word
            self._restart()
            return False

    def _restart(self):
        """Kill and restart the Word instance after a failure."""
        try:
            if self._word is not None:
                self._word.Quit(0)
        except Exception:
            pass
        self._word = None
        # Small delay before retrying
        time.sleep(1)

    def quit(self):
        """Cleanly shut down Word and COM."""
        try:
            if self._word is not None:
                self._word.Quit(0)
        except Exception:
            pass
        self._word = None
        try:
            if self._initialized:
                import pythoncom

                pythoncom.CoUninitialize()
        except Exception:
            pass


def process_file(
    input_file: Path,
    output_file: Path,
    markitdown: MarkItDown,
    word_server: WordComServer,
    prefix: str = "",
):
    """Convert a single .doc or .docx file to markdown."""
    ext = input_file.suffix.lower()

    temp_docx = None
    try:
        target_file = input_file

        if ext == ".doc":
            # Create a temporary .docx file
            temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
            os.close(temp_fd)
            temp_docx = Path(temp_path)

            print(f"{prefix}Converting {input_file.name} (.doc -> .docx)...")
            success = word_server.convert(
                str(input_file.absolute()), str(temp_docx.absolute())
            )
            if not success:
                print(f"{prefix}SKIPPED (doc->docx conversion failed)")
                return False
            target_file = temp_docx

        print(f"{prefix}Converting to markdown...")
        result = markitdown.convert(str(target_file))

        output_file.parent.mkdir(parents=True, exist_ok=True)
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(result.markdown)

        print(f"{prefix}OK -> {output_file}")
        return True

    except Exception as e:
        print(f"{prefix}FAILED: {e}")
        return False
    finally:
        if temp_docx and temp_docx.exists():
            try:
                os.remove(temp_docx)
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(
        description="Batch convert .doc and .docx files to Markdown, maintaining folder structure.",
        prog="markitdown-batch",
    )

    parser.add_argument(
        "-v", "--version", action="version", version=f"%(prog)s {__version__}"
    )
    parser.add_argument(
        "input_dir", help="Input directory containing .doc/.docx files."
    )
    parser.add_argument("output_dir", help="Output directory for .md files.")
    parser.add_argument(
        "-p", "--use-plugins", action="store_true", help="Use 3rd-party plugins."
    )

    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()

    if not input_dir.is_dir():
        print(f"Error: Not a valid directory: {input_dir}")
        sys.exit(1)

    # ── Discover files ──────────────────────────────────────────────
    files_to_process = []
    for root, _dirs, files in os.walk(input_dir):
        for filename in files:
            file_ext = Path(filename).suffix.lower()
            if file_ext in (".doc", ".docx"):
                files_to_process.append(Path(root) / filename)

    total = len(files_to_process)
    if total == 0:
        print(f"No .doc or .docx files found in {input_dir}")
        return

    print(f"Input : {input_dir}")
    print(f"Output: {output_dir}")
    print(f"Found {total} file(s) to process.\n")

    # ── Set up converters ───────────────────────────────────────────
    markitdown = MarkItDown(enable_plugins=args.use_plugins)
    word_server = WordComServer()

    succeeded = 0
    failed = 0
    skipped = 0

    try:
        for idx, input_file in enumerate(files_to_process, start=1):
            rel_path = input_file.relative_to(input_dir)
            # Output: change extension to .md (strip .doc/.docx, add .md)
            output_file = output_dir / rel_path.with_suffix(".md")

            prefix = f"[{idx}/{total}] {rel_path.name} — "

            ok = process_file(input_file, output_file, markitdown, word_server, prefix)
            if ok:
                succeeded += 1
            else:
                failed += 1
    except KeyboardInterrupt:
        print("\n\nInterrupted by user.")
        skipped = total - succeeded - failed
    finally:
        word_server.quit()

    # ── Summary ─────────────────────────────────────────────────────
    print("\n" + "=" * 50)
    print("BATCH CONVERSION SUMMARY")
    print("=" * 50)
    print(f"  Total files : {total}")
    print(f"  Succeeded   : {succeeded}")
    print(f"  Failed      : {failed}")
    if skipped > 0:
        print(f"  Skipped     : {skipped}")
    print("=" * 50)


if __name__ == "__main__":
    main()
