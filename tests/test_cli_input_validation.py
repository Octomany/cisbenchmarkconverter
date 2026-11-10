"""CLI input-validation tests for the CIS Benchmark Converter."""

import subprocess
import sys
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent.parent / "cis_benchmark_converter.py"


def run_cli(*args):
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        capture_output=True,
        text=True,
    )


class InputValidationTestCase(unittest.TestCase):
    def test_missing_input_file_exits_with_usage_error(self):
        # A missing input PDF must be rejected with a clean usage error
        # (exit code 2) instead of an unhandled traceback.
        result = run_cli("-i", "does-not-exist.pdf", "-f", "json")
        self.assertEqual(result.returncode, 2)
        self.assertIn("Input file does not exist", result.stderr)
        self.assertNotIn("Traceback", result.stderr)


if __name__ == "__main__":
    unittest.main()
