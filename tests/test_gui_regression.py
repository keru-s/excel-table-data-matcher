from __future__ import annotations

import os
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class GuiRegressionTest(unittest.TestCase):
    def test_gui_smoke_flow(self) -> None:
        env = os.environ.copy()
        env["QT_QPA_PLATFORM"] = "offscreen"

        result = subprocess.run(
            [sys.executable, str(ROOT / "gui_smoke_test.py")],
            cwd=ROOT,
            env=env,
            capture_output=True,
            text=True,
            timeout=30,
        )

        if result.returncode != 0:
            self.fail(
                "GUI smoke test failed\n"
                f"stdout:\n{result.stdout}\n"
                f"stderr:\n{result.stderr}"
            )

        self.assertIn("output_rows", result.stdout)
        self.assertIn("导出完成", result.stdout)


if __name__ == "__main__":
    unittest.main()
