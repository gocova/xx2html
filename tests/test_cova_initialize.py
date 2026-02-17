import json
import tempfile
import unittest
from pathlib import Path

from xx2html.cova import initialize


class CovaInitializeTests(unittest.TestCase):
    def test_initialize_returns_merged_config_without_mutating_default(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir)
            config_file = config_path / "config.json"
            config_file.write_text(json.dumps({"user": True, "a": 2}), encoding="utf-8")

            default_config = {"a": 1, "nested": {"x": 1}}
            system_config = {"sys": True}

            result = initialize(
                app_config_path=str(config_path),
                config_filename="config.json",
                default_config=default_config,
                system_config=system_config,
            )

            self.assertEqual(1, default_config["a"])
            self.assertEqual({"x": 1}, default_config["nested"])

            self.assertEqual(
                {"a": 2, "nested": {"x": 1}, "user": True, "sys": True},
                result,
            )


if __name__ == "__main__":
    unittest.main()
