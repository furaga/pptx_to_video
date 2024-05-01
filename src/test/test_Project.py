import unittest
from pathlib import Path
import yaml

from src.Project import Project

sample_dir = Path(__file__).parent.parent.parent / "samples"

class TestProject(unittest.TestCase):
    def test_png_txt(self):
        with open(sample_dir / "from_png_txt/config.yml", "r", encoding="utf8") as f:
            config = yaml.safe_load(f)

        with Project(config) as project:
            out_path = Path(project.config["output"]["path"])
            out_path.unlink(missing_ok=True)
            project.export_video()
            self.assertTrue(out_path.exists())
            out_path.unlink()

    def test_pptx(self):
        with open(sample_dir / "from_pptx/config.yml", "r", encoding="utf8") as f:
            config = yaml.safe_load(f)

        with Project(config) as project:
            out_path = Path(project.config["output"]["path"])
            out_path.unlink(missing_ok=True)
            project.export_video()
            self.assertTrue(out_path.exists())
            out_path.unlink()

