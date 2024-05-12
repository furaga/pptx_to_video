import argparse
from pathlib import Path
import yaml
import time

from src.Project import Project


def parse_args():
    parser = argparse.ArgumentParser(description="")
    parser.add_argument(
        "--watch_dir", type=Path, default="__pptx_import_server__watch_dir__"
    )
    args = parser.parse_args()
    return args


def main(args):
    while True:
        all_config_paths = Path(args.watch_dir).glob("*.yml")
        for config_path in all_config_paths:
            try:
                with open(config_path, "r", encoding="utf8") as f:
                    config = yaml.safe_load(f)
                Project(config)
            except Exception as e:
                import traceback

                print(e, traceback.format_exc())
            finally:
                config_path.unlink()

        time.sleep(1)


if __name__ == "__main__":
    import dotenv

    dotenv.load_dotenv()
    main(parse_args())
