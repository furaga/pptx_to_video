import argparse
from pathlib import Path
from typing import List
import datetime

from src.Project import Project

def parse_args():
    parser = argparse.ArgumentParser(description="")
   #  parser.add_argument("--pptx_path", type=Path, required=True)
#    parser.add_argument("--workdir", type=Path, required=True)
#    parser.add_argument("--out_video_path", type=Path, required=True)
    parser.add_argument("--config_path", type=Path, default="data/config.yml")
    args = parser.parse_args()
    return args


def main(args):
    import yaml
    with open(args.config_path, "r", encoding="utf8") as f:
        config = yaml.safe_load(f)
    print(f"{config=}")

    project = Project(config["input"]["directory"])

    ok = project.export_video(config["output"]["filename"])
    if not ok:
        print("Failed to export video")
        return

    print("Successefully exported video to", config["output"]["filename"])

if __name__ == "__main__":
    main(parse_args())
