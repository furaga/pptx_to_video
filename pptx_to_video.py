import argparse
from pathlib import Path
import yaml

from src.Project import Project

def parse_args():
    parser = argparse.ArgumentParser(description="")
   #  parser.add_argument("--pptx_path", type=Path, required=True)
    parser.add_argument("--config_path", type=Path, default="data/config.yml")
    parser.add_argument("--input_dir", type=Path, default=None)
    parser.add_argument("--output_path", type=Path, default=None)
    args = parser.parse_args()
    return args


def main(args):
    with open(args.config_path, "r", encoding="utf8") as f:
        config = yaml.safe_load(f)
    print(f"{config=}")

    if args.input_dir is not None:
        config["input"]["directory"] = args.input_dir

    if args.output_path is not None:
        config["output"]["filename"] = args.output_path

    project = Project(config["input"]["directory"])

    ok = project.export_video(config["output"]["filename"])
    if not ok:
        print("Failed to export video")
        return

    print("Successefully exported video to", config["output"]["filename"])

if __name__ == "__main__":
    main(parse_args())
