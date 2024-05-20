import argparse
from pathlib import Path
import yaml
import dotenv

dotenv.load_dotenv()
from src.Project import Project


def parse_args():
    parser = argparse.ArgumentParser(description="")
    parser.add_argument(
        "--config_path", type=Path, default="samples/from_pptx/config.yml"
    )
    parser.add_argument("--pptx_path", type=Path, default=None)
    parser.add_argument("--png_txt_dir", type=Path, default=None)
    parser.add_argument("--output_path", type=Path, default=None)
    args = parser.parse_args()
    return args


def main(args):
    # Load config
    with open(args.config_path, "r", encoding="utf8") as f:
        config = yaml.safe_load(f)

    # Overwrite config with args
    if config["input"]["type"] == "png_txt" and args.png_txt_dir is not None:
        config["input"]["path"] = args.png_txt_dir

    if config["input"]["type"] == "pptx" and args.pptx_path is not None:
        config["input"]["path"] = args.pptx_path

    if args.output_path is not None:
        config["output"]["filename"] = args.output_path

    print("config:", config)

    # Create project
    with Project(config) as project:
        # Export video
        print(f"Exporting video... (working directory: {project.workdir})")
        ok = project.export_video()
        if not ok:
            print("Failed to export video")
            return
        print("Successefully exported video to", config["output"]["path"])


if __name__ == "__main__":
    main(parse_args())
