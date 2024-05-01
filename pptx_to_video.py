import argparse
from pathlib import Path
from typing import List
import datetime

from src.Project import Project

def parse_args():
    parser = argparse.ArgumentParser(description="")
   #  parser.add_argument("--pptx_path", type=Path, required=True)
    parser.add_argument("--workdir", type=Path, required=True)
    parser.add_argument("--out_video_path", type=Path, required=True)
    parser.add_argument("--config_path", type=Path, default=None)
    args = parser.parse_args()
    return args


def main(args):
    # workdir = Path("./__workdir__") / datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    workdir = args.workdir
    print("workdir:", workdir)

    project = Project(workdir)
    # ok = project.from_pptx(args.pptx_path)

    # if not ok:
    #     print("Failed to load pptx")
    #     return

    if args.config_path is not None:
        pass

    ok = project.export_video(args.out_video_path)
    if not ok:
        print("Failed to export video")
        return

    print("Successefully exported video to", args.out_video_path)

if __name__ == "__main__":
    main(parse_args())
