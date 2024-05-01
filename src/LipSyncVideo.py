from pathlib import Path
from typing import List, Tuple


class LipSyncVideo:
    def __init__(self, image_dir: Path = Path("data/zundamon")):
        self.image_dir = image_dir

    def create_video(
        self, duration: float, fps: float, speak_times: List[Tuple[float, float]]
    ):
        pass
