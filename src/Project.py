import shutil
from pathlib import Path

class Project:
    def __init__(self, workdir: Path) -> None:
        self.workdir = workdir
        self.workdir.mkdir(parents=True, exist_ok=True)
        print("Created workdir:", self.workdir)

    def close(self):
        shutil.rmtree(self.workdir)
        print("Removed workdir:", self.workdir)
        
    def from_pptx(self, pptx_path: Path) -> bool:
        # TODO: Implement this method
        return True

    def export_video(self, out_video_path: Path) -> bool:
        all_img_paths = list(self.workdir.glob("slides/*.png"))
        all_manuscrpt_paths = list(self.workdir.glob("slides/*.txt"))
        for img_path, manuscrpt_path in zip(all_img_paths, all_manuscrpt_paths):
            # TODO: Implement this method
            print(img_path, manuscrpt_path)
        return True