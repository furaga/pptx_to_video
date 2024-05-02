import shutil
from typing import Dict
from pathlib import Path
import moviepy.editor
import comtypes.client
import uuid
import os

from .TextToSpeech import TextToSpeech


class Project:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self.config["input"]["type"] == "pptx":
             self.close()

        # 例外は伝搬させたいので False を返す
        return False

    def __init__(self, config: Dict) -> None:
        self.config = config

        if config["input"]["type"] == "pptx":
            self.pptx_path = Path(config["input"]["path"])
            self.workdir = self.pptx_path.with_suffix("")
            self.workdir.mkdir(parents=True, exist_ok=True)
            self.import_pptx()
        else:
            self.workdir = Path(config["input"]["path"])
            self.workdir.mkdir(parents=True, exist_ok=True)

        self.tts = TextToSpeech()

    def close(self):
        shutil.rmtree(self.workdir)

    def import_pptx(self) -> bool:
        try:
            application = comtypes.client.CreateObject("Powerpoint.Application")
            pptx_path = str(self.pptx_path.resolve())
            print(f"{pptx_path=}")
            presentation = application.Presentations.open(pptx_path)

            presentation.Export(str(self.workdir.resolve()), FilterName="png")
            print("Exported images to", self.workdir)
            for slide in presentation.Slides:
                # Extract notes from slide
                notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
                notes = notes.replace("\r", "\n")
                (self.workdir / f"スライド{slide.SlideIndex}.txt").write_text(
                    notes, encoding="utf8"
                )
                print(
                    "Exported note to", self.workdir / f"スライド{slide.SlideIndex}.txt"
                )

            presentation.close()
            application.quit()

            return True
        except Exception as e:
            import traceback

            print(e, "\n", traceback.format_exc())
            return False

    def make_clip(
        self,
        img_path: Path,
        manuscrpt: str,
        fps: float,
        manuscript_margin: float,
        line_interval: float,
    ) -> None:
        lines = manuscrpt.split("\n")
        lines = [line.strip() for line in lines if len(line.strip()) > 0]

        all_clips = []

        if len(lines) <= 0:
            lines = ["（セリフが未設定です）"]

        for i, line in enumerate(lines):
            wav = self.tts.tts(line, speed=1.1, speaker=3)
            # Save audio to temporary file
            audio_path = self.workdir / f"{str(uuid.uuid4())}.wav"
            audio_path.write_bytes(wav)

            try:
                # Load audio clip
                audio_clip = moviepy.editor.AudioFileClip(str(audio_path))
                audio_duration = len(wav) // (2 * 24000)

                # Load image clip
                start = line_interval / 2 if i > 0 else manuscript_margin
                end = line_interval / 2 if i < len(lines) - 1 else manuscript_margin
                video_clip = moviepy.editor.ImageClip(
                    str(img_path), duration=audio_duration + start + end
                )
                video_clip.fps = fps
                video_clip.audio = moviepy.editor.CompositeAudioClip(
                    [audio_clip.set_start(start)]
                )

                # Create text clip
                fontsize = video_clip.size[0] // 30
                txt_clip = moviepy.editor.TextClip(
                    line,
                    fontsize=fontsize,
                    color="green",
                    font=os.environ["MANUSCRIPTS_FONT"],
                )
                txt_clip.duration = video_clip.duration
                txt_clip = txt_clip.set_position(("center", 0.9), relative=True)

                # Composite clips
                clip = moviepy.editor.CompositeVideoClip([video_clip, txt_clip])
                clip.duration = video_clip.duration

                # Append clip
                all_clips.append(clip)

            except Exception as e:
                import traceback

                print(e, "\n", traceback.format_exc())
            finally:
                audio_path.unlink()

        return moviepy.editor.concatenate_videoclips(all_clips)

    def export_video(self) -> bool:
        try:
            # Get all image and manuscript paths
            all_img_paths = list(self.workdir.glob("*.png"))
            all_manuscrpt_paths = list(self.workdir.glob("*.txt"))

            # Make clips for slides
            all_clips = []

            all_pairs = list(zip(all_img_paths, all_manuscrpt_paths))

            from tqdm import tqdm

            for img_path, manuscrpt_path in tqdm(all_pairs):
                clip = self.make_clip(
                    img_path,
                    manuscrpt_path.read_text(encoding="utf8"),
                    self.config["output"]["fps"],
                    self.config["output"]["manuscript_slide_margin"],
                    self.config["output"]["manuscript_line_interval"],
                )
                all_clips.append(clip)

            if len(all_clips) <= 0:
                print("No clips to concatenate")
                return False

            # Concatenate clips
            video = moviepy.editor.concatenate_videoclips(all_clips)

            # Export video
            video.write_videofile(self.config["output"]["path"])
            return True
        except Exception as e:
            import traceback

            print(e, "\n", traceback.format_exc())
            return False
