from typing import Dict
from pathlib import Path
import moviepy.editor
import moviepy.video.fx.resize
import comtypes.client
import uuid
import yaml
import os
import time
import traceback

from .TextToSpeech import TextToSpeech


class Project:
    def __enter__(self):
        self.errors = ""
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        if self.config["input"]["type"] == "pptx":
            self.close()

        # 例外は伝搬させたいので False を返す
        return False

    def __init__(self, config: Dict, user_import_server: bool = False) -> None:
        self.config = config

        if config["input"]["type"] == "pptx":
            self.pptx_path = Path(config["input"]["path"])
            self.workdir = self.pptx_path.with_suffix("")
            self.workdir.mkdir(parents=True, exist_ok=True)
            if user_import_server:
                self.request_to_pptx_import_server()
            else:
                self.import_pptx()
        else:
            self.workdir = Path(config["input"]["path"])
            self.workdir.mkdir(parents=True, exist_ok=True)

        self.speaker_id = config["output"].get("speaker_id", 3)
        self.speaker_speed = config["output"].get("speaker_speed", 1.1)
        self.voicevox_url = config["output"].get("voicevox_url", "http://127.0.0.1:50021")
        self.tts = TextToSpeech(voicevox_url=self.voicevox_url)
        self.errors = ""

    def close(self):
        #   shutil.rmtree(self.workdir)
        pass

    def request_to_pptx_import_server(self) -> bool:
        config_path = Path(f"__pptx_import_server__watch_dir__/{str(uuid.uuid4())}.yml")
        config_path.parent.mkdir(exist_ok=True, parents=True)
        with open(config_path, "w", encoding="utf8") as f:
            yaml.safe_dump(self.config, f)

        while not Path(self.workdir / "__pptx_imported__").exists():
            time.sleep(1)

        print("[request_to_pptx_import_server] PPTX imported")
        return True

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

            (self.workdir / "__pptx_imported__").touch()
            return True
        except Exception as e:
            import traceback

            print(e, "\n", traceback.format_exc())
            return False

    def log_error(self, msg):
        self.errors += msg + "\n"

    def parse_command(self, line: str):
        line = line.strip()
        line = line.replace(" ", "")

        # <video:xxx.mp4>
        if line.startswith("<video:") and ">" in line:
            p_end = line.find(">")
            video_path = line[len("<video:") : p_end]
            return "insert_video", self.workdir.parent / video_path

        # <wait:waittime>
        if line.startswith("<wait:") and ">" in line:
            p_end = line.find(">")
            try:
                wait_time = float(line[len("<wait:") : p_end])
            except Exception:
                self.log_error(f"Failed to parse wait command {line}")
            return "wait", wait_time

        # <speaker:id=yy,speed=zz>
        if line.startswith("<speaker:") and ">" in line:
            p_end = line.find(">")
            speaker = line[len("<speaker:") : p_end]
            id, speed = None, None
            try:
                for cfg in speaker.split(","):
                    if "=" not in cfg:
                        continue
                    key, value = cfg.split("=")[:2]
                    if key == "id":
                        id = int(value)
                    if key == "speed":
                        speed = float(value)
            except Exception:
                self.log_error(f"Failed to parse speaker command {line}")

            return "speaker", (id, speed)

        return "", None

    def make_clip(
        self,
        img_path: Path,
        manuscript: str,
        fps: float,
        manuscript_margin: float,
        line_interval: float,
        fontsize_ratio: float,
        fontcolor: str,
    ) -> None:
        cmd, args = self.parse_command(manuscript)
        if cmd == "insert_video":
            video_path = args
            video_clip = moviepy.editor.VideoFileClip(str(video_path))
            ref_img_clip = moviepy.editor.ImageClip(str(img_path))
            video_clip = video_clip.resize(ref_img_clip.size)
            return video_clip

        lines = manuscript.split("\n")
        lines = [line.strip() for line in lines if len(line.strip()) > 0]

        all_clips = []

        if len(lines) <= 0:
            # 台本未設定の場合
            return moviepy.editor.ImageClip(str(img_path), duration=5.0)

        for i, line in enumerate(lines):
            speaker_id = self.speaker_id
            speaker_speed = self.speaker_speed

            cmd, args = self.parse_command(line)
            if cmd == "speaker":
                new_speaker, new_speed = args
                if new_speaker is not None:
                    speaker_id = new_speaker
                if new_speed is not None:
                    speaker_speed = new_speed

            wait_time = 0
            if cmd == "wait":
                wait_time = args

            # remove <speaker:xx>
            tag_start = line.find("<speaker:")
            tag_end = line.rfind(">")
            if tag_start >= 0 and tag_end >= 0:
                line = line[:tag_start] + line[tag_end + 1 :]

            tag_start = line.find("<wait:")
            tag_end = line.rfind(">")
            if tag_start >= 0 and tag_end >= 0:
                line = line[:tag_start] + line[tag_end + 1 :]

            if line.strip() == "":
                continue

            wav = self.tts.tts(line, speed=speaker_speed, speaker=speaker_id)

            # Save audio to temporary file
            audio_path = self.workdir / f"{str(uuid.uuid4())}.wav"
            audio_path.write_bytes(wav)

            # Load audio clip
            audio_clip = moviepy.editor.AudioFileClip(str(audio_path))
            # 音の最後のノイズが乗ることがあるので除去
            audio_clip.duration = audio_clip.duration - 0.01
            audio_duration = audio_clip.duration

            # Load image clip
            start = line_interval / 2 if i > 0 else manuscript_margin
            start += wait_time
            end = line_interval / 2 if i < len(lines) - 1 else manuscript_margin
            video_clip = moviepy.editor.ImageClip(
                str(img_path), duration=audio_duration + start + end
            )
            video_clip.fps = fps
            video_clip.audio = moviepy.editor.CompositeAudioClip(
                [audio_clip.set_start(start)]
            )

            # Create text clip
            fontsize = int(video_clip.size[0] * fontsize_ratio)
            print("========", line, "==========")
            txt_clip = moviepy.editor.TextClip(
                line,
                fontsize=fontsize,
                color=fontcolor,
                font=os.environ["MANUSCRIPTS_FONT"],
            )
            txt_clip.duration = video_clip.duration
            txt_clip = txt_clip.set_position(("center", 0.9), relative=True)

            # Composite clips
            clip = moviepy.editor.CompositeVideoClip([video_clip, txt_clip])
            clip.duration = video_clip.duration

            # Append clip
            all_clips.append(clip)

        return moviepy.editor.concatenate_videoclips(all_clips)

    def export_video(self, on_progress=None) -> bool:
        try:
            # Get all image and manuscript paths
            all_img_paths = list(self.workdir.glob("*.png"))
            all_manuscrpt_paths = list(self.workdir.glob("*.txt"))

            # Make clips for slides
            all_clips = []

            all_pairs = list(zip(all_img_paths, all_manuscrpt_paths))

            def _is_int(s: str):
                try:
                    _ = int(s)
                    return True
                except Exception:
                    return False

            all_pairs = [(p, m) for p, m in all_pairs if _is_int(p.stem[4:])]

            # スライドXXのXXの数で並び替える
            all_pairs = sorted(all_pairs, key=lambda pm: int(pm[0].stem[4:]))

            from tqdm import tqdm

            for i, (img_path, manuscrpt_path) in tqdm(enumerate(all_pairs)):
                if on_progress is not None:
                    on_progress(
                        (i + 1) / (len(all_pairs) + 2),
                        f"Exporting a slide {img_path.name}",
                    )
                clip = self.make_clip(
                    img_path,
                    manuscrpt_path.read_text(encoding="utf8"),
                    self.config["output"].get("fps", 30),
                    self.config["output"].get("manuscript_slide_margin", 1.0),
                    self.config["output"].get("manuscript_line_interval", 0.5),
                    self.config["output"].get("fontsize_ratio", 0.025),
                    self.config["output"].get("font_color", "green"),
                )
                all_clips.append(clip)

            if len(all_clips) <= 0:
                print("No clips to concatenate")
                return False

            # Concatenate clips
            video = moviepy.editor.concatenate_videoclips(all_clips)

            # Export video
            if on_progress is not None:
                on_progress(
                    (len(all_pairs) + 1) / (len(all_pairs) + 2), "Writing a video file"
                )
            video.write_videofile(self.config["output"]["path"], audio_codec="aac")
            return True
        except Exception as e:
            self.log_error(f"{e}\n{traceback.format_exc()}\n")
            return False
