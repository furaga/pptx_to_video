import shutil
from pathlib import Path
import moviepy.editor
from .TextToSpeech import TextToSpeech

class Project:
    def __init__(self, workdir: Path) -> None:
        self.workdir = Path(workdir)
        self.workdir.mkdir(parents=True, exist_ok=True)
        self.tts = TextToSpeech()
        print("Created workdir:", self.workdir)

    def close(self):
        shutil.rmtree(self.workdir)
        print("Removed workdir:", self.workdir)
        
    def from_pptx(self, pptx_path: Path) -> bool:
        # TODO: Implement this method
        return True
    
    def make_clip(self, img_path: Path, manuscrpt: str, fps: float) -> None:
        manuscript_margin = 0.8
        line_interval = 0.4
        lines = manuscrpt.split("\n")
        lines = [line.strip() for line in lines if len(line.strip()) > 0]

        all_clips = []
        for i, line in enumerate(lines):
            wav = self.tts.tts(line, speed=1.1, speaker=3)

            # generate random string
            import uuid
            audio_path = self.workdir / f"{str(uuid.uuid4())}.wav"
            audio_path.write_bytes(wav)

            try:
                wav_clip = moviepy.editor.AudioFileClip(str(audio_path))
                wav_duration = len(wav) // (2 * 24000)

                start = line_interval / 2 if i > 0 else manuscript_margin
                end = line_interval / 2 if i < len(lines) - 1 else manuscript_margin

                video_clip = moviepy.editor.ImageClip(str(img_path), duration=wav_duration + start + end)
                video_clip.fps = fps
                video_clip.audio = moviepy.editor.CompositeAudioClip([wav_clip.set_start(start)])

                fontsize = video_clip.size[0] // 30
                txt_clip = moviepy.editor.TextClip(line, fontsize=fontsize, color="green", font="C:/Windows/Fonts/msgothic.ttc")
                txt_clip.duration = video_clip.duration
                txt_clip = txt_clip.set_position(('center', 0.85), relative=True)
                
                clip = moviepy.editor.CompositeVideoClip([video_clip, txt_clip])
                clip.duration = video_clip.duration

                all_clips.append(clip)

            except Exception as e:
                import traceback
                print(e, "\n", traceback.format_exc())
            finally:
                audio_path.unlink()

        
        return moviepy.editor.concatenate_videoclips(all_clips)
        
    def export_video(self, out_video_path: Path, fps: float=10) -> bool:
        try:
            all_img_paths = list(self.workdir.glob("*.png"))
            all_manuscrpt_paths = list(self.workdir.glob("*.txt"))
            print(f"{len(all_img_paths)=}, {len(all_manuscrpt_paths)=}")

            all_clips = []
            for img_path, manuscrpt_path in zip(all_img_paths, all_manuscrpt_paths):
                clip = self.make_clip(img_path, manuscrpt_path.read_text(encoding="utf8"), fps)
                all_clips.append(clip)
                print(img_path, manuscrpt_path)

            if len(all_clips) <= 0:
                print("No clips to concatenate")
                return False

            video = moviepy.editor.concatenate_videoclips(all_clips)
            video.write_videofile(str(out_video_path))
            return True
        except Exception as e:
            import traceback
            print(e, "\n", traceback.format_exc())
            return False