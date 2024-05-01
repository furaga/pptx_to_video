import shutil
from pathlib import Path
import moviepy.editor
import simpleaudio
from .TextToSpeech import TextToSpeech

class Project:
    def __init__(self, workdir: Path) -> None:
        self.workdir = workdir
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
        wav = self.tts.tts(manuscrpt, speed=1.1, speaker=3)
        audio_path = self.workdir / "__wav__.wav"
        audio_path.write_bytes(wav)
        wav_clip = moviepy.editor.AudioFileClip(str(audio_path))
        wav_duration = len(wav) // (2 * 24000)

        audio_margin = 1.0

        clip = moviepy.editor.ImageClip(str(img_path), duration=wav_duration + audio_margin * 2)
        clip.fps = fps
        clip.audio = moviepy.editor.CompositeAudioClip([wav_clip.set_start(audio_margin)])
        return clip
        
    def export_video(self, out_video_path: Path, fps: float=10) -> bool:
        try:
            all_img_paths = list(self.workdir.glob("slides/*.png"))
            all_manuscrpt_paths = list(self.workdir.glob("slides/*.txt"))

            all_clips = []
            for img_path, manuscrpt_path in zip(all_img_paths, all_manuscrpt_paths):
                clip = self.make_clip(img_path, manuscrpt_path.read_text(encoding="utf8"), fps)
                all_clips.append(clip)
                print(img_path, manuscrpt_path)

            video = moviepy.editor.concatenate_videoclips(all_clips)
            video.write_videofile(str(out_video_path))
            return True
        except Exception as e:
            import traceback
            print(e, "\n", traceback.format_exc())
            return False