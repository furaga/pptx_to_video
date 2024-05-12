import yaml
import streamlit as st
from pathlib import Path
import uuid
import sys
import shutil
import datetime


sys.path.append(str(Path(__file__).parent.parent))
from src.Project import Project

def on_progress(progress_bar, percentage: float, status: str):
    progress_bar.progress(percentage, status)

def remove_old_tmp_dirs():
    for d in Path("tmp").iterdir():
        if not d.is_dir():
            continue

        try:
            dt = datetime.datetime.strptime(d.name, "%Y%m%d-%H%M%S")
            if (datetime.datetime.now() - dt).days > 1:
                shutil.rmtree(d)
                print(f"Removed {d}")
        except ValueError:
            continue

def main():
    st.title("PowerPoint to Video Converter")

    pptx_file = st.file_uploader("Upload a pptx file", type=["pptx"])

    if st.button("Convert to Video"):
        if pptx_file is not None:
            content = pptx_file.read()
            st.write(f"Converting to video... (filesize={len(content)}")

            with open("samples/from_pptx/config.yml", "r", encoding="utf8") as f:
                config = yaml.safe_load(f)

            dirname = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            input_path = Path(f"tmp/{dirname}/{pptx_file.name}")
            input_path.parent.mkdir(exist_ok=True, parents=True)
            input_path.write_bytes(content)
            config["input"]["path"] = str(input_path)

            output_path = input_path.parent / f"{pptx_file.name}.mp4"
            config["output"]["path"] = str(output_path)

            progress_bar = st.progress(0, "Exporting a video...")

            with Project(config, user_import_server=True) as project:
                project.export_video(lambda v, t: on_progress(progress_bar, v, t))

            input_path.unlink()

            if output_path.exists():
                progress_bar.progress(100, "Successfully exported a video!")
                st.download_button(
                    label=f"Download {output_path.name}",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime="video/mp4",
                )
            else:
                progress_bar.progress(100, "Failed to export a video...")



if __name__ == "__main__":
    import dotenv

    dotenv.load_dotenv()

    main()
