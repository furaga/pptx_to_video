import yaml
import streamlit as st
from pathlib import Path
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


def save_uploaded_file(file, save_dir):
    if file is None:
        return None
    content = file.read()
    save_path = save_dir / file.name
    save_path.parent.mkdir(exist_ok=True, parents=True)
    save_path.write_bytes(content)
    return save_path


def main():
    st.title("PowerPoint to Video Converter")

    pptx_file = st.file_uploader("パワーポイント(.pptx)", type=["pptx"])

    additional_vidoe_files = st.file_uploader(
        "動画素材（.mp4）(<video>コマンドで指定する動画)",
        type=["mp4", "mkv"],
        accept_multiple_files=True,
    )

    if st.button("Convert to Video"):
        if pptx_file is not None:
            st.write("Converting to video...")

            with open("samples/from_pptx/config.yml", "r", encoding="utf8") as f:
                config = yaml.safe_load(f)

            tmp_dir = Path(f'tmp/{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}')
            input_path = save_uploaded_file(pptx_file, tmp_dir)
            config["input"]["path"] = str(input_path)

            if additional_vidoe_files is not None:
                for f in additional_vidoe_files:
                    save_uploaded_file(f, tmp_dir)

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
