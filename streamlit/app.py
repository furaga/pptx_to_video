import yaml
import streamlit as st
from pathlib import Path
import uuid
import sys

sys.path.append(str(Path(__file__).parent.parent))
from src.Project import Project


def main():
    st.title("My Streamlit App")
    st.write("Welcome to my app!")

    pptx_file = st.file_uploader("Upload a file", type=["pptx"])

    if st.button("Convert to Video"):
        if pptx_file is not None:
            content = pptx_file.read()
            st.write("Converting to video...", len(content))

            with open("samples/from_pptx/config.yml", "r", encoding="utf8") as f:
                config = yaml.safe_load(f)

            input_path = Path(f"tmp/{uuid.uuid4()}/{pptx_file.name}")
            input_path.parent.mkdir(exist_ok=True, parents=True)
            input_path.write_bytes(content)
            config["input"]["path"] = str(input_path)

            output_path = input_path.parent / f"{pptx_file.name}.mp4"
            config["output"]["path"] = str(output_path)
            with Project(config) as project:
                project.export_video()
            input_path.unlink()

            if output_path.exists():
                st.download_button(
                    label=f"Download Video {output_path.name}",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime="video/mp4",
                )



if __name__ == "__main__":
    import dotenv

    dotenv.load_dotenv()

    main()
