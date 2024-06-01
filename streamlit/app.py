import yaml
import streamlit as st
from pathlib import Path
import sys
import shutil
import datetime
import dotenv

dotenv.load_dotenv(".env")
import moviepy.editor

sys.path.append(str(Path(__file__).parent.parent))
from src.Project import Project


@st.cache_resource
def get_manuscript_colors():
    colors = moviepy.editor.TextClip.list("color")[3:]
    colors = [c.decode() for c in colors]
    print(f"{len(colors)} font color options were found.")
    return colors


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

    with st.sidebar:
        st.markdown("""
    # 概要
    パワーポイント(.pptx)を動画化するツールです。  
    各スライドのノートをVOICEVOXで読み上げた動画を生成します。
                    
    # 使い方
    1. "Drag and drop file here"に動画化したいpptxファイルを指定
    2. "変換"ボタンを押す  
                
    変換成功すると生成動画(.mp4)のダウンロードボタンが表示されます  
    また、"詳細設定"で字幕などの設定を変更できます

    # 特殊コマンド
    ```
    <video:{video_name}.mp4> 
    ノート冒頭に記載すると、そのスライドの代わりに {video_name}.mp4 が挿入される 
    {video_name}.mp4 は、"詳細設定>動画素材（.mp4）(<video>コマンドで指定する動画)"で指定した動画名と一致している必要があります
    
    <wait:{wait_time_in_sec}> 
    ノートの各行頭に記載すると、読み上げ開始が通常より {wait_time_in_sec} 秒遅延される
                    
    <speaker:id={speaker_id},speed={speak_speed}>
    ノートの各行頭に記載すると、一時的にVOICEVOXのspeaker id, 読み上げ速度が変更される
    ```

    # 実装
    https://github.com/furaga/pptx_to_video
    """)

    with st.expander("詳細設定", expanded=False):
        st.title("詳細設定")
        fontsize_percentage = st.number_input("字幕サイズ（フォントサイズ/スライド横幅[%]）", min_value=0.0, max_value=100.0, value=2.5)
        color_options = get_manuscript_colors()
        fontcolor = st.selectbox("字幕の色", color_options, color_options.index("green"))

        voicevox_url = st.text_input("読み上げAPIのURL", value="http://127.0.0.1:50021")
        speaker_id = st.number_input("話者ID", min_value=0, max_value=100, value=3)
        speaker_speed = st.number_input("読み上げ速度", min_value=0.01, max_value=100.0, value=1.0)

        manuscript_slide_margin = st.number_input("スライド間の読み上げの余白[秒]", min_value=0.0, max_value=100.0, value=1.5)
        manuscript_line_interval = st.number_input("行間の読み上げの余白[秒]", min_value=0.0, max_value=100.0, value=0.5)
        additional_vidoe_files = st.file_uploader(
            "動画素材（.mp4）(<video>コマンドで指定する動画)",
            type=["mp4", "mkv"],
            accept_multiple_files=True,
        )
       # with st.expander("話者ID一覧", expanded=False):
        st.write("""参考：話者ID一覧
|ID|名前|スタイル|
|---|---|---|
|0|四国めたん|あまあま|
|1|ずんだもん|あまあま|
|2|四国めたん|ノーマル|
|3|ずんだもん|ノーマル|
|4|四国めたん|セクシー|
|5|ずんだもん|セクシー|
|6|四国めたん|ツンツン|
|7|ずんだもん|ツンツン|
|8|春日部つむぎ|ノーマル|
|9|波音リツ|ノーマル|
|10|雨晴はう|ノーマル|
|11|玄野武宏|ノーマル|
|12|白上虎太郎|ふつう|
|13|青山龍星|ノーマル|
|14|冥鳴ひまり|ノーマル|
|15|九州そら|あまあま|
|16|九州そら|ノーマル|
|17|九州そら|セクシー|
|18|九州そら|ツンツン|
|19|九州そら|ささやき|
|20|もち子さん|ノーマル|
|21|剣崎雌雄|ノーマル|
|22|ずんだもん|ささやき|
|23|WhiteCUL|ノーマル|
|24|WhiteCUL|たのしい|
|25|WhiteCUL|かなしい|
|26|WhiteCUL|びえーん|
|27|後鬼|人間ver.|
|28|後鬼|ぬいぐるみver.|
|29|No.7|ノーマル|
|30|No.7|アナウンス|
|31|No.7|読み聞かせ|
|32|白上虎太郎|わーい|
|33|白上虎太郎|びくびく|
|34|白上虎太郎|おこ|
|35|白上虎太郎|びえーん|
|36|四国めたん|ささやき|
|37|四国めたん|ヒソヒソ|
|38|ずんだもん|ヒソヒソ|
|39|玄野武宏|喜び|
|40|玄野武宏|ツンギレ|
|41|玄野武宏|悲しみ|
|42|ちび式じい|ノーマル|
|43|櫻歌ミコ|ノーマル|
|44|櫻歌ミコ|第二形態|
|45|櫻歌ミコ|ロリ|
|46|小夜/SAYO|ノーマル|
|47|ナースロボ＿タイプＴ|ノーマル|
|48|ナースロボ＿タイプＴ|楽々|
|49|ナースロボ＿タイプＴ|恐怖|
|50|ナースロボ＿タイプＴ|内緒話|
|51|†聖騎士 紅桜†|ノーマル|
|52|雀松朱司|ノーマル|
|53|麒ヶ島宗麟|ノーマル|
                 
""")        

    if st.button("変換"):
        remove_old_tmp_dirs()

        st.session_state["output_path"] = None

        if pptx_file is not None:
            with st.spinner("Converting to video..."):
                # ベースの設定
                with open("samples/from_pptx/config.yml", "r", encoding="utf8") as f:
                    config = yaml.safe_load(f)

                tmp_dir = Path(f'tmp/{datetime.datetime.now().strftime("%Y%m%d-%H%M%S")}')
                input_path = save_uploaded_file(pptx_file, tmp_dir)
                config["input"]["path"] = str(input_path)

                if additional_vidoe_files is not None:
                    for f in additional_vidoe_files:
                        save_uploaded_file(f, tmp_dir)

                output_path = input_path.parent / f"{pptx_file.name}.mp4"
                config["output"]["voicevox_url"] = voicevox_url
                config["output"]["speaker_id"] = speaker_id
                config["output"]["speaker_speed"] = speaker_speed
                config["output"]["fontsize_ratio"] = fontsize_percentage / 100
                config["output"]["font_color"] = fontcolor
                config["output"]["manuscript_slide_margin"] = manuscript_slide_margin
                config["output"]["manuscript_line_interval"] = manuscript_line_interval
                config["output"]["path"] = str(output_path)

                progress_bar = st.progress(0, "Exporting a video...")

                error_logs = ""
                with Project(config, user_import_server=True) as project:
                    project.export_video(lambda v, t: on_progress(progress_bar, v, t))
                    error_logs = project.errors

                input_path.unlink()
    
                if output_path.exists():
                    progress_bar.progress(100, "Successfully exported a video!")
                    st.session_state["output_path"] = output_path
                else:
                    progress_bar.progress(100, "Failed to export a video...")
                    st.error(error_logs)

    if "output_path" in st.session_state and st.session_state["output_path"] is not None:
        output_path = st.session_state["output_path"]
        if output_path.exists():
            st.download_button(
                label=f"Download {output_path.name}",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="video/mp4",
            )


if __name__ == "__main__":
    main()
