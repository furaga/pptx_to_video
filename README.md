# pptx_to_video

パワーポイント（.pptx）からゆっくり音声付き動画を自動生成するツール  
各スライドのノートを台本として読み上げた動画生成を行います 

# Setup

1. 以下のコマンドでpythonをセットアップしてください（
python 3.11.9で動作確認）


```
pip install -r requirements.txt
```

2. VOICEVOXをローカルで起動してください

# How to use

# Streamlit App (Windows)

1. 以下の２つのコマンドを別ターミナルで起動してください

```
python pptx_import_server.py
```

```
streamlit run streamlit/app.py
```

2. http://localhost:8501/ にアクセスしてください

3. 変換したいPPTXファイルを選択して、"Convert to Video" をクリックしてください

4. 変換成功したら出力動画のダウンロードボタンが表示されます


# CLIツール（Windows, Mac）

