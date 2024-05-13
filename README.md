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

3. [ImageMagick](https://imagemagick.org/script/download.php)をインストール

4. .env.sampleを参考に.envを作成

4. PowerPointをインストールする（Mac版は不要）

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


## 特殊コマンド

1. ノート冒頭に `<video:動画ファイル名>` を指定すると、streamlit UIの "動画素材" で指定した動画をスライドに挿入できます

2. ノートの各行の冒頭に `<speaker:id=xxx,speed=yyy>` をつけると、その行を読み上げる音声ID・スピードを変更できます


# Mac環境で動かしたい場合

Mac環境ではpptxファイルから画像・台本抽出する処理が動きません。  
一方で、 `samples/from_png_txt/slides` のように画像(.png)と台本(.txt)を用意して、そこから動画生成することはできます。
（powerpointの画像書き出し機能を作ると用意しやすい）

```
python pptx_to_video.py --config_path samples/from_png_txt/config.yml 
```
