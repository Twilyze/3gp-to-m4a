## ソニーの3gp音楽ファイルをm4aに変換するスクリプト

### 概要
- ソニーの音楽アプリ（x-アプリ）で生成された3gp音楽ファイルを想定
  - 時期によってフォーマットの仕様が変わっているので変換できない物もあるかも
- 以下の形式のファイルに対して変換処理を行う
  - 拡張子: `.3gp`
  - FileType: `3GP` or `MP4`
  - AudioFormat: `mp4a`
- 音声データはそのまま（ストリームコピー）
- 引き継げるメタデータ
  - `Title` `Artist` `Album` `AlbumArtist(※)` `Year` `Track(※)` `DiscNumber(※)` `Genre` `CoverArt(※)`
  - `AlbumArtist` `Track` `DiscNumber`はファイルパスから推測するため完全ではない
    - `\/:*?"<>|.`などのパスに使用できない記号が`_`に変換されているため
      - `Feat_`→`Feat.`のみ自動で修正
    - タグから`AlbumArtist`を取得できる時はそちらを優先(FileTypeがMP4の場合など)
  - `CoverArt`はファイルパスに[特定の文字](https://www.google.com/search?q=0x5c%E5%95%8F%E9%A1%8C)が含まれる場合読み出せない
  - 基本的に読み出せるのはID3v2形式で記録されているものだけ
  - 変換時ソニーの独自形式部分に記録された読み仮名やカバーアートは破棄される
- タイムスタンプは変換前のファイルから引き継ぐ



### 動作確認環境
- Windows10 64bit
- PowerShell 5.1 (Win10に付属)
- exiftool 12.4.4.0
- MediaInfo 22.6.0.0
- ffmpeg 2022-09-12-git-3ce6fa6b6d-essentials_build



### 準備
#### 必要なファイル
- [run.bat, script.ps1](https://github.com/Twilyze/3gp-to-m4a/releases)
  - script.ps1を実行するバッチファイルとPowerShellスクリプト
  - `Source code (zip)`
- [exiftool.exe](https://exiftool.org/)
  - 3gpファイルのメタデータ抽出に使用
  - `Windows Executable: exiftool-nn.nn.zip (n.n MB)`
    - `exiftool(-k).exe`を`exiftool.exe`にリネーム
- [MediaInfo.exe](https://mediaarea.net/en/MediaInfo/Download/Windows)
  - カバーアートのファイルタイプ判別に使用
  - `64 bit - CLI`
- [ffmpeg.exe](https://ffmpeg.org/download.html#build-windows)
  - 3gp→m4aの変換に使用
  - `Windows builds from gyan.dev`
    - `ffmpeg-git-essentials.7z`


#### 注意
- Windowsセキュリティのコントロールされたフォルダーアクセスがオンの場合、`ffmpeg.exe`を許可しないと動きません
- 変換時に同じ名前のm4aファイルは上書きされます（デフォルト）
- 変換後の選択肢で[はい]を選ぶと変換前の3gpファイルをゴミ箱へ送りますが、もしゴミ箱へ送れない場所にファイルがあった場合そのまま削除されます
- 3gpファイルの情報を全て引き継げるわけではないので、消えたら困る場合はバックアップをしてください



### 使用方法
#### 3gpファイルが保存されているフォルダを開く
- デフォルトの場所  
`C:\Users\Public\Music\Sony MediaPlayerX\Shared\Music`


#### 必要なファイルをフォルダに配置
- run.bat
- script.ps1
- exiftool.exe
- MediaInfo.exe
- ffmpeg.exe
```
Sony MediaPlayerX\Shared\Music
│  run.bat
│  script.ps1
│  exiftool.exe
│  MediaInfo.exe
│  ffmpeg.exe
├─アーティスト
│  ├─アルバム
│  └─アルバム
├─アーティスト
│  └─アルバム
...
```
※exiftool, MediaInfo, ffmpegはパスを通したフォルダに配置してもOK


#### run.batをダブルクリックして実行
- 出力されるファイル
  - `_all-3gp.tsv`, `_meta.tsv`, `_3gp.m3u8`, `_converted-m4a.m3u8`
  - 各アルバムフォルダ内
    - `cover.jpg` or `cover.png` or `cover.bmp`
    - `.3gp`ファイルを変換した`.m4a`ファイル
  - カバーアートが設定できなかったアルバムは`_no-picture.m3u8`, `_no-picture-album.tsv`に記録されます
  - 予期しない形式の3gpファイルは`_unexpected.m3u8`, `_unexpected.tsv`に記録されます
- 削除されるファイル
  - 変換前の`.3gp`ファイル
  - `_all-3gp.tsv`, `_meta.tsv`, `_3gp.m3u8`
  - ※`_unexpected.tsv`に記録されたファイルは削除されません
- ファイル削除前にもう一度実行すると各処理をスキップするか選択肢が出ます


#### 出力されるプレイリストファイルについて
`.m3u8`は音楽ソフトなどで読み込めるプレイリストファイル。  
`.tsv`は追加の情報を少し含めたリストで表計算ソフトなどで読み込めます。  
(※相対パスなので別のフォルダに移すと読み込めなくなる)

- `_converted-m4a.m3u8`
  - 変換したm4aファイル全てのプレイリスト
- `_no-picture.m3u8`
  - カバーアートが設定できなかったm4aファイル全てのプレイリスト
- `_no-picture-album.tsv`
  - カバーアートが設定できなかったアルバム一覧
- `_unexpected.m3u8`, `_unexpected.tsv`
  - (FileType: `3GP` or `MP4`) と (AudioFormat: `mp4a`) に当てはまらなかったファイル一覧


#### 使用したファイルを削除
- run.bat
- script.ps1
- exiftool.exe
- MediaInfo.exe
- ffmpeg.exe
- _converted-m4a.m3u8
- _no-picture.m3u8
- _no-picture.tsv
- _unexpected.m3u8
- _unexpected.tsv



### ライセンス
[cc0 1.0](/LICENSE)
