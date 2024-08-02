# もくじ
1. プログラムの概要
2. 動作環境
3. 操作方法
4. 初期設定（linux PSの場合）
5. 動きの説明
6. 今後の展望

# プログラムの概要
Google driveから電子の売上データを取得し、  
指定された期間内のセールアイテムに推奨されるアイテム一覧を出力します  

# 動作環境
1. linux OSまたはGoogle colaboratory
2. python 3.11
3. Google Chrome

# 操作方法
## linux OSの場合
1. 仮想環境の起動
`source .venv/bin/activate`
2. pythonコマンドの実行
`pyhon sale-picker_1_0.py --noshare --outbrowser --local`
3. ID,PASSの表示
シェルスクリプトに自動で表示されます  
`ID: discover`
`pass:ランダムな文字列`
5. URLの表示
シェルスクリプトに自動で表示されます
`Running on local URL:  http://127.0.0.1:7860`
6. URLへのアクセス
表示されたURLをブラウザにコピー&ペーストする
7. セール必要情報の記入
ジャンル（単一選択）、ストア（複数選択）、開始日付、終了日付、アイテム点数を入力
「セールアイテムを出力」ボタンをクリック
8. 出力結果の確認
ボタンの下にあるデータフレーム部分にアイテムと過去30日間売上金額が降順で表示されます
9. 出力結果のスプレッドシート化
フォーマット（単一選択）を選択し、「スプレッドシートに転記」ボタンをクリック
10秒ほどで「出力URL」欄にスプレッドシートのURLが表示されます
出力されたスプレッドシートはhttps://drive.google.com/drive/folders/1kq3716S0RYLroAYROBLVGWaUnHTPdMd3に格納されています
## Google Colaboratoryの場合
1. 必要ライブラリのインストール
`! pip install -r requirements.txt`
2. pythonコマンドの実行
`pyhon sale-picker_1_0.py --share --inbrowser --global`
3. D,PASSの表示
シェルスクリプトに自動で表示されます  
`ID: discover`
`pass:ランダムな文字列`
5. URLの表示
シェルスクリプトに自動で表示されます
`https://（ランダムな文字列）.gradio.live`
6. URLへのアクセス
表示されたURLをクリック  
7. セール必要情報の記入
ジャンル（単一選択）、ストア（複数選択）、開始日付、終了日付、アイテム点数を入力
「セールアイテムを出力」ボタンをクリック
8. 出力結果の確認
ボタンの下にあるデータフレーム部分にアイテムと過去30日間売上金額が降順で表示されます
9. 出力結果のスプレッドシート化
フォーマット（単一選択）を選択し、「スプレッドシートに転記」ボタンをクリック
10秒ほどで「出力URL」欄にスプレッドシートのURLが表示されます
出力されたスプレッドシートはhttps://drive.google.com/drive/folders/1kq3716S0RYLroAYROBLVGWaUnHTPdMd3に格納されています
