# もくじ
1. プログラムの概要
2. 動作環境
3. 操作方法
4. 初期設定（linux OSの場合）
5. 動きの説明
6. 今後の展望

# プログラムの概要
Google driveから電子の売上データを取得し、  
指定された期間内のセールアイテムに推奨されるアイテム一覧を出力します  

# 動作環境
1. linux OSまたはGoogle Colaboratory
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
出力されたスプレッドシートはhttps://drive.google.com/drive/folders/1kq3716S0RYLroAYROBLVGWaUnHTPdMd3
に格納されています
  
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
表示されたURLをクリック（gradio.liveで終わる方のURLをクリック）  
8. セール必要情報の記入
ジャンル（単一選択）、ストア（複数選択）、開始日付、終了日付、アイテム点数を入力
「セールアイテムを出力」ボタンをクリック
9. 出力結果の確認
ボタンの下にあるデータフレーム部分にアイテムと過去30日間売上金額が降順で表示されます
10. 出力結果のスプレッドシート化
フォーマット（単一選択）を選択し、「スプレッドシートに転記」ボタンをクリック
10秒ほどで「出力URL」欄にスプレッドシートのURLが表示されます
出力されたスプレッドシートはhttps://drive.google.com/drive/folders/1kq3716S0RYLroAYROBLVGWaUnHTPdMd3
に格納されています
  
# 初期設定（linux OSの場合）
1. 仮想環境の構築
`python -m venv .venv`
2. 仮想環境の起動
`source .venv/bin/activate`
3. 必要ライブラリのインストール
`pip install -r requirements.txt`
  
# 動きの説明
1. コマンドラインの認識
最初のアーギュメントでウェブページを公開状態にするかどうかを認識します  
--noshareで非公開状態、--shareで公開状態になります  
linux OSなら--noshare、Google Colaboratoryなら--shareで実行します  
次のアーギュメントでブラウザ内部で実行しているのか外部で実行しているのかを認識します  
--outbrowserで外部、--inbrowserで内部で実行する状態です  
linux OSなら--outbrowser、Google Colaboraoryなら--inbrowserで実行します  
3番目のアーギュメントでwebページにauth認証を付けます  
--adminで実行します
4番目のアーギュメントでローカルで実行しているのかグローバルで実行しているのかを認識します
linux OSなら--local、Google Colaboratoryなら--globalで実行します
2. 売上データの取得
Google Drive上に格納された売上データを金額ベース、冊数ベースでそれぞれ取得します  
4番目のアーギュメントが--localならGoogleサービスアカウントで、--globalならマウントされたドライブ上からglobで取得します
3. 売上データのジョイン・DF化
金額ベースの売上、冊数ベースの売上、Googleサービスアカウントでドライブ上から取得したコンテンツID変換テーブルをジョインします  
最終的に商品コード、書店グループ名、割引があったかどうか、実売日、売上金額（上代）、セール時の売上金額（予測）で構成されたデータフレームになります
セール時の売上金額はセール期間の日売れが1.2倍になった場合（仮）xセール日数で計算しています
5. セール対象アイテムの選定
・Kindleの場合  
　過去30日間に割引がなかったアイテムを選定します  
・そのほかのストアの場合  
　過去8週間のうち、開始日までと終了日まででそれぞれ4週間以上連続してセールがなかったアイテムを選定します
ストアが複数ある場合、選定されたアイテムの中で重複するもの＝すべてのストアで選定条件を満たすものを選定します
6. アイテムの売上取得
セール許可アイテムの売上を取得し、上位20%,中位60%,下位20%に分類します  
その後上位が20%,中位が60%,下位が20%の割合になるように入力された点数のセールアイテムを選出します  
その後、ゴール金額より高く、ゴール金額の倍より少なくなるようにセールアイテムを調整します  
調整に時間がかかることを避けるため、調整は最大20回に設定されています  
調整を行った後、選定アイテムに過去30日間売上金額、セール時の予測売上金額、メディアドゥID、MBJID、ヨドバシIDを付与してデータフレームとして出力されます  
8. 選定アイテムのスプレッドシート転記
選定アイテムについて、変換テーブルとして取得した商品マスタに基づいて必要情報を取得しgspreadで作成した新規シートに転記します
この時、フォーマットを電子取次ごとに選ぶことで、選ばれた電子取次のコンテンツIDを転記します
9. スプレッドシートのURL表示
転記が完了したスプレッドシートのURLを表示します

# 今後の展望
1. 選定アルゴリズムの見直し
現在の出力結果をそのまま採用すると売上金額の高いアイテムばかりが先んじてセールに投入されてしまう＝将来的にセールアイテム全体の売上が下がってしまうので、
目標売上金額をもとにバランスの取れたセール選定アルゴリズムに変更します  
2. 振り返り機能の追加
セール・ストアごとに各商品の売上がどれくらい上がったかの効果検証機能を付与します
