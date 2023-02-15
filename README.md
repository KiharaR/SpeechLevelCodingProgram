# スピーチレベルコーディングプログラム
BTSJのコーディングルールに基づき，対話ログのスピーチレベルやスピーチレベルシフトの判定を行うプログラムです．

## プログラム概要
本プログラムは，BTSJのコーディングルールに基づき，対話ログに対してスピーチレベルやスピーチレベルシフトのコーディングを行うプログラムです．  
なお，本プログラムはWindows上で開発したため，以降Windowsを前提とした環境構築方法について記載しています．

## システム要件
* Anaconda（Python3.7以降）
* MeCab
* Cコンパイラ

## インストール
システム要件を満たしていない場合，以下の方法でインストールをしてください．

### Anaconda
本プログラムはPythonファイルであるため，Pythonの実行環境が必要です．  
以下のリンクからAnacondaをインストールしてください．  
* ダウンロード：[https://www.anaconda.com/products/distribution](https://www.anaconda.com/products/distribution)
* インストール方法（Windows版）：[https://www.python.jp/install/anaconda/windows/install.html](https://www.python.jp/install/anaconda/windows/install.html)

### MeCab
本プログラムでは，発話文のスピーチレベルを判定するためにMeCabを使用します．  
以下のリンクからMeCabをインストールしてください．  
* ダウンロード：[https://github.com/ikegami-yukino/mecab/releases](https://github.com/ikegami-yukino/mecab/releases)
* インストール方法（Windows10）：[https://self-development.info/mecabをインストールしてpythonで使う【windows】/](https://self-development.info/mecab%E3%82%92%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB%E3%81%97%E3%81%A6python%E3%81%A7%E4%BD%BF%E3%81%86%E3%80%90windows%E3%80%91/)  
  パス設定は，以下のサイトを参考にしてください．
  * パス設定：[http://realize.jounin.jp/path.html](http://realize.jounin.jp/path.html)
* インストール方法（Windows11）：[https://self-development.info/windows-11へのmecabのインストール/](https://self-development.info/windows-11%E3%81%B8%E3%81%AEmecab%E3%81%AE%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB/)

### Cコンパイラ
プログラムの実行に必要なコンパイラです．  
* ダウンロード：[https://visualstudio.microsoft.com/ja/visual-cpp-build-tools/](https://visualstudio.microsoft.com/ja/visual-cpp-build-tools/)
  - インストーラー実行後の選択画面では「C++によるデスクトップ開発」をクリックし，インストールしてください．

上記3つをインストールしたら環境構築は完了です．

## 実行方法
※以下，環境構築済みであることを前提とする．

1. 以下のいずれかの方法でソースコードをダウンロード
  * cloneする場合：[https://codelikes.com/git-clone/](https://codelikes.com/git-clone/)
    * git cloneの実行には[git](https://git-scm.com/)が必要です．  
    * Anaconda Prompt上で実行する場合，gitがインストールされていなければ
      ```sh
      conda install git
      ```
      と入力し，Gitをインストールしてから行ってください．
  * 「Download ZIP」からダウンロードする場合
    * 「Download ZIP」を選択することでzipファイルがダウンロードされるため，PC側で解凍してください．
2. 「スタート」メニューから「Anaconda Prompt」と入力し，Anaconda Promptを起動
3. Anaconda Prompt上で以下のコマンドを入力し，Enter
    ```sh
    cd BTSJexcelLog.pyのパス
    ```
    ※パスはBTSJexcelLog.pyを右クリックし，プロパティを開くことで確認できる．
    ※BTSJexcelLog.pyの保存先ドライブとAnaconda Promptに表示されているドライブが異なる場合は，まず以下のコマンドを入力しドライブ移動する．
      ```sh
      BTSJexcelLog.pyが保存されているドライブ:
      ```
4. Anaconda Prompt上で以下のコマンドを入力し，Enterを押すとプログラムが実行
    ```sh
    python BTSJexcelLog.py
    ```
5. プログラムの実行が終了後，「dialogueLog」フォルダ内の各フォルダにコーディング後の出力結果（ファイル名に「-label」が含まれるファイル）が保存される．また，Anaconda Prompt上には各フォルダのスピーチレベルの合計や平均の丁寧体率（丁寧体の使用割合）が表示される．

__注：プログラムを再度実行する場合，実行前に必ず「-label」を含むファイルを削除する必要がある（削除せずに実行すると，コーディング済みのファイルも再度コーディングされてしまうため）__

## 本プログラムにおけるコーディング基準
### スピーチレベルのラベル
本プログラムでは，スピーチレベルのラベルを以下のように定義する．

| ラベル | 定義                                                                    |
| :---: | ----------------------------------------------------------------------- |
| P     | 挨拶などの定型表現，「です。/ます。」，「ください。」                        |
| P*    | 「です。/ます。」+α（※1），「ください。」+α（※2）                         |
| P**   | 「です、/ます、」+α（※1）                                                |
| NM    | あいづち・応答，形容動詞終了，名詞止め，言いよどみ，名詞＋助詞，中途終了型発話 |
| N     | 上記以外                                                                 |

※1 αには「ね/よ/か/けど/から/って」が入る  
※2 αには「ね/よ」が入る  

なお，Pにおける挨拶などの定型表現は宇佐美[1][2]のスピーチレベルのコーディングに基づく．ただし，「はじめまして（初めまして）」と「こんばんは」は独自に定義したものである．

| P（定型表現）                                                                 |
| ---------------------------------------------------------------------------- |
| よろしくお願いいたします　よろしくお願いします　申し訳ございません　申し訳ありません　おはようございます　こんにちは　ありがとうございます　ごめんなさい　すみません　はじめまして（初めまして）　こんばんは |

また，NMにおけるあいづち・応答については，プログラムの使用上すべてのあいづち・応答をコーディングすることができないため，本プログラムでは，宇佐美[1][2]のスピーチレベルのコーディングと吉田ら[3]の相槌表現認定基準に当てはまる発話とする．

| NM（あいづち・応答）                                                                 |
| ---------------------------------------------------------------------------- |
| うん　そう　はい　あ　え　うーん　ね　どうぞ　まあね　よろ　おう　ありあり　どうも　いえいえ　もちろん　あー　えー　いやいや　そだねー　そうだねぇ　そうだね　そうだな　せやな　はーい　ぜひ　まあ　いえ　はいよ　うい　ううん　そだね　うむ　そうね　うんうん　そうだねー　　そうだなー　いいえ　だなー　だね　いや　なるほど　わかる　わかった　わかったよ　いいよ　いいよー　いいよね　ええんやで　いいだろう　なんだろう　とんでもない　なる　そんなことないよ　とくには　いいけど　まあーまあー　もちろん　それほどでも　あまり　また　特に　うふふ　たまに　そうかも　いますぐ　|

※ 上記は，我々の実験にて取得した対話ログで見受けられたものであるため，すべてのあいづち・応答は網羅できていない

### スピーチレベルシフト
本プログラムでは，宇佐美[2]の定義に基づき，スピーチレベルシフトを以下のように定義する．

| シフト        |　定義                  |
| :----------: | ---------------------- |
| ダウンシフト  | P/P*/P**からNへのシフト |
| アップシフト  | NからP/P*/P**へのシフト |

## コーディング方法
### スピーチレベルのコーディング
発話文のスピーチレベルのコーディングは，1文ごとに行う（1発話ターンに複数文含まれている場合でも，1文ずつコーディングする）．なお，以下に示す記号が含まれている場合に1文としてコーディングするため，適切な1文をコーディングできない可能性がある．

| 記号                    |
| ----------------------- |
| 。　！　？　☺　（笑）　笑 |

※ 上記は我々の実験にて取得したアイワログにて見受けられたものを基準としている

### スピーチレベルシフトのコーディング
※ コーディング例はマニュアルファイルを参照

スピーチレベルシフトのコーディングは，対象の発話文とその前ターンの発話文のスピーチレベルを比較することによりコーディングする.

ただし，例外として以下のような場合がある．
1. 発話文のスピーチレベルコーディングに P/P*/P**と N が含まれる場合は，必ずダウンシフトかアップシフトとコーディング
2. 1 発話のスピーチレベルのコーディング結果が複数かつ NM を含む場合は，NM を考慮せずスピーチレベルシフトをコーディング
3. 前ターンの発話文においてスピーチレベルのコーディング結果が NM の場合は，その前ターンのコーディング結果を参照してコーディング

## プログラムの注意点
作成したプログラムは，我々の実験で使用することを前提として開発したため，以下の点に注意が必要である.
* コーディングに使用する対話ログは,xlsx ファイルのみに対応
* 現行バージョンでは,各 15 ターンまでのコーディングにのみ対応
* 現行バージョンでは,対話ログに含まれる「ユーザ年齢」の入力がなければプログラム実行時にエラーが出る可能性あり
* 対話ログファイルに含まれる「system」の「発話内容」におけるスピーチレベルのコーディングは,BTSJ のコーディングルールに基づくものではなく,「rule」フォルダ内のtxt ファイルとの比較によりコーディング
  * 「rule」フォルダ内の txt ファイルと「system」の「発話内容」を書き換えることで,スピーチレベルのコーディングをすることが可能(ただし,コーディングは「丁寧体,混在(丁寧体と非丁寧体),非丁寧体」の 3 つ)
* スピーチレベルのコーディングは,BTSJ のコーディングルールが基本となるが,我々の実験にて取得した対話ログにおける発話文が正しくコーディングできないものがあったため,一部独自に定義している
* プログラム実行後に出力されるものは,それぞれ以下の意味を示す
  * 〇〇条件〇カウント  
    各条件のスピーチレベルの合計を出力している.  
    なお,条件とはシステム発話のスピーチレベルのことを指している.以下に具体的に示す.  
    * 丁寧体条件(teineitai):常に丁寧体
    * 文末制御条件(switching):丁寧体と非丁寧体を使い分ける
    * 非丁寧体条件(hiteineitai):常に非丁寧体  
    上記の()に示すとおり,各条件は「dialogueLog」フォルダ内の各フォルダと対応している.  
  * 〇〇条件丁寧体率  
    各条件における user の丁寧体率の平均を示している.なお,丁寧体率は以下のように算出している.  
    ```math
    丁寧体率 = \frac{userのPの合計 + userのP ∗ の合計 + userのP ∗∗ の合計}{userのスピーチレベルの合計}
    ```

## プログラムの改良方法
開発したプログラムは,BTSJexcelLog.py において以下に示す行のコードを書き換えることで,様々な対話ログのコーディングを行うことができる.なお,プログラムを書き換える際には,[VScode](https://code.visualstudio.com/download) を使用することを推奨する.  
1. スピーチレベルにおけるコーディング基準の変更
  114-132 行目は,定型表現や NM におけるあいづち・応答のコーディングリストである.以下にそれぞれのリストの説明を示す.
    * teikeiList1:P における定型表現のリスト(「。」で終わる定型表現)
    * teikeiList2:P における定型表現のリスト(「!/?」で終わる定型表現)
    * NMlist:NM におけるあいづち・応答のリスト
    コーディングリストの中身を変更したい場合,[ ]内を書き換えることで変更することができる.ただし,以下の 2 点に注意する必要がある.
    * teikeiList1 は「。」で終わる表現,teikeiList2 は「!/?」で終わる表現として分類しているため,適切なリストを書き換えなければスピーチレベルのコーディングに誤りが生じる可能性がある
    * teikeiList2 と NMlist において新たに文章を追加する場合,記号(「。/!/?」)の数だけ文章を追加する必要がある.
2. dialogueLog に含まれている対話ログ以外の対話ログ(独自に用意した対話ログなど)のコーディング
  配布時に dialogueLog に含まれている対話ログ以外の対話ログもコーディング可能である.ただし,以下の要件を満たしていない場合はコーディングすることができない.
    * template.xlsx の「発話内容」(C2-C33 セル)※と「ユーザ年齢」(I1 セル)に入力後,「名前を付けて保存」から適切なファイル名で保存されていること
    ※「話者」が「system」となっている「発話内容」は書き換えても正しくコーディングされないため,入力済みのものからの書き換えは不要である
    * 上記ファイルを「dialogueLog」フォルダ内のいずれかのフォルダに格納されていること

## 参考文献
[1] 宇佐美まゆみ:スピーチレベルのコーディング_ 修正版_110128,2011  
[2] 宇佐美まゆみ:【最新版】スピーチレベルとシフトのコーディングのルール 宇佐美 13  
[3] 吉田奈央,高梨克也,伝康晴:対話におけるあいづち表現の認定とその問題点について,言語処理学会第 15 回年次大会発表論文集,2009
