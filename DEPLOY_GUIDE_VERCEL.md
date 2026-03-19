# せどりポイポイ Web公開手順書 (Vercel編)

このガイドでは、あなたのパソコンにある「せどりポイポイ」を、インターネット上で使えるようにする（デプロイする）手順を説明します。
中学生でもできるように、ゆっくり丁寧に説明します。

---

## 準備するもの
1.  **GitHub（ギットハブ）のアカウント**
    - コードを保存する場所です。持っていない場合は [GitHub公式サイト](https://github.com/) で無料登録してください。
2.  **Vercel（バーセル）のアカウント**
    - アプリを動かす場所です。持っていない場合は [Vercel公式サイト](https://vercel.com/) で「Sign Up」し、「Continue with GitHub」を選んで登録してください。

---

## 手順1: VS CodeからGitHubにアップロードする

あなたのパソコンにあるコードを、まずはGitHubにアップロードします。

1.  **VS Code** で「せどりポイポイ」のプロジェクトを開きます。
2.  左側のメニューにある **「ソース管理」（枝分かれした線のアイコン）** をクリックします。
3.  **「GitHub に公開 (Publish to GitHub)」** という青いボタンがあればクリックします。
    - ※もしボタンがない場合は、メッセージ欄に「first commit」と入力して「コミット」ボタンを押し、その後に「ブランチを公開 (Publish Branch)」ボタンを押してください。
4.  上の方に検索バーが出るので、**「Publish to GitHub public repository（公開リポジトリとして公開）」** または **「private repository（非公開リポジトリ）」** を選びます。
    - 自分だけで使うなら **Private (非公開)** がおすすめです。
5.  数秒待つと、「GitHub に正常に公開されました」と右下に出ます。「GitHub で開く」を押して、コードがアップロードされたか確認しましょう。

---

## 手順2: Vercelでアプリを作成する

次に、GitHubにあるコードをVercelに取り込んで動かします。

1.  [Vercelのダッシュボード](https://vercel.com/dashboard) にアクセスします。
2.  右上の **「Add New...」** ボタンを押し、**「Project」** を選びます。
3.  **「Import Git Repository」** の画面になります。
    - 先ほどアップロードした「sedori_poipoi」が表示されているはずです。
    - 横にある **「Import」** ボタンを押します。

---

## 手順3: 環境変数の設定 (一番重要！)

アプリが動くための「鍵（APIキー）」を設定します。これを忘れると動きません。

1.  **「Configure Project」** という画面になります。
2.  **「Environment Variables」** という項目をクリックして開きます。
3.  あなたのパソコンの `.env.local` ファイルの中身を1つずつ入力します。
    - `.env.local` をVS Codeで開いて確認してください。

    **入力例:**
    
    | Key (名前) | Value (値) |
    | :--- | :--- |
    | `NEXT_PUBLIC_KEEPA_API_KEY` | (あなたのKeepaキー) |
    | `RAKUTEN_APP_ID` | (あなたの楽天ID) |
    | `YAHOO_CLIENT_ID` | (あなたのYahoo ID) |

    ※ `.env.local` にあるものは基本的に全て追加してください。
    ※ `NEXT_PUBLIC_...` で始まるものも重要です。

4.  全て入力し終わったら、下の **「Deploy」** ボタンを押します。

---

## 手順4: 完成！

画面が切り替わり、紙吹雪が舞ったら成功です！🎊
**「Domain (ドメイン)」** というところに、あなたのアプリのURL（例: `sedori-poipoi.vercel.app`）が表示されます。

このURLをクリックして、アプリが開くか確認してください。
もしエラーが出る場合は、環境変数の設定が間違っている可能性が高いです（Settings > Environment Variables から修正できます）。

---

## 最後に: AppSheetへの設定

このURLを使って、AppSheet連携ツールを設定します。
`developer/sedori_appsheet_tool/Code.gs` を開き、行頭の `API_URL` を修正してください。

```javascript
// 変更前
const API_URL = "https://[あなたのアプリのURL]/api/appsheet/scan";

// 変更後 (例)
const API_URL = "https://sedori-poipoi-xxxxx.vercel.app/api/appsheet/scan";
```

これで完了です！お疲れ様でした！
