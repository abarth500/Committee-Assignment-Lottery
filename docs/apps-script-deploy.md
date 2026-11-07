# Apps Script の自動反映

同じ `apps-script/` のコードを、GitHub のイベントに応じて次のシートへ反映します。

| タイミング | 反映先 | 反映するコード |
| --- | --- | --- |
| `main` に push | [テスト用シート](https://docs.google.com/spreadsheets/d/1FySsnrRI4uZb1SCxDsng3LAQQZvBHw1tB8MaKTorfkE/edit) | その push のコミット |
| GitHub Release を公開 | [本番用シート](https://docs.google.com/spreadsheets/d/1sDSV2-9KGVr04cRFovnKopbwBwM9mJtFHU108_2bxJE/edit) | 公開したリリースのタグが指すコミット |

スクリプト ID はワークフロー内に固定しています。

| 環境 | スクリプト ID |
| --- | --- |
| テスト | `1ba_yMwjRt-z2PHcyCRrNkYtPnP5TSYxFJH993ywaL-dUFKp2l-e5Samw` |
| 本番 | `1EfpkaHJ7BmEaIIKBZl72HzdjtoSJmVssOxd86-x-q6RFWjWrzr2i-GjG` |

通常の開発・デバッグ・抽選の試行はテスト用シートで行います。`main` の更新は本番のスクリプトを更新しません。コードは対象シートに紐づく `SpreadsheetApp.getActiveSpreadsheet()` を使う方針とし、本番シートの ID をコードに固定しないでください。

## 初回設定

1. 両方のスクリプトとシートに編集権限を持つ Google アカウントで https://script.google.com/home/usersettings の **Google Apps Script API** を有効にします。
2. 手元の PC で Node.js 24 を用意し、次を実行して Google の認証画面で許可します。

   ```sh
   npx --yes @google/clasp@3.4.1 login
   ```

3. リポジトリのルートで、テスト用プロジェクトの初期ファイルとマニフェストを取り込みます。`apps-script/` に既存ファイルがある場合は上書きせず、差分を確認してください。

   ```sh
   npx --yes @google/clasp@3.4.1 clone-script 1ba_yMwjRt-z2PHcyCRrNkYtPnP5TSYxFJH993ywaL-dUFKp2l-e5Samw --rootDir apps-script
   npx --yes @google/clasp@3.4.1 show-file-status
   ```

   以後は `apps-script/` 内でプログラムを作成します。初期のサンプル関数を同期するだけでは目的の機能は実装されません。TypeScript を直接送信する設定ではありません。

4. GitHub の **Settings → Secrets and variables → Actions** で次を登録します。

   | 種類 | 名前 | 値 |
   | --- | --- | --- |
   | Repository secret | `CLASP_AUTH_JSON` | ログインで生成された `~/.clasprc.json` の内容全体 |

   同じ Secret をテスト・本番の両方で使用します。Repository variable `APPS_SCRIPT_ID` による反映先の上書きは使用しません。

   認証ファイルはチャット、Git、ログに貼り付けないでください。`clasp login` は既定ユーザーで実行します。OAuth 権限が失効した場合は再認証して Secret を更新します。

   手元の PC で GitHub CLI にログイン済みの場合は、ファイルを画面に表示せず登録できます。

   ```sh
   gh secret set CLASP_AUTH_JSON --repo abarth500/Committee-Assignment-Lottery < ~/.clasprc.json
   ```

   このクラウドからの GitHub API の Secret 一覧取得は `Forbidden` だったため、既存 Secret の有無は未確認です。登録済みで使用可能なら、その設定を再利用してください。

5. コードと設定を `main` に反映し、Actions の **Sync Apps Script** が成功したことを確認して、テスト用シートで動作を確認します。

## リリースの手順

1. テスト用シートで確認したコミットに、例として `v1.0.0` のタグを付けます。
2. GitHub の **Releases → Draft a new release** でそのタグを選び、**Publish release** を実行します。
3. Actions の **Sync to production** が成功したことと、本番用シートの動作を確認します。

タグの push だけ、リリースの下書き保存、公開後の説明文の編集では本番へ反映しません。`release.published` を使用するため、プレリリースも公開すると本番への反映対象です。公開するタグは必ずテスト済みのものを選んでください。

リリース後に `main` が進んでも、本番へ送るのはリリースのタグの内容です。テスト・本番は別々の実行グループで同期します。手動実行による本番反映は設定していません。

## 同期の動作と現在の状態

これはシートに紐づくスクリプトのコード同期です。Web アプリなどのバージョン固定のデプロイ更新や、抽選関数の自動実行、トリガーの作成は行いません。

`clasp push --force` は対象のオンラインプロジェクト全体（マニフェストを含む）をローカルの内容で置き換えます。ローカルにないオンラインのファイルは削除対象になります。既存コードがあるプロジェクトへ反映する場合や、オンラインエディタで後から変更した場合は、すべて Git に取り込んでから同期します。

ワークフローは送信前に、イベントと宛先、マニフェスト、ソースの存在、JavaScript 構文、認証形式を確認します。これは Apps Script の実行テストではないため、Google 側での認可、メニュー、ボタン、トリガーなどの動作確認は別途必要です。

初回用の `apps-script/CommitteeLottery.js` とマニフェストは登録済みです。ローカルの構文・イベント分岐・送信対象ファイルの検証に加え、実際の同期結果は Actions の各実行結果で確認してください。成功後、対象シートを開いて `myFunction` を手動実行すると確認用の通知が表示されます。デプロイによって関数が自動実行されるわけではありません。
