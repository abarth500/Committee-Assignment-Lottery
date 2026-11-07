function myFunction() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '初回デプロイのコードが反映されています！',
    'デプロイ確認',
    10
  );
}