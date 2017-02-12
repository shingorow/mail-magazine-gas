# Mail Magazine GAS
Google Spreadsheet のリストにメールを送る Google Apps Script

## 使用方法
Google Apps Script のスクリプトエディタでライブラリを読み込みます。API ID は下記のとおり。

API ID: API ID: MbJOSVUHvWQuWWvCLcZM4nbFSwZrFpnJA
サンプルコード

```javascript
function myFunction() {
  var config = {}
  
  config['spreadsheet'] = {
    ssId: 'スプレッドシートの ID',
    sheetName: 'リスト', // メール送信リストのあるシート名
    messageSheet: 'メッセージ', // 送信するメッセージのテンプレートが記載されているシートの名前
    testSheet: 'テスト',  // テストメッセージの送信先一覧が記載されているシート名
    displayValues: false, // スプレッドシートの値を取得する際、getDisplayValues() を使う場合は true にする
    emailLabel: 'メールアドレス' // リストのメールアドレスが記載されている項目の名前
  }
  
  config['email'] = {
    attachments: '',
    bcc: '',
    cc: '',
    from: '',
    inlineImages: '',
    name: '',
    replyTo: ''
  }
  
  
  MailMagazine.send(config);
  MailMagazine.test(config);
}
```
config は二つの子要素を設定できます。
一つはスプレッドシートに関する設定で、キーが "spreadsheet" というもの。
これは必須の要素です。ただし、ssId 以外の要素は上記のコードにある値がデフォルト値として設定されます。
ssId が無い状態でスクリプトを実行すると例外を投げます。

もう一つはメール送信に関するもので、キーが "email"。こちらの設定は必須ではありません。
設定していない要素については、GmailApp のデフォルト設定が適用されます。

## メソッド

### send()
send メソッドを実行するとアドレスリストに対してメールを送信します。
### test()

test メソッドを実行するとテスト用のリストにメールを送信します。
社内のメンバーなどを登録しておいて、差し込みがうまくいっているかなどを確認するのに使うと良いかと思います。

## お約束
このスクリプトを使用して起こった損害等について、当方では一切の責任を負いません。ご了承下さい。
