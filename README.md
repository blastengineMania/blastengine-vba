# blastengine SDK for VBA

ExcelなどのVBAで、blastengineを利用するためのSDKです。

## 使い方

リリースページから最新のバージョンをダウンロードするか、コードを vbac でコンパイルしてください。

## 参照設定の追加

VBAエディタで、ツール > 参照設定 から、以下の参照設定を追加してください。

- Microsoft ActiveX Data Objects 6.1 Library
- Microsoft XML, v6.0

※ 将来的には、参照設定を不要にする予定です。

## 初期化

SDKを初期化する際には、blastengineのユーザーIDとAPIキーを指定してください。

```vb
Dim client As new Blastengine
client.UserId = "YOUR_USER_ID"
client.ApiKey = "YOUR_API_KEY"
```

## 即時配信メールの送信

即時配信メール（Transaction Mail）を送信するには、以下のようにします。

```vb
Dim transaction As Transaction
Set transaction = client.Transaction

transaction.Email = "user@example.jp"
transaction.From "info@example.com"
transaction.Subject = "テストメール from Excel"
transaction.TextPart = "メールの本文です __name__"

' 置き換え文字列
Dim insertCode As Dictionary
Set insertCode = New Dictionary
insertCode.Add "name", "Test"
transaction.InsertCode = insertCode
If transaction.Send Then
		Debug.Print transaction.DeliveryId
Else
		Debug.Print JsonConverter.ConvertToJson(transaction.Error)
End If
```

### 添付ファイル付きメール

添付ファイルは、ファイルをパスで指定してください。

```vb
transaction.Attachments(0) = "C:\path\to\test.png"
transaction.Attachments(1) = "C:\path\to\test.pdf"
```

## ライセンス

MIT


