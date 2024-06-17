VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blastengine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_ApiKey As String
Private m_UserId As String
Private m_Token As String
Private m_EndPoint As String

Private Sub Class_Initialize()
    m_ApiKey = ""
    m_UserId = ""
    m_EndPoint = "https://app.engn.jp"
End Sub

Property Let ApiKey(ByVal ApiKey As String)
    m_ApiKey = ApiKey
End Property

Property Let UserId(ByVal UserId As String)
    m_UserId = UserId
End Property

Function Transaction() As Transaction
    Set Transaction = New Transaction
    Transaction.client = Me
End Function

Property Get Token() As String
    Token = EncodeToBase64(LCase(SHA256(m_UserId & m_ApiKey)))
End Property

Property Get EndPoint() As String
    EndPoint = m_EndPoint
End Property



Private Function SHA256(s As String) As String
    Dim objSHA256
    Dim objUTF8

    Dim bytes() As Byte
    Dim hash() As Byte

    Dim i
    Dim wk

    '// INIT
    Set objSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    Set objUTF8 = CreateObject("System.Text.UTF8Encoding")

    '// 文字列を UTF8 にエンコードし、バイト配列に変換
    bytes = objUTF8.GetBytes_4(s)

    '// ハッシュ値を計算（バイナリ）
    hash = objSHA256.ComputeHash_2((bytes))

    '// バイナリを16進数文字列に変換
    For i = 1 To UBound(hash) + 1
        wk = wk & Right("0" & Hex(AscB(MidB(hash, i, 1))), 2)
    Next i

    '// 結果を返す
    SHA256 = LCase(wk)

End Function

' https://ribbit.konomi.app/vba/base64-encode/
'*-------------------------------------------------------------
'* テキストをBase64でエンコード
'*
'* @param text 変換する値
'* @return Base64フォーマットデータ
'*-------------------------------------------------------------
Public Function EncodeToBase64(ByRef text As String) As String

  ' オブジェクトの準備
  Dim node As Object
  Set node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")

  ' エンコード
  node.DataType = "bin.base64"
  node.nodeTypedValue = ConvertToBinary(text)

  ' 関数で取り除けない改行を削除して返却
  EncodeToBase64 = Replace(node.text, vbLf, "")
End Function

'*-------------------------------------------------------------
'* 文字列をバイナリデータに変換
'*
'* @param text 変換する値
'* @return バイナリデータ
'*-------------------------------------------------------------
Private Function ConvertToBinary(ByRef text As String)

  ' オブジェクトの準備
  Dim BinaryStream As Object
  Set BinaryStream = CreateObject("ADODB.Stream")

  ' Streamの設定
  With BinaryStream
    .Type = 2
    .Charset = "us-ascii"
    .Open
    .WriteText text
    .Position = 0
    .Type = 1
    .Position = 0
  End With

  ConvertToBinary = BinaryStream.Read
End Function
