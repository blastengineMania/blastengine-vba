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

Function Bulk() As Bulk
    Set Bulk = New Bulk
    Bulk.client = Me
End Function

Property Get Token() As String
    Token = EncodeToBase64(LCase(SHA256(m_UserId & m_ApiKey)))
End Property

Property Get EndPoint() As String
    EndPoint = m_EndPoint
End Property

Function SendText(method As WebMethod, strPath As String, Query As Dictionary, Body As Dictionary) As Dictionary
    Dim client As New WebClient
    client.BaseUrl = EndPoint
    Dim Request As New WebRequest
    Request.Resource = strPath
    Request.method = method
    Request.AddHeader "Authorization", "Bearer " & Token
    Request.AddHeader "Content-Type", "application/json"
    If (method = HttpPost Or httpPatch Or httpPut) And Body.Count > 0 Then
        Set Request.Body = Body
    End If
    WebHelpers.EnableLogging = False ' デバッグ用
    Dim Response As WebResponse
    Set Response = client.Execute(Request)
    Set SendText = JsonConverter.ParseJson(Response.Content)
    
    Set client = Nothing
    Set Request = Nothing
    Set Response = Nothing
End Function


Function SendFiles(httpMethod As WebMethod, strPath As String, Data As Dictionary, aryFiles() As String) As Dictionary
    ' Dim Data As Dictionary
    ' Set Data = BuildParams
    Dim Boundary As String
    Boundary = RandomString(24)
    Dim FilePath
    Dim stream As ADODB.stream
    Set stream = New ADODB.stream
    stream.Charset = "UTF-8"
    stream.Open
    stream.Position = 0
    stream.Type = adTypeText
    stream.WriteText "----" & Boundary & vbCrLf
    stream.WriteText "Content-Disposition: form-data; name=""data""; filename=""base.json""" & vbCrLf
    stream.WriteText "Content-Type: application/json" & vbCrLf
    stream.WriteText "" & vbCrLf
    stream.WriteText JsonConverter.ConvertToJson(Data)
    
    Dim p As Long
    If (Not aryFiles) <> -1 Then
        For Each FilePath In aryFiles
            AddAttachmentToStream stream, FilePath, Boundary
        Next FilePath
    End If
    
    ChangeStreamType stream, adTypeText
    stream.WriteText vbCrLf & "----" & Boundary & "--" & vbCrLf
    ChangeStreamType stream, adTypeBinary
    stream.Position = 0
    
    Dim objHTTP As XMLHTTP60
    Dim Params As Variant
    Dim strResult As String
    
    Set objHTTP = New XMLHTTP60
    Dim strMethod As String
    strMethod = IIf(httpMethod = HttpPost, "POST", "GET")
    objHTTP.Open strMethod, EndPoint & strPath, False
    objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=--" & Boundary
    objHTTP.setRequestHeader "Authorization", "Bearer " & Token
    Params = stream.Read
    objHTTP.Send Params
    Set stream = Nothing
    
    strResult = StrConv(objHTTP.responsebody, vbUnicode)
    If objHTTP.Status <> 201 Then
        Set SendFiles = JsonConverter.ParseJson(strResult)
    Else
        ' Dim Json As Dictionary
        ' Set Json = JsonConverter.ParseJson(strResult)
        ' m_DeliveryId = Json("delivery_id")
        ' SendFiles = True
        Set SendFiles = JsonConverter.ParseJson(strResult)
    End If
    Set objHTTP = Nothing
End Function

Function InsertCodeToArray(InsertCode As Dictionary) As Dictionary()
    Dim Key As Variant
    Dim i As Long
    i = 0
    Dim InsertCodes() As Dictionary
    ReDim InsertCodes(UBound(InsertCode.Keys)) As Dictionary
    For Each Key In InsertCode.Keys
        Dim Params As Dictionary
        Set Params = New Dictionary
        Params.Add "key", "__" & Key & "__"
        Params.Add "value", InsertCode.Item(Key)
        Set InsertCodes(i) = Params
        i = i + 1
    Next Key
    
    InsertCodeToArray = InsertCodes
End Function

Function ConvertToISO8601(dateValue As Date, Optional offset As Integer = 540) As String
    ' Declare variables
    Dim yearStr As String
    Dim monthStr As String
    Dim dayStr As String
    Dim hourStr As String
    Dim minuteStr As String
    Dim secondStr As String
    Dim timezoneStr As String
    ' Dim offset As Integer
    'If offset > 0 Then
    '    dateValue = DateAdd("n", offset * -1, dateValue)
    'End If
    
    ' Format date and time parts
    yearStr = Year(dateValue)
    monthStr = Right("0" & Month(dateValue), 2)
    dayStr = Right("0" & Day(dateValue), 2)
    hourStr = Right("0" & Hour(dateValue), 2)
    minuteStr = Right("0" & Minute(dateValue), 2)
    secondStr = Right("0" & Second(dateValue), 2)

    ' Get the timezone offset in minutes
    ' offset = Round((dateValue - Now) * 24 * 60)
    
    ' Format the timezone offset
    If offset = 0 Then
        timezoneStr = "Z"
    Else
        timezoneStr = IIf(offset > 0, "+", "-") & _
                      Right("0" & Abs(offset) \ 60, 2) & ":" & _
                      Right("0" & Abs(offset) Mod 60, 2)
    End If

    ' Concatenate parts into ISO8601 format
    ConvertToISO8601 = yearStr & "-" & monthStr & "-" & dayStr & "T" & _
                       hourStr & ":" & minuteStr & ":" & secondStr & timezoneStr
End Function

Private Sub AddAttachmentToStream(stream As ADODB.stream, FilePath As Variant, Boundary As String)
    Dim file As ADODB.stream
    Set file = New ADODB.stream
    file.Open
    file.Type = adTypeBinary
    file.LoadFromFile FilePath
    
    ChangeStreamType stream, adTypeText
    stream.WriteText vbCrLf & "----" & Boundary & vbCrLf
    stream.WriteText "Content-Disposition: form-data; name=""file""; filename=""" & Dir(FilePath) & """" & vbCrLf
    stream.WriteText "Content-Type: application/octet-stream" & vbCrLf
    stream.WriteText "" & vbCrLf
    
    ChangeStreamType stream, adTypeBinary
    stream.Write file.Read()
    
    file.Close
    Set file = Nothing
End Sub

Private Sub ChangeStreamType(ByRef stream As ADODB.stream, adType As Integer)
    Dim p As Long
    p = stream.Position
    stream.Position = 0
    stream.Type = adType
    If adType = adTypeText Then
        stream.Charset = "UTF-8"
    End If
    stream.Position = p
End Sub


Function RandomString(lngLengh As Long) As String
    Dim CharacterBank As Variant
    Dim x As Long
    Dim str As String

    CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
      "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
      "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
      "A", "B", "C", "D", "E", "F", "G", "H", _
      "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
      "W", "X", "Y", "Z")
      

    For x = 1 To lngLengh
      Randomize
      str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
    Next x

    RandomString = str
End Function


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
