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
    Transaction.Client = Me
End Function

Function Bulk() As Bulk
    Set Bulk = New Bulk
    Bulk.Client = Me
End Function

Function Mail() As Mail
    Set Mail = New Mail
    Mail.Client = Me
End Function

Function MailQuery() As MailQuery
    Set MailQuery = New MailQuery
    MailQuery.Client = Me
End Function

Function Log() As Log
    Set Log = New Log
    Log.Client = Me
End Function

Function LogQuery() As LogQuery
    Set LogQuery = New LogQuery
    LogQuery.Client = Me
End Function

Property Get Token() As String
    Token = EncodeToBase64(LCase(SHA256(m_UserId & m_ApiKey)))
End Property

Property Get EndPoint() As String
    EndPoint = m_EndPoint
End Property

Function SendText(method As WebMethod, strPath As String, Query As Dictionary, Body As Dictionary) As Dictionary
    Dim Client As New WebClient
    Client.BaseUrl = EndPoint
    Dim Request As New WebRequest
    Request.method = method
    Dim QueryString() As String
    ReDim QueryString(UBound(Query.Items))
    
    Dim i As Integer
    For i = 0 To UBound(Query.Items)
        QueryString(i) = Query.Keys(i) & "=" & WorksheetFunction.EncodeURL(Query.Item(Query.Keys(i)))
    Next i
    
    strPath = Replace(strPath, "https://app.engn.jp/", "")
    Request.Resource = strPath & "?" & Join(QueryString, "&")

    Request.AddHeader "Authorization", "Bearer " & Token
    Request.AddHeader "Content-Type", "application/json"
    ' Debug.Print JsonConverter.ConvertToJson(Body)
    If (method = httpPost Or httpPatch Or httpPut) And Body.Count > 0 Then
        Set Request.Body = Body
    End If
    WebHelpers.EnableLogging = True ' デバッグ用
    Dim Response As WebResponse
    Set Response = Client.Execute(Request)
        
    Set SendText = JsonConverter.ParseJson(Response.Content)
    
    Set Client = Nothing
    Set Request = Nothing
    Set Response = Nothing
End Function

Function GetFile(strUrl As String) As String
    Dim objHTTP As XMLHTTP60
    Dim Params As Variant
    Dim strResult As String
    Set objHTTP = New XMLHTTP60
    Dim strMethod As String
    Dim zipFile As String
    zipFile = CreateTemporaryFile("zip")
    ' Debug.Print zipFile
    Dim Stream As ADODB.Stream
    Dim shellApp As Object
    Dim fileName As String
    Dim fileNum As Integer
    Dim fileFullPath As String

    strMethod = "GET"
    objHTTP.Open strMethod, strUrl, False
    objHTTP.setRequestHeader "Authorization", "Bearer " & Token
    objHTTP.Send
    If objHTTP.Status = 200 Then
        Set Stream = New ADODB.Stream
        Stream.Type = 1
        Stream.Open
        Stream.Write objHTTP.responseBody
        Stream.SaveToFile zipFile, 2
        Stream.Close
        Application.WAIT (Now + TimeValue("0:00:02"))
        extractFolderPath = CreateTemporaryFolder()
        Set shellApp = CreateObject("Shell.Application")
        shellApp.Namespace(extractFolderPath).CopyHere shellApp.Namespace(CVar(zipFile)).Items
        fileName = Dir(extractFolderPath & "\*.csv")
        Stream.Open
        Stream.Type = adTypeBinary
        Stream.LoadFromFile extractFolderPath & "\" & fileName
        GetFile = GetString(Stream.Read, "UTF-8")
        Stream.Close
    Else
        GetFile = StrConv(objHTTP.responseBody, vbUnicode)
    End If
    Set objHTTP = Nothing
    Set Stream = Nothing
    Set shellApp = Nothing
End Function

Sub SetDate(ByRef Val As Date, value As Variant)
    Dim str As String
    If IsNull(value) = False Then
        str = CStr(value)
        Val = ConvertISO8601ToDate(str)
    End If
End Sub

Public Function GetString(ByRef bin, ByVal encoding As String) As String
    Const adTypeBinary = 1
    Const adTypeText = 2
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeBinary
        .Write bin
        .Position = 0
        .Type = adTypeText
        .Charset = encoding
        GetString = .ReadText
        .Close
    End With
End Function


Function QueryRequest(strPath As String, Query As Dictionary) As Dictionary
    Dim objHTTP As XMLHTTP60
    Dim Params As Variant
    Dim strResult As String

    Dim QueryString() As String
    If UBound(Query.Items) > -1 Then
        ReDim QueryString(UBound(Query.Items))
        
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = True
        re.Pattern = "\[[0-9]+\]"
        
        Dim i As Integer
        For i = 0 To UBound(Query.Items)
            QueryString(i) = re.Replace(Query.Keys(i), "[]") & "=" & WorksheetFunction.EncodeURL(Query.Item(Query.Keys(i)))
        Next i
        strPath = Replace(strPath, "https://app.engn.jp/", "") & "?" & Join(QueryString, "&")
    End If
    Set objHTTP = New XMLHTTP60
    Dim strMethod As String
    strMethod = "GET"
    objHTTP.Open "GET", EndPoint & strPath, False
    objHTTP.setRequestHeader "Content-Type", "application/json;"
    objHTTP.setRequestHeader "Authorization", "Bearer " & Token
    objHTTP.Send
    
    strResult = ConvertString(objHTTP.responseBody, "UTF-8")
    
    If objHTTP.Status <> 200 Then
        Set QueryRequest = JsonConverter.ParseJson(strResult)
    Else
        Dim Json As Dictionary
        Set QueryRequest = JsonConverter.ParseJson(strResult)
    End If
    Set objHTTP = Nothing
End Function

Public Function ConvertString(ByRef utf8Text, ByVal encoding As String) As String
    Const adTypeBinary = 1
    Const adTypeText = 2
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeBinary
        .Write utf8Text
        .Position = 0
        .Type = adTypeText
        .Charset = encoding
        ConvertString = .ReadText
        .Close
    End With
End Function

Function SendFiles(httpMethod As WebMethod, strPath As String, Data As Dictionary, aryFiles() As String) As Dictionary
    ' Dim Data As Dictionary
    ' Set Data = BuildParams
    Dim Boundary As String
    Boundary = RandomString(24)
    Dim FilePath
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    Stream.Charset = "UTF-8"
    Stream.Open
    Stream.Position = 0
    Stream.Type = adTypeText
    Stream.WriteText "----" & Boundary & vbCrLf
    Stream.WriteText "Content-Disposition: form-data; name=""data""; filename=""base.json""" & vbCrLf
    Stream.WriteText "Content-Type: application/json" & vbCrLf
    Stream.WriteText "" & vbCrLf
    Stream.WriteText JsonConverter.ConvertToJson(Data)
        
    Dim p As Long
    If (Not aryFiles) <> -1 Then
        For Each FilePath In aryFiles
            AddAttachmentToStream Stream, FilePath, Boundary
        Next FilePath
    End If
    
    ChangeStreamType Stream, adTypeText
    Stream.WriteText vbCrLf & "----" & Boundary & "--" & vbCrLf
    ChangeStreamType Stream, adTypeBinary
    Stream.Position = 0
    
    Dim objHTTP As XMLHTTP60
    Dim Params As Variant
    Dim strResult As String
    
    Set objHTTP = New XMLHTTP60
    Dim strMethod As String
    strMethod = IIf(httpMethod = httpPost, "POST", "GET")
    objHTTP.Open strMethod, EndPoint & strPath, False
    objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=--" & Boundary
    objHTTP.setRequestHeader "Authorization", "Bearer " & Token
    Params = Stream.Read
    objHTTP.Send Params
    Set Stream = Nothing
    
    strResult = StrConv(objHTTP.responseBody, vbUnicode)
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

Function Delete(DeliveryId As Long) As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/" & DeliveryId
    Set dictResult = SendText(HttpDelete, strPath, New Dictionary, New Dictionary)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Delete = False
    Else
        Delete = True
    End If
    Set dictResult = Nothing
End Function

Function InsertCodeToArray(InsertCode As Dictionary) As Dictionary()
    Dim Key As Variant
    Dim i As Long
    i = 0
    Dim InsertCodes() As Dictionary
    If InsertCode.Count = 0 Then
        Exit Function
    End If
    ReDim InsertCodes(InsertCode.Count - 1) As Dictionary
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

Function StringToStatus(value As String) As StatusType
    Select Case value
    Case "EDIT"
        StringToStatus = EDIT
    Case "IMPORTING"
        StringToStatus = IMPORTING
    Case "RESERVE"
        StringToStatus = RESERVE
    Case "WAIT"
        StringToStatus = WAIT
    Case "SENDING"
        StringToStatus = SENDING
    Case "SENT"
        StringToStatus = SENT
    Case "FAILED"
        StringToStatus = FAILED
    End Select
End Function

Function CreateTemporaryFile(extension As String) As String
    Dim tempPath As String
    Dim tempFileName As String
    Dim fullPath As String
    Dim fileNum As Integer
    
    ' テンポラリフォルダのパスを取得
    tempPath = Environ("TEMP")

    ' テンポラリファイル名を生成
    tempFileName = "TempFile_" & RandomString(10) & "." & extension

    ' フルパスを生成
    CreateTemporaryFile = tempPath & "\" & tempFileName
End Function

Function CreateTemporaryFolder() As String
    Dim tempFolderPath As String
    Dim newTempFolderPath As String
    Dim fso As Object
    Dim uniqueID As String

    ' FileSystemObjectを作成
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Windowsの一時フォルダパスを取得
    tempFolderPath = Environ("TEMP")

    ' ユニークなフォルダ名を生成
    uniqueID = fso.GetTempName

    ' 新しいテンポラリフォルダのパスを作成
    newTempFolderPath = fso.BuildPath(tempFolderPath, uniqueID)

    ' テンポラリフォルダを作成
    If Not fso.FolderExists(newTempFolderPath) Then
        fso.CreateFolder newTempFolderPath
        CreateTemporaryFolder = newTempFolderPath
    Else
        CreateTemporaryFolder = ""
    End If
    ' オブジェクトを解放
    Set fso = Nothing
End Function


Function IndexOf(arr As Variant, value As Variant) As Long
    Dim i As Long
    ' 初期値を-1に設定
    IndexOf = -1
    
    ' 配列の要素をループ
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Sub ArrayToString(ByRef dict As Dictionary, strKey As String, aryString() As String)
    If (Not aryString) <> -1 Then
        Dim Key As Variant
        Dim i As Integer
        i = 0
        For Each Key In aryString
            dict.Add strKey & "[" & i & "]", Key
            i = i + 1
        Next Key
    End If
End Sub

Sub AddStringQuery(ByRef dict As Dictionary, strKey As String, value As String)
    If value <> "" Then
        dict.Add strKey, value
    End If
End Sub

Sub AddIntegerQuery(ByRef dict As Dictionary, strKey As String, value As Integer)
    If value > 0 Then
        dict.Add strKey, value
    End If
End Sub

Sub AddLongQuery(ByRef dict As Dictionary, strKey As String, value As Long)
    If value > 0 Then
        dict.Add strKey, value
    End If
End Sub

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

Function ConvertISO8601ToDate(isoDate As String) As Date
    Dim datePart As String
    Dim timePart As String
    Dim tzPart As String
    Dim dt As Date
    Dim timeZoneOffset As Integer
    
    ' ??????????
    datePart = Left(isoDate, 10)
    timePart = Mid(isoDate, 12, 8)
    tzPart = Mid(isoDate, 20, 3)
    
    dt = CDate(datePart & " " & timePart)
    
    timeZoneOffset = CInt(tzPart)
    
    ConvertISO8601ToDate = DateAdd("h", -timeZoneOffset, dt)
End Function

Private Sub AddAttachmentToStream(Stream As ADODB.Stream, FilePath As Variant, Boundary As String)
    Dim file As ADODB.Stream
    Set file = New ADODB.Stream
    file.Open
    file.Type = adTypeBinary
    file.LoadFromFile FilePath
    
    ChangeStreamType Stream, adTypeText
    Stream.WriteText vbCrLf & "----" & Boundary & vbCrLf
    Stream.WriteText "Content-Disposition: form-data; name=""file""; filename=""" & Dir(FilePath) & """" & vbCrLf
    Stream.WriteText "Content-Type: application/octet-stream" & vbCrLf
    Stream.WriteText "" & vbCrLf
    
    ChangeStreamType Stream, adTypeBinary
    Stream.Write file.Read()
    
    file.Close
    Set file = Nothing
End Sub

Private Sub ChangeStreamType(ByRef Stream As ADODB.Stream, adType As Integer)
    Dim p As Long
    p = Stream.Position
    Stream.Position = 0
    Stream.Type = adType
    If adType = adTypeText Then
        Stream.Charset = "UTF-8"
    End If
    Stream.Position = p
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

