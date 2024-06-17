VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Client As Blastengine
Private m_DeliveryId As Long
Private m_From As From
Private m_To As String
Private m_Cc() As String
Private m_Bcc() As String
Private m_InsertCode As Dictionary
Private m_Subject As String
Private m_ListUnsubscribe As ListUnsubscribe
Private m_Encode As String
Private m_TextPart As String
Private m_HtmlPart As String
Private m_Error As Dictionary
Private m_Attachments() As String
Private m_Response As Dictionary

Private Sub Class_Initialize()
    m_Encode = "UTF-8"
    Set m_InsertCode = New Dictionary
    Set m_Error = New Dictionary
End Sub

Property Let client(client As Blastengine)
    Set m_Client = client
End Property

Function Send() As Boolean
    Dim strPath As String
    strPath = "/api/v1/deliveries/transaction"
    If (Not m_Attachments) = -1 Then
        Send = SendText(HttpPost, strPath)
    Else
        Send = SendAttachment(HttpPost, strPath)
    End If
End Function

Sub HandleError(Content As String)
    Dim Error As Dictionary
    Set Error = JsonConverter.ParseJson(Content)
    If Error.Exists("error_messages") Then
        Set m_Error = Error("error_messages")
    End If
End Sub

Function SendText(method As WebMethod, strPath As String) As Boolean
    Dim client As New WebClient
    client.BaseUrl = m_Client.EndPoint
    Dim Request As New WebRequest
    Request.Resource = strPath
    Request.method = WebMethod.HttpPost
    Request.AddHeader "Authorization", "Bearer " & m_Client.Token
    Request.AddHeader "Content-Type", "application/json"
    Set Request.Body = BuildParams
    WebHelpers.EnableLogging = False ' デバッグ用
    Dim Response As WebResponse
    Set Response = client.Execute(Request)
    If Response.StatusCode <> Created Then
        HandleError (Response.Content)
        SendText = False
    Else
        ' JSONをパース（デフォルトでは失敗するため、VBA-JSONを利用）
        Dim Json As Dictionary
        Set Json = JsonConverter.ParseJson(Response.Content)
        m_DeliveryId = Json("delivery_id")
        SendText = True
    End If
    Set client = Nothing
    Set Request = Nothing
    Set Response = Nothing
End Function

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

Function SendAttachment(httpMethod As WebMethod, strPath As String) As Boolean
    Dim Data As Dictionary
    Set Data = BuildParams
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
    stream.WriteText JsonConverter.ConvertToJson(BuildParams)
    
    Dim p As Long
    If (Not m_Attachments) <> -1 Then
        
        For Each FilePath In m_Attachments
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
    objHTTP.Open strMethod, m_Client.EndPoint & strPath, False
    objHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=--" & Boundary
    objHTTP.setRequestHeader "Authorization", "Bearer " & m_Client.Token
    Params = stream.Read
    objHTTP.Send Params
    Set stream = Nothing
    
    strResult = StrConv(objHTTP.responsebody, vbUnicode)
    If objHTTP.Status <> 201 Then
        HandleError strResult
        SendAttachment = False
    Else
        Dim Json As Dictionary
        Set Json = JsonConverter.ParseJson(strResult)
        m_DeliveryId = Json("delivery_id")
        SendAttachment = True
    End If
    Set objHTTP = Nothing
End Function

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

Property Get DeliveryId() As Long
    DeliveryId = m_DeliveryId
End Property

Property Get Error() As Dictionary
    Set Error = m_Error
End Property

Sub From(Email As String, Optional Name As String = "")
    Dim From As From
    From.Email = Email
    From.Name = Name
    m_From = From
End Sub

Property Let Email(ByVal Email As String)
    m_To = Email
End Property

Property Let Cc(i As Long, ByVal Email As String)
    ReDim Preserve m_Cc(i) As String
    m_Cc(i) = Email
End Property

Property Let Bcc(i As Long, ByVal Email As String)
    ReDim Preserve m_Bcc(i) As String
    m_Bcc(i) = Email
End Property

Property Let Attachments(i As Long, ByVal FilePath As String)
    ReDim Preserve m_Attachments(i) As String
    m_Attachments(i) = FilePath
End Property

Property Let InsertCode(ByVal InsertCode As Dictionary)
    Set m_InsertCode = InsertCode
End Property

Property Let Subject(ByVal Subject As String)
    m_Subject = Subject
End Property

Property Let Encode(ByVal Encode As String)
    m_Encode = Encode
End Property

Property Let TextPart(ByVal TextPart As String)
    m_TextPart = TextPart
End Property

Property Let HtmlPart(ByVal HtmlPart As String)
    m_HtmlPart = HtmlPart
End Property

Property Let ListUnsubscribe(ListUnsubscribe As ListUnsubscribe)
    m_ListUnsubscribe = ListUnsubscribe
End Property


Function BuildParams() As Dictionary
    Set BuildParams = New Dictionary
    Dim From As Dictionary
    Set From = New Dictionary
    From.Add "email", m_From.Email
    If m_From.Name <> "" Then
        From.Add "name", m_From.Name
    End If
    BuildParams.Add "from", From
    BuildParams.Add "to", m_To
    If (Not m_Cc) <> -1 Then
        BuildParams.Add "cc", m_Cc
    End If
    If (Not m_Bcc) <> -1 Then
        BuildParams.Add "bcc", m_Bcc
    End If
    Dim Key As Variant
    Dim i As Long
    i = 0
    Dim InsertCodes() As Dictionary
    ReDim InsertCodes(UBound(m_InsertCode.Keys)) As Dictionary
    For Each Key In m_InsertCode.Keys
        Dim InsertCode As Dictionary
        Set InsertCode = New Dictionary
        InsertCode.Add "key", "__" & Key & "__"
        InsertCode.Add "value", m_InsertCode.Item(Key)
        Set InsertCodes(i) = InsertCode
    Next Key
    If (Not InsertCodes) <> -1 Then
        BuildParams.Add "insert_code", InsertCodes
    End If
    BuildParams.Add "subject", m_Subject
    BuildParams.Add "text_part", m_TextPart
    If m_HtmlPart <> "" Then
        BuildParams.Add "html_part", m_HtmlPart
    End If
    BuildParams.Add "encode", m_Encode
    
    If m_ListUnsubscribe.Email <> "" Or m_ListUnsubscribe.Url <> "" Then
        Dim ListUnsubscribe As Dictionary
        Set ListUnsubscribe = New Dictionary
        If m_ListUnsubscribe.Email <> "" Then
            ListUnsubscribe.Add "mailto", "mailto:" & m_ListUnsubscribe.Email
        End If
        If m_ListUnsubscribe.Url <> "" Then
            ListUnsubscribe.Add "url", "mailto:" & m_ListUnsubscribe.Url
        End If
        BuildParams.Add "list_unsubscribe", ListUnsubscribe
    End If

End Function
