VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Job"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Id As Long
Private m_Client As Blastengine
Private m_Error As New Dictionary
Private m_Percentage As Long
Private m_Status As String
Private m_SuccessCount As Long
Private m_FailedCount As Long
Private m_TotalCount As Long
Private m_ErrorFileUrl As String

Property Let Client(Client As Blastengine)
    Set m_Client = Client
End Property

Property Let Id(Id As Long)
    m_Id = Id
End Property

Public Function Error() As Dictionary()
    Dim strResult As String
    Dim Lines() As String
    Dim Line() As String
    Dim dict As Dictionary
    Dim i As Long
    Dim j As Long
    Dim aryError() As Dictionary
    
    strResult = m_Client.GetFile(ErrorFileUrl)
    Lines = Split(strResult, vbCrLf)
    ' ヘッダーを分割して配列に格納
    Headers = Split(Lines(0), ",")

    ' 結果を格納する配列のサイズを決定
    ReDim aryError(1 To UBound(Lines) - 1) As Dictionary

    ' データ行ごとに処理
    For i = 1 To UBound(Lines) - 1 ' ヘッダー分を差し引く
        ' 新しいDictionaryオブジェクトを作成
        Set dict = New Dictionary
        
        ' 現在の行をカンマで分割してフィールドごとに処理
        Line = Split(Lines(i), ",")
        
        For j = 0 To UBound(Headers)
            ' Dictionaryにヘッダー名をキーとしてデータを格納
            dict(Replace(Headers(j), """", "")) = Replace(Line(j), """", "")
        Next j
        
        ' Dictionaryを配列に格納
        Set aryError(i) = dict
    Next i
    Error = aryError
End Function

Property Get Status() As String
    Status = m_Status
End Property

Property Get Percentage() As Long
    Percentage = m_Percentage
End Property

Property Get ErrorFileUrl() As String
    ErrorFileUrl = m_ErrorFileUrl
End Property

Property Get IsError() As Long
    If m_Status = "FAILED" Then
        IsError = True
    ElseIf m_Status = "FINISHED" Then
        IsError = False
    End If
End Property


Function Finished() As Boolean
    Dim dictResult As Dictionary
    Dim strPath As String
    strPath = "/api/v1/deliveries/-/emails/import/" & m_Id
    Set dictResult = m_Client.SendText(httpGet, strPath, New Dictionary, New Dictionary)
    If dictResult.Exists("error_messages") Then
        ' Error
        Set m_Error = dictResult("error_messages")
        Finished = False
        Exit Function
    End If
    ' Debug.Print JsonConverter.ConvertToJson(dictResult)
    m_Percentage = CLng(dictResult("percentage"))
    m_Status = dictResult("status")
    m_SuccessCount = CLng(dictResult("success_count"))
    m_FailedCount = CLng(dictResult("failed_count"))
    m_TotalCount = CLng(dictResult("total_count"))
    m_ErrorFileUrl = dictResult("error_file_url")
    If m_Percentage = 100 Then
        Finished = True
    Else
        Finished = False
    End If
End Function
