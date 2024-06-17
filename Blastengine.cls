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

    '// ������� UTF8 �ɃG���R�[�h���A�o�C�g�z��ɕϊ�
    bytes = objUTF8.GetBytes_4(s)

    '// �n�b�V���l���v�Z�i�o�C�i���j
    hash = objSHA256.ComputeHash_2((bytes))

    '// �o�C�i����16�i��������ɕϊ�
    For i = 1 To UBound(hash) + 1
        wk = wk & Right("0" & Hex(AscB(MidB(hash, i, 1))), 2)
    Next i

    '// ���ʂ�Ԃ�
    SHA256 = LCase(wk)

End Function

' https://ribbit.konomi.app/vba/base64-encode/
'*-------------------------------------------------------------
'* �e�L�X�g��Base64�ŃG���R�[�h
'*
'* @param text �ϊ�����l
'* @return Base64�t�H�[�}�b�g�f�[�^
'*-------------------------------------------------------------
Public Function EncodeToBase64(ByRef text As String) As String

  ' �I�u�W�F�N�g�̏���
  Dim node As Object
  Set node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")

  ' �G���R�[�h
  node.DataType = "bin.base64"
  node.nodeTypedValue = ConvertToBinary(text)

  ' �֐��Ŏ�菜���Ȃ����s���폜���ĕԋp
  EncodeToBase64 = Replace(node.text, vbLf, "")
End Function

'*-------------------------------------------------------------
'* ��������o�C�i���f�[�^�ɕϊ�
'*
'* @param text �ϊ�����l
'* @return �o�C�i���f�[�^
'*-------------------------------------------------------------
Private Function ConvertToBinary(ByRef text As String)

  ' �I�u�W�F�N�g�̏���
  Dim BinaryStream As Object
  Set BinaryStream = CreateObject("ADODB.Stream")

  ' Stream�̐ݒ�
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