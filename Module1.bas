Attribute VB_Name = "Module1"
Private Declare PtrSafe Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringW" ( _
    ByVal pbBinary As LongPtr, _
    ByVal cbBinary As Long, _
    ByVal dwFlags As Long, _
    ByVal pszString As LongPtr, _
    ByVal pcchString As LongPtr _
    ) As Long

Private Const CRYPT_STRING_BASE64 As Long = &H1&

'バイト配列をBASE64でエンコードしたUnicode文字列にする関数
'この関数にバイト配列を渡す際に、UTF-8変換したバイト配列を渡すのか
'SJIS変換したバイト配列を渡すのか、UTF16変換(無変換)したバイト配列を渡すのかで
'当然ながら出力される文字列が変わってきます。
Public Function CryptBytesToString(ByRef bData() As Byte) As String
    Dim pbBinary As LongPtr
    Dim cbBinary As Long
    If UBound(bData) = -1 Then
        Exit Function
    End If

    'バイト配列の先頭アドレス　と　長さを取得
    pbBinary = VarPtr(bData(0))
    cbBinary = UBound(bData) - LBound(bData) + 1

    Dim nBufferSize As Long
    '変換後文字列の長さを取得し、変換後文字列を格納するだけのバッファを用意
    If CryptBinaryToString(pbBinary, cbBinary, CRYPT_STRING_BASE64, _
                                0, VarPtr(nBufferSize)) Then
        CryptBytesToString = String(nBufferSize, vbNullChar)
        '必要な長さ分用意した空文字列バッファの先頭アドレス(StrPtr)から変換結果で上書き
        If CryptBinaryToString(pbBinary, cbBinary, CRYPT_STRING_BASE64, _
                                StrPtr(CryptBytesToString), VarPtr(nBufferSize)) Then
        End If
    End If
End Function


