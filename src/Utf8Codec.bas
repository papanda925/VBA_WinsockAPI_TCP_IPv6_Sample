Attribute VB_Name = "Utf8Codec"
Option Explicit

'===============================================================================
' Utf8Codec.bas
'-------------------------------------------------------------------------------
' VBAのString（内部ではUTF-16）と、ネットワーク送信用のUTF-8バイト列を
' 相互変換します。StrConvに任せるとWindowsの言語設定に依存する場合があるため、
' Windows APIで文字コードを明示しています。
'===============================================================================

Private Const CP_UTF8 As Long = 65001
Private Const MB_ERR_INVALID_CHARS As Long = &H8
Private Const WC_ERR_INVALID_CHARS As Long = &H80

#If VBA7 Then
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal codePage As Long, ByVal flags As Long, ByVal wideText As LongPtr, _
    ByVal wideCharCount As Long, ByVal outputBytes As LongPtr, _
    ByVal outputByteCount As Long, ByVal defaultChar As LongPtr, _
    ByVal usedDefaultChar As LongPtr) As Long

Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codePage As Long, ByVal flags As Long, ByVal inputBytes As LongPtr, _
    ByVal inputByteCount As Long, ByVal outputText As LongPtr, _
    ByVal outputCharCount As Long) As Long
#Else
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal codePage As Long, ByVal flags As Long, ByVal wideText As Long, _
    ByVal wideCharCount As Long, ByVal outputBytes As Long, _
    ByVal outputByteCount As Long, ByVal defaultChar As Long, _
    ByVal usedDefaultChar As Long) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codePage As Long, ByVal flags As Long, ByVal inputBytes As Long, _
    ByVal inputByteCount As Long, ByVal outputText As Long, _
    ByVal outputCharCount As Long) As Long
#End If

'StringをUTF-8へ変換し、実際のバイト数を返します。
'空文字の場合は0を返します。outputには安全のため1要素だけ確保します。
Public Function StringToUtf8(ByVal value As String, ByRef output() As Byte) As Long
    Dim requiredBytes As Long
    Dim convertedBytes As Long

    If Len(value) = 0 Then
        ReDim output(0 To 0)
        StringToUtf8 = 0
        Exit Function
    End If

    requiredBytes = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, _
                                        StrPtr(value), Len(value), 0, 0, 0, 0)
    If requiredBytes <= 0 Then
        Err.Raise vbObjectError + 2101, "StringToUtf8", _
                  "UTF-8変換に必要なバイト数を取得できませんでした。"
    End If

    ReDim output(0 To requiredBytes - 1)
    convertedBytes = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, _
                                         StrPtr(value), Len(value), _
                                         VarPtr(output(0)), requiredBytes, 0, 0)
    If convertedBytes <> requiredBytes Then
        Err.Raise vbObjectError + 2102, "StringToUtf8", _
                  "UTF-8への変換結果が想定した長さと一致しません。"
    End If

    StringToUtf8 = convertedBytes
End Function

'UTF-8バイト列の先頭byteCountバイトをVBAのStringへ戻します。
Public Function Utf8ToString(ByRef value() As Byte, ByVal byteCount As Long) As String
    Dim requiredChars As Long
    Dim convertedChars As Long
    Dim availableBytes As Long
    Dim firstIndex As Long
    Dim result As String

    If byteCount < 0 Then
        Err.Raise vbObjectError + 2103, "Utf8ToString", _
                  "byteCountに負の値は指定できません。"
    ElseIf byteCount = 0 Then
        Utf8ToString = vbNullString
        Exit Function
    End If

    firstIndex = LBound(value)
    availableBytes = UBound(value) - firstIndex + 1
    If byteCount > availableBytes Then
        Err.Raise vbObjectError + 2104, "Utf8ToString", _
                  "byteCountが受信バッファの範囲を超えています。"
    End If

    requiredChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, _
                                        VarPtr(value(firstIndex)), byteCount, 0, 0)
    If requiredChars <= 0 Then
        Err.Raise vbObjectError + 2105, "Utf8ToString", _
                  "受信データをUTF-8として解釈できませんでした。"
    End If

    result = String$(requiredChars, vbNullChar)
    convertedChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, _
                                         VarPtr(value(firstIndex)), byteCount, _
                                         StrPtr(result), requiredChars)
    If convertedChars <> requiredChars Then
        Err.Raise vbObjectError + 2106, "Utf8ToString", _
                  "UTF-8からの変換結果が想定した長さと一致しません。"
    End If

    Utf8ToString = result
End Function
