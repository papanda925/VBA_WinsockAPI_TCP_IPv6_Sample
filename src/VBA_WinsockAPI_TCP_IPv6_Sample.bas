Attribute VB_Name = "VBA_WinsockAPI_TCP_IPv6_Sample"
Option Explicit

'===============================================================================
' Excel VBA Winsock TCP/IPv6 クライアント・サーバー サンプル
'-------------------------------------------------------------------------------
' 目的
'   WindowsのWinsock APIをVBAから直接呼び出し、IPv6ループバックアドレス
'   「::1」上でTCP通信を試します。外部のexe、PowerShell、curlは使用しません。
'
' 安全な試し方
'   最初はDEFAULT_SERVER_ADDRESSを「::1」のまま使用してください。
'   ::1は自分のPCだけを指すIPv6ループバックアドレスです。
'
' サーバーを別Excelプロセスにする理由
'   accept/recvは相手からデータが届くまで待つ同期APIです。同じExcelプロセスで
'   サーバーとクライアントを動かすと操作画面まで待たされるため、学習用として
'   受信側を別プロセスに分けています。DoEventsだけではAPIの待機は解除できません。
'===============================================================================

Private Const AF_INET6 As Long = 23
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const IPPROTO_IPV6 As Long = 41
Private Const IPV6_V6ONLY As Long = 27
Private Const SOL_SOCKET As Long = &HFFFF&
' WindowsではSO_REUSEADDRが、使用中ポートへの強制bindや不定な配送を招く
' 場合があります。本教材は同じポートを共有しないため、排他的に確保します。
Private Const SO_EXCLUSIVEADDRUSE As Long = -5
Private Const SO_RCVTIMEO As Long = &H1006&
Private Const SO_SNDTIMEO As Long = &H1005&
Private Const SOCKET_ERROR As Long = -1
Private Const INVALID_SOCKET_VALUE As Long = -1

Private Const DEFAULT_SERVER_ADDRESS As String = "::1"
Private Const DEFAULT_SERVER_PORT As Long = 60052
Private Const LISTEN_BACKLOG As Long = 5
Private Const RECEIVE_BUFFER_SIZE As Long = 8192
Private Const SOCKET_TIMEOUT_MS As Long = 5000

' 別プロセスで起動したExcelへの参照を、起動マクロの終了後も保持します。
' ローカル変数だけにするとCOM参照が解放され、環境によってはサーバー用Excelが
' 意図せず終了するため、モジュール変数として寿命を明示します。
Private m_serverExcelInstance As Excel.Application
Private m_serverWorkbookInstance As Excel.Workbook

'Win64版のWSADATAは、32bit版と一部メンバーの並び順が異なります。
'条件付きコンパイルでWindows SDKと同じ並びになるよう定義します。
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
#If Win64 Then
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
#End If
    'WinSockのヘッダーではchar配列です。VBAの固定長StringはUnicodeで
    '1文字2バイトになるため、構造体のサイズを崩さないByte配列を使います。
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
#If Not Win64 Then
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As LongPtr
#End If
End Type

'Windows SDKのSOCKADDR_IN6（28バイト）に対応する構造体です。
'IPv6アドレス本体は16バイトなので、Byte配列として保持します。
Private Type SOCKADDR_IN6
    sin6_family As Integer
    sin6_port As Integer
    sin6_flowinfo As Long
    sin6_addr(0 To 15) As Byte
    sin6_scope_id As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" ( _
    ByVal requestedVersion As Integer, ByRef data As WSADATA) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function socket Lib "ws2_32.dll" ( _
    ByVal addressFamily As Long, ByVal socketType As Long, _
    ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function closesocket Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr) As Long
Private Declare PtrSafe Function bind Lib "ws2_32.dll" Alias "bind" ( _
    ByVal socketHandle As LongPtr, ByRef socketAddress As SOCKADDR_IN6, _
    ByVal addressLength As Long) As Long
Private Declare PtrSafe Function listen Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByVal backlog As Long) As Long
Private Declare PtrSafe Function accept Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByRef clientAddress As SOCKADDR_IN6, _
    ByRef addressLength As Long) As LongPtr
Private Declare PtrSafe Function connect Lib "ws2_32.dll" Alias "connect" ( _
    ByVal socketHandle As LongPtr, ByRef socketAddress As SOCKADDR_IN6, _
    ByVal addressLength As Long) As Long
Private Declare PtrSafe Function send Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByRef buffer As Any, _
    ByVal bufferLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function recv Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByRef buffer As Any, _
    ByVal bufferLength As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function setsockopt Lib "ws2_32.dll" ( _
    ByVal socketHandle As LongPtr, ByVal level As Long, ByVal optionName As Long, _
    ByRef optionValue As Any, ByVal optionLength As Long) As Long
Private Declare PtrSafe Function htons Lib "ws2_32.dll" ( _
    ByVal hostValue As Integer) As Integer
Private Declare PtrSafe Function ntohs Lib "ws2_32.dll" ( _
    ByVal networkValue As Integer) As Integer
Private Declare PtrSafe Function InetPtonW Lib "ws2_32.dll" ( _
    ByVal addressFamily As Long, ByVal addressText As LongPtr, _
    ByRef binaryAddress As Any) As Long
Private Declare PtrSafe Function InetNtopW Lib "ws2_32.dll" ( _
    ByVal addressFamily As Long, ByRef binaryAddress As Any, _
    ByVal outputText As LongPtr, ByVal outputCharCount As LongPtr) As LongPtr
#Else
'このプロジェクトはVBA 7（Excel 2010以降）を対象とします。
#End If

'-------------------------------------------------------------------------------
' 初心者向けの実行マクロ
'-------------------------------------------------------------------------------

'受信側を別のExcelプロセスで開始します。
'同じブックが読み取り専用で開かれても、Traceシートへの記録はそのプロセス内で
'確認できます。終了後、保存せずにサーバー側Excelを閉じてください。
Public Sub StartTcpIPv6ServerInNewExcel()
    'Excel VBA自身の型を使う早期バインディングです。Open/Closeの名前付き引数を
    'コンパイル時に確認でき、遅延バインディングの引数解決エラーを避けます。
    Dim errorText As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "先にこのブックをマクロ有効ブックとして保存してください。", _
               vbInformation
        Exit Sub
    End If

    If Not m_serverExcelInstance Is Nothing Then
        MsgBox "このブックから起動したサーバー用Excelは、すでに存在します。", _
               vbInformation
        Exit Sub
    End If

    On Error GoTo ERROR_HANDLER

    Set m_serverExcelInstance = CreateObject("Excel.Application")
    m_serverExcelInstance.Visible = True
    Set m_serverWorkbookInstance = m_serverExcelInstance.Workbooks.Open( _
                         Filename:=ThisWorkbook.FullName, _
                         UpdateLinks:=False, _
                         ReadOnly:=True, _
                         IgnoreReadOnlyRecommended:=True)

    'Application.Runから呼ぶプロシージャはPublicである必要があります。
    m_serverExcelInstance.Run WorkbookMacroReference( _
                        CStr(m_serverWorkbookInstance.Name), "ScheduleTcpIPv6Server")

    MsgBox "IPv6 TCPサーバー用のExcelを起動しました。" & vbCrLf & _
           "サーバー側のTraceシートにSERVER_READYが表示されてから送信してください。", _
           vbInformation
    Exit Sub

ERROR_HANDLER:
    errorText = Err.Description
    '起動途中で失敗した場合は、見えないExcelプロセスや読み取り専用ブックを
    '残さないよう、このマクロが作成したものだけを後始末します。
    On Error Resume Next
    If Not m_serverWorkbookInstance Is Nothing Then _
        m_serverWorkbookInstance.Close SaveChanges:=False
    If Not m_serverExcelInstance Is Nothing Then m_serverExcelInstance.Quit
    Set m_serverWorkbookInstance = Nothing
    Set m_serverExcelInstance = Nothing
    On Error GoTo 0
    MsgBox "サーバー用Excelを起動できませんでした。" & vbCrLf & _
           errorText, vbExclamation
End Sub

'Application.Runへすぐ応答を返してから、1秒後にブロッキング処理を始めます。
Public Sub ScheduleTcpIPv6Server()
    Application.OnTime Now + TimeSerial(0, 0, 1), _
                       WorkbookMacroReference( _
                           ThisWorkbook.Name, "RunTcpIPv6Server")
End Sub

'ネイティブAPIを呼ぶ前に、VBA側の構造体サイズがWindows SDKと一致するかを
'単独で確認できます。32bit/64bit両方の実機テストで最初に実行してください。
Public Sub CheckTcpIPv6StructureSizes()
    ValidateWinsockStructureSizes
    MsgBox "構造体サイズは想定どおりです。" & vbCrLf & _
           "SOCKADDR_IN6=28 bytes", vbInformation
End Sub

Public Sub SendHello()
    SendTcpIPv6Command "HELLO", vbNullString
End Sub

Public Sub SendEchoSample()
    SendTcpIPv6Command "ECHO", "こんにちは。IPv6 TCP通信の確認です。"
End Sub

'サーバーへ終了要求を送り、応答を受け取ってから受信ループを終了させます。
Public Sub StopTcpIPv6Server()
    SendTcpIPv6Command "QUIT", vbNullString
End Sub

' StartTcpIPv6ServerInNewExcelで作成した読み取り専用ブックを保存せず閉じ、
' そのためだけに作成したExcelプロセスを終了します。StopTcpIPv6Serverの後、
' サーバー側TraceのSHUTDOWNを確認してから実行してください。
Public Sub CloseTcpIPv6ServerExcel()
    Dim closeError As String

    If m_serverExcelInstance Is Nothing Then Exit Sub

    On Error Resume Next
    If Not m_serverWorkbookInstance Is Nothing Then
        m_serverWorkbookInstance.Close SaveChanges:=False
    End If
    m_serverExcelInstance.Quit
    If Err.Number <> 0 Then closeError = Err.Description
    Set m_serverWorkbookInstance = Nothing
    Set m_serverExcelInstance = Nothing
    On Error GoTo 0

    If Len(closeError) > 0 Then
        MsgBox "サーバー用Excelの終了処理を確認してください。" & vbCrLf & _
               closeError, vbExclamation
    End If
End Sub

'-------------------------------------------------------------------------------
' TCP/IPv6 サーバー
'-------------------------------------------------------------------------------

Public Sub RunTcpIPv6Server()
    Dim winsockData As WSADATA
    Dim listenSocket As LongPtr
    Dim clientSocket As LongPtr
    Dim serverAddress As SOCKADDR_IN6
    Dim clientAddress As SOCKADDR_IN6
    Dim clientAddressLength As Long
    Dim returnCode As Long
    Dim optionValue As Long
    Dim winsockStarted As Boolean
    Dim exitRequested As Boolean
    Dim requestText As String
    Dim responseText As String
    Dim fields As Variant
    Dim traceId As String
    Dim commandName As String
    Dim payload As String
    Dim serverTraceId As String
    Dim startedAt As Double

    On Error GoTo ERROR_HANDLER

    listenSocket = INVALID_SOCKET_VALUE
    clientSocket = INVALID_SOCKET_VALUE
    serverTraceId = NewTraceId("TCP6-SERVER")
    startedAt = Timer

    WriteTrace serverTraceId, "SERVER", "STARTUP", "LOCAL", _
               "Winsock 2.2を初期化します。", "START", 0

    ValidateWinsockStructureSizes
    returnCode = WSAStartup(MakeWord(2, 2), winsockData)
    If returnCode <> 0 Then
        Err.Raise vbObjectError + 2201, "RunTcpIPv6Server", _
                  "WSAStartupに失敗しました。エラーコード=" & CStr(returnCode)
    End If
    winsockStarted = True

    listenSocket = socket(AF_INET6, SOCK_STREAM, IPPROTO_TCP)
    If listenSocket = INVALID_SOCKET_VALUE Then
        RaiseLastSocketError "IPv6 TCPソケットを作成できませんでした。"
    End If

    '同じポートを別プロセスが同時に奪わないよう、bind前に排他利用を設定します。
    '設定済みポートへ2個目の教材サーバーを起動した場合はbindエラーになります。
    optionValue = 1
    returnCode = setsockopt(listenSocket, SOL_SOCKET, SO_EXCLUSIVEADDRUSE, _
                            optionValue, LenB(optionValue))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "SO_EXCLUSIVEADDRUSEを設定できませんでした。"
    End If

    'この教材はIPv6を明確に学ぶため、IPv4-mapped IPv6の受け付けを禁止します。
    optionValue = 1
    returnCode = setsockopt(listenSocket, IPPROTO_IPV6, IPV6_V6ONLY, _
                            optionValue, LenB(optionValue))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "IPV6_V6ONLYを設定できませんでした。"
    End If

    SetSocketTimeouts listenSocket, SOCKET_TIMEOUT_MS

    If Not FillIPv6Address(serverAddress, DEFAULT_SERVER_ADDRESS, DEFAULT_SERVER_PORT) Then
        Err.Raise vbObjectError + 2202, "RunTcpIPv6Server", _
                  "サーバーIPv6アドレスを変換できませんでした。"
    End If

    returnCode = bind(listenSocket, serverAddress, LenB(serverAddress))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "IPv6アドレスとポートのbindに失敗しました。"
    End If

    returnCode = listen(listenSocket, LISTEN_BACKLOG)
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "listenに失敗しました。"
    End If

    WriteTrace serverTraceId, "SERVER", "SERVER_READY", "LOCAL", _
               EndpointText(serverAddress) & " で接続を待ちます。", _
               "READY", ElapsedMilliseconds(startedAt)

    Do While Not exitRequested
        clientAddressLength = LenB(clientAddress)

        'acceptはクライアントが接続するまで戻りません。
        'このためサーバーを別Excelプロセスで動かしています。
        clientSocket = accept(listenSocket, clientAddress, clientAddressLength)
        If clientSocket = INVALID_SOCKET_VALUE Then
            RaiseLastSocketError "acceptに失敗しました。"
        End If

        SetSocketTimeouts clientSocket, SOCKET_TIMEOUT_MS
        WriteTrace serverTraceId, "SERVER", "ACCEPT", "IN", _
                   EndpointText(clientAddress) & " から接続しました。", _
                   "OK", ElapsedMilliseconds(startedAt)

        requestText = ReceiveText(clientSocket)
        fields = Split(TrimLineEnding(requestText), vbTab, 3)

        If UBound(fields) < 1 Or Not IsSafeTraceId(CStr(fields(0))) Then
            traceId = serverTraceId
            commandName = "INVALID"
            payload = vbNullString
        Else
            traceId = CStr(fields(0))
            commandName = UCase$(CStr(fields(1)))
            If UBound(fields) >= 2 Then payload = CStr(fields(2)) Else payload = vbNullString
        End If

        WriteTrace traceId, "SERVER", "REQUEST_RECEIVED", "IN", _
                   "Command=" & commandName & ", Payload=" & payload, _
                   "OK", ElapsedMilliseconds(startedAt)

        Select Case commandName
            Case "HELLO"
                responseText = traceId & vbTab & "OK" & vbTab & _
                               "HELLO from the Excel VBA IPv6 TCP server." & vbLf
            Case "ECHO"
                responseText = traceId & vbTab & "OK" & vbTab & payload & vbLf
            Case "QUIT"
                responseText = traceId & vbTab & "OK" & vbTab & _
                               "The IPv6 TCP server will stop." & vbLf
                exitRequested = True
            Case Else
                responseText = traceId & vbTab & "ERROR" & vbTab & _
                               "Unknown command." & vbLf
        End Select

        SendAllText clientSocket, responseText
        WriteTrace traceId, "SERVER", "RESPONSE_SENT", "OUT", _
                   TrimLineEnding(responseText), "SENT", _
                   ElapsedMilliseconds(startedAt)

        closesocket clientSocket
        clientSocket = INVALID_SOCKET_VALUE
    Loop

    WriteTrace serverTraceId, "SERVER", "SHUTDOWN", "LOCAL", _
               "QUITを受信したため待受を終了します。", "COMPLETED", _
               ElapsedMilliseconds(startedAt)

CLEANUP:
    If clientSocket <> INVALID_SOCKET_VALUE Then closesocket clientSocket
    If listenSocket <> INVALID_SOCKET_VALUE Then closesocket listenSocket
    If winsockStarted Then WSACleanup
    Exit Sub

ERROR_HANDLER:
    WriteTrace serverTraceId, "SERVER", "ERROR", "LOCAL", _
               Err.Description, "FAILED", ElapsedMilliseconds(startedAt)
    MsgBox "IPv6 TCPサーバーでエラーが発生しました。" & vbCrLf & _
           Err.Description, vbExclamation
    Resume CLEANUP
End Sub

'-------------------------------------------------------------------------------
' TCP/IPv6 クライアント
'-------------------------------------------------------------------------------

Private Function SendTcpIPv6Command( _
    ByVal commandName As String, _
    ByVal payload As String) As Boolean
    Dim winsockData As WSADATA
    Dim clientSocket As LongPtr
    Dim serverAddress As SOCKADDR_IN6
    Dim returnCode As Long
    Dim winsockStarted As Boolean
    Dim traceId As String
    Dim requestText As String
    Dim responseText As String
    Dim startedAt As Double

    On Error GoTo ERROR_HANDLER

    clientSocket = INVALID_SOCKET_VALUE
    traceId = NewTraceId("TCP6")
    startedAt = Timer
    requestText = traceId & vbTab & UCase$(commandName) & vbTab & payload & vbLf

    WriteTrace traceId, "CLIENT", "STARTUP", "LOCAL", _
               "Winsock 2.2を初期化します。", "START", 0

    ValidateWinsockStructureSizes
    returnCode = WSAStartup(MakeWord(2, 2), winsockData)
    If returnCode <> 0 Then
        Err.Raise vbObjectError + 2211, "SendTcpIPv6Command", _
                  "WSAStartupに失敗しました。エラーコード=" & CStr(returnCode)
    End If
    winsockStarted = True

    clientSocket = socket(AF_INET6, SOCK_STREAM, IPPROTO_TCP)
    If clientSocket = INVALID_SOCKET_VALUE Then
        RaiseLastSocketError "IPv6 TCPクライアントソケットを作成できませんでした。"
    End If
    SetSocketTimeouts clientSocket, SOCKET_TIMEOUT_MS

    If Not FillIPv6Address(serverAddress, DEFAULT_SERVER_ADDRESS, DEFAULT_SERVER_PORT) Then
        Err.Raise vbObjectError + 2212, "SendTcpIPv6Command", _
                  "接続先IPv6アドレスを変換できませんでした。"
    End If

    WriteTrace traceId, "CLIENT", "CONNECT", "OUT", _
               EndpointText(serverAddress) & " へ接続します。", _
               "START", ElapsedMilliseconds(startedAt)

    returnCode = connect(clientSocket, serverAddress, LenB(serverAddress))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "IPv6 TCPサーバーへ接続できませんでした。"
    End If

    WriteTrace traceId, "CLIENT", "CONNECTED", "OUT", _
               EndpointText(serverAddress), "OK", ElapsedMilliseconds(startedAt)

    SendAllText clientSocket, requestText
    WriteTrace traceId, "CLIENT", "REQUEST_SENT", "OUT", _
               "Command=" & UCase$(commandName) & ", Payload=" & payload, _
               "SENT", ElapsedMilliseconds(startedAt)

    responseText = ReceiveText(clientSocket)
    ValidateResponseTraceId responseText, traceId
    WriteTrace traceId, "CLIENT", "RESPONSE_RECEIVED", "IN", _
               TrimLineEnding(responseText), "COMPLETED", _
               ElapsedMilliseconds(startedAt)

    MsgBox "IPv6 TCPサーバーから応答を受信しました。" & vbCrLf & _
           TrimLineEnding(responseText), vbInformation
    SendTcpIPv6Command = True

CLEANUP:
    If clientSocket <> INVALID_SOCKET_VALUE Then closesocket clientSocket
    If winsockStarted Then WSACleanup
    Exit Function

ERROR_HANDLER:
    WriteTrace traceId, "CLIENT", "ERROR", "LOCAL", _
               Err.Description, "FAILED", ElapsedMilliseconds(startedAt)
    MsgBox "IPv6 TCPクライアントでエラーが発生しました。" & vbCrLf & _
           Err.Description, vbExclamation
    Resume CLEANUP
End Function

'-------------------------------------------------------------------------------
' Winsock補助処理
'-------------------------------------------------------------------------------

Private Function FillIPv6Address(ByRef result As SOCKADDR_IN6, _
                                 ByVal addressText As String, _
                                 ByVal portNumber As Long) As Boolean
    If portNumber < 0 Or portNumber > 65535 Then
        Err.Raise vbObjectError + 2221, "FillIPv6Address", _
                  "ポート番号は0～65535で指定してください。"
    End If

    result.sin6_family = AF_INET6
    result.sin6_port = htons(ToSignedInteger(portNumber))
    result.sin6_flowinfo = 0
    result.sin6_scope_id = 0

    FillIPv6Address = (InetPtonW(AF_INET6, StrPtr(addressText), _
                       result.sin6_addr(0)) = 1)
End Function

Private Sub SetSocketTimeouts(ByVal socketHandle As LongPtr, ByVal timeoutMs As Long)
    Dim returnCode As Long

    returnCode = setsockopt(socketHandle, SOL_SOCKET, SO_RCVTIMEO, _
                            timeoutMs, LenB(timeoutMs))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "受信タイムアウトを設定できませんでした。"
    End If

    returnCode = setsockopt(socketHandle, SOL_SOCKET, SO_SNDTIMEO, _
                            timeoutMs, LenB(timeoutMs))
    If returnCode = SOCKET_ERROR Then
        RaiseLastSocketError "送信タイムアウトを設定できませんでした。"
    End If
End Sub

'TCPのsendは、要求した全バイトを1回で送る保証がありません。
'戻り値の送信済みバイト数を確認し、残りがなくなるまで繰り返します。
Private Sub SendAllText(ByVal socketHandle As LongPtr, ByVal value As String)
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim sentTotal As Long
    Dim sentNow As Long

    byteCount = StringToUtf8(value, bytes)
    Do While sentTotal < byteCount
        sentNow = send(socketHandle, bytes(sentTotal), byteCount - sentTotal, 0)
        If sentNow = SOCKET_ERROR Then
            RaiseLastSocketError "sendに失敗しました。"
        ElseIf sentNow = 0 Then
            Err.Raise vbObjectError + 2222, "SendAllText", _
                      "sendが0を返したため、送信を継続できません。"
        End If
        sentTotal = sentTotal + sentNow
    Loop
End Sub

'TCPはメッセージ境界を持たないため、1回のrecvだけで要求全体を受け取れるとは
'限りません。このサンプルはLFを終端文字とし、LFが届くまでrecvを繰り返します。
Private Function ReceiveText(ByVal socketHandle As LongPtr) As String
    Dim buffer(0 To RECEIVE_BUFFER_SIZE - 1) As Byte
    Dim receivedBytes As Long
    Dim totalBytes As Long
    Dim scanStart As Long
    Dim byteIndex As Long
    Dim lineCompleted As Boolean

    Do
        If totalBytes >= RECEIVE_BUFFER_SIZE Then
            Err.Raise vbObjectError + 2223, "ReceiveText", _
                      "受信データが教材の上限8,192バイトを超えました。"
        End If

        scanStart = totalBytes
        receivedBytes = recv(socketHandle, buffer(totalBytes), _
                             RECEIVE_BUFFER_SIZE - totalBytes, 0)
        If receivedBytes = SOCKET_ERROR Then
            RaiseLastSocketError "recvに失敗またはタイムアウトしました。"
        ElseIf receivedBytes = 0 Then
            Err.Raise vbObjectError + 2224, "ReceiveText", _
                      "終端文字を受信する前に相手が接続を閉じました。"
        End If

        totalBytes = totalBytes + receivedBytes

        ' LFが受信塊の末尾にあるとは限りません。今回増えた範囲を先頭から調べ、
        ' 最初のLFまでを1要求として採用します。LF後の余分なデータは、この教材が
        ' 1接続1要求でソケットを閉じるため破棄します。
        For byteIndex = scanStart To totalBytes - 1
            If buffer(byteIndex) = 10 Then
                totalBytes = byteIndex + 1
                lineCompleted = True
                Exit For
            End If
        Next byteIndex
        If lineCompleted Then Exit Do
    Loop

    ReceiveText = Utf8ToString(buffer, totalBytes)
End Function

Private Function EndpointText(ByRef value As SOCKADDR_IN6) As String
    Dim addressBuffer As String
    Dim convertedPointer As LongPtr
    Dim nullPosition As Long

    addressBuffer = String$(46, vbNullChar)
    convertedPointer = InetNtopW(AF_INET6, value.sin6_addr(0), _
                                 StrPtr(addressBuffer), Len(addressBuffer))
    If convertedPointer = 0 Then
        EndpointText = "[IPv6 address conversion failed]"
    Else
        nullPosition = InStr(addressBuffer, vbNullChar)
        If nullPosition > 0 Then addressBuffer = Left$(addressBuffer, nullPosition - 1)
        EndpointText = "[" & addressBuffer & _
                       "]:" & CStr(UnsignedInteger(ntohs(value.sin6_port)))
    End If
End Function

'受信した応答が、送信した要求と同じTraceIdを返しているか確認します。
'相関IDを確認しないと、別要求の応答を誤って成功扱いする可能性があります。
Private Sub ValidateResponseTraceId( _
    ByVal responseText As String, _
    ByVal expectedTraceId As String)

    Dim fields As Variant

    fields = Split(TrimLineEnding(responseText), vbTab, 3)
    If UBound(fields) < 1 Then
        Err.Raise vbObjectError + 2225, "ValidateResponseTraceId", _
                  "サーバー応答の形式が不正です。"
    End If

    If StrComp(CStr(fields(0)), expectedTraceId, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 2226, "ValidateResponseTraceId", _
                  "サーバー応答のTraceIdが要求と一致しません。"
    End If
End Sub

'サーバー側ログの相関IDとして採用できる文字だけに制限します。
'ローカルの別プロセスから不正な式や改行を送り、ログを混乱させることを防ぎます。
Private Function IsSafeTraceId(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String

    If Len(value) = 0 Or Len(value) > 128 Then Exit Function

    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If Not ((character >= "0" And character <= "9") Or _
                (character >= "A" And character <= "Z") Or _
                (character >= "a" And character <= "z") Or _
                character = "-") Then Exit Function
    Next index

    IsSafeTraceId = True
End Function

'WSADATAは32bit/64bitで並びが異なります。サイズが違う状態でWSAStartupを
'呼ぶと、VBAが確保した範囲外へ書き込まれる危険があるため、先に停止します。
Private Sub ValidateWinsockStructureSizes()
    Dim winsockData As WSADATA
    Dim ipv6Address As SOCKADDR_IN6
    Dim expectedWsaDataBytes As Long

#If Win64 Then
    expectedWsaDataBytes = 408
#Else
    expectedWsaDataBytes = 400
#End If

    If LenB(winsockData) <> expectedWsaDataBytes Then
        Err.Raise vbObjectError + 2227, "ValidateWinsockStructureSizes", _
                  "WSADATAのサイズが想定と異なります。actual=" & _
                  CStr(LenB(winsockData)) & ", expected=" & _
                  CStr(expectedWsaDataBytes)
    End If

    If LenB(ipv6Address) <> 28 Then
        Err.Raise vbObjectError + 2228, "ValidateWinsockStructureSizes", _
                  "SOCKADDR_IN6のサイズが想定と異なります。actual=" & _
                  CStr(LenB(ipv6Address)) & ", expected=28"
    End If
End Sub

'ブック名にアポストロフィが含まれてもApplication.Run/OnTimeで正しく参照します。
Private Function WorkbookMacroReference( _
    ByVal workbookName As String, _
    ByVal macroName As String) As String

    WorkbookMacroReference = "'" & Replace(workbookName, "'", "''") & _
                             "'!" & macroName
End Function

Private Function TrimLineEnding(ByVal value As String) As String
    Do While Len(value) > 0 And _
             (Right$(value, 1) = vbCr Or Right$(value, 1) = vbLf)
        value = Left$(value, Len(value) - 1)
    Loop
    TrimLineEnding = value
End Function

'VBAのIntegerは符号付きですが、Winsockのポート番号は16bit符号なしです。
'32768以上をビット列を保ったままIntegerへ格納します。
Private Function ToSignedInteger(ByVal unsignedValue As Long) As Integer
    If unsignedValue <= 32767 Then
        ToSignedInteger = CInt(unsignedValue)
    Else
        ToSignedInteger = CInt(unsignedValue - 65536)
    End If
End Function

Private Function UnsignedInteger(ByVal signedValue As Integer) As Long
    UnsignedInteger = signedValue And &HFFFF&
End Function

Private Function MakeWord(ByVal lowByte As Byte, ByVal highByte As Byte) As Integer
    MakeWord = lowByte Or (CLng(highByte) * 256&)
End Function

Private Sub RaiseLastSocketError(ByVal message As String)
    Dim errorCode As Long

    errorCode = WSAGetLastError()
    Err.Raise vbObjectError + 2299, "Winsock", _
              message & " Winsockエラー=" & CStr(errorCode)
End Sub
