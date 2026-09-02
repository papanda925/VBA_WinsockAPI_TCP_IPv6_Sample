# Excel VBA Winsock TCP/IPv6 Client–Server Sample

Excel VBAからWindowsのWinsock APIを直接呼び出し、IPv6のTCPクライアントと簡易サーバーをローカルPC内で動かす学習用サンプルです。

`WScript.Shell`、PowerShell、`curl`、外部ActiveXコントロールは使用しません。クライアントとサーバーの処理を`Trace`シートで追跡できるため、正常時だけでなく「どこまで処理が進んだか」も確認できます。

> [!IMPORTANT]
> このサンプルはネットワーク学習用です。認証、暗号化、複数クライアントの並行処理、大容量データ転送は実装していません。初めは接続先をIPv6ループバックアドレス`::1`から変更しないでください。

## 学べること

- IPv4の`127.0.0.1`に相当するIPv6ループバックアドレス`::1`
- `SOCKADDR_IN6`に含まれるIPv6アドレス、ポート、スコープID
- `WSAStartup` / `socket` / `bind` / `listen` / `accept`
- `connect` / `send` / `recv` / `closesocket` / `WSACleanup`
- `InetPtonW`と`InetNtopW`によるIPv6アドレス変換
- UTF-16とUTF-8の変換
- TCPでは1回の`send`で全データを送れるとは限らないこと
- ブロッキングAPIを別Excelプロセスへ分離する理由
- TraceIdを使ってクライアントとサーバーの記録を対応付ける方法

## ファイル構成

| ファイル | 内容 |
| --- | --- |
| `src/VBA_WinsockAPI_TCP_IPv6_Sample.bas` | IPv6 TCPサーバー、クライアント、実行マクロ |
| `src/TraceLogger.bas` | Traceシートとイミディエイトウィンドウへの記録 |
| `src/Utf8Codec.bas` | VBA文字列とUTF-8バイト列の相互変換 |
| `docs/TESTING.md` | 動作確認項目と期待結果 |
| `docs/REVIEW.md` | 10ペルソナ・100観点レビューの記録 |
| `SECURITY.md` | 安全上の注意と脆弱性の連絡方法 |

## 動作環境

- Windows 10またはWindows 11
- Excel 2010以降（VBA 7）
- 32ビット版または64ビット版Office
- IPv6が有効なWindows環境
- `.xlsm`または`.xlsb`形式のマクロ有効ブック

macOS版ExcelではWindows DLLを呼び出せないため動作しません。

## 使い方

### 1. モジュールを取り込む

1. Excelでマクロ有効ブックを作成し、一度保存します。
2. `Alt` + `F11`でVisual Basic Editorを開きます。
3. 「ファイル」→「ファイルのインポート」から、`src`内の3ファイルを取り込みます。
4. 「デバッグ」→「VBAProjectのコンパイル」を実行します。
5. Excelへ戻り、マクロを有効にします。

GitHub上の`.bas`はUTF-8です。VBEで日本語が文字化けする場合は、UTF-8対応エディターでCP932（Shift_JIS）へ変換してからインポートするか、VBEの新規モジュールへコードを貼り付けてください。日本語文字列もあるため、コンパイル前に表示を確認します。

TCP版とUDP版には同名の初心者向けマクロがあります。両リポジトリのモジュールを同じブックへ混在させず、それぞれ別の学習用ブックへ取り込んでください。

### 2. サーバーを起動する

`StartTcpIPv6ServerInNewExcel`を実行します。

同じブックが別のExcelプロセスで読み取り専用として開き、1秒後に`::1:60052`で接続を待ちます。サーバー側Excelの`Trace`シートに`SERVER_READY`が表示されるまで待ってください。

### 3. クライアントから送信する

最初のExcelへ戻り、次のマクロを実行します。

| マクロ | 内容 | サーバーの応答 |
| --- | --- | --- |
| `SendHello` | 接続確認 | 固定のHELLOメッセージ |
| `SendEchoSample` | 日本語をUTF-8で送信 | 受信内容をそのまま返信 |
| `StopTcpIPv6Server` | 終了要求 | 応答後にサーバーを停止 |
| `CloseTcpIPv6ServerExcel` | 別Excelの終了 | 読み取り専用ブックを保存せず閉じる |

`StopTcpIPv6Server`で受信ループを終了し、サーバー側Traceの`SHUTDOWN`を確認してから`CloseTcpIPv6ServerExcel`を実行します。起動元が保持している参照を解放し、サーバー専用Excelを終了します。

## 通信の流れ

### サーバー

```text
WSAStartup
  → socket(AF_INET6, SOCK_STREAM)
  → IPV6_V6ONLY
  → bind([::1]:60052)
  → listen
  → accept
  → recv
  → send
  → closesocket
  → WSACleanup
```

### クライアント

```text
WSAStartup
  → socket(AF_INET6, SOCK_STREAM)
  → connect([::1]:60052)
  → send
  → recv
  → closesocket
  → WSACleanup
```

## Traceシート

| 列 | 意味 |
| --- | --- |
| `Time` | ミリ秒付きの記録時刻 |
| `TraceId` | 1回の要求をクライアント・サーバー間で結ぶ識別子 |
| `Side` | `CLIENT`または`SERVER` |
| `Step` | `CONNECT`、`REQUEST_SENT`、`RESPONSE_RECEIVED`など |
| `Direction` | `IN`、`OUT`、`LOCAL` |
| `Detail` | 接続先、コマンド、処理内容 |
| `Result` | `START`、`OK`、`SENT`、`FAILED`など |
| `ElapsedMs` | 操作開始からの経過ミリ秒 |

クライアントが作成したTraceIdを要求電文へ含め、サーバーも同じTraceIdで記録します。これにより、両方のExcelで同じ通信を探せます。

## TCPの重要な制限

TCPは「メッセージ」ではなくバイトストリームです。このサンプルでは、1回の`recv`で全体が届くと決めつけず、LF終端文字を受け取るまで`recv`を繰り返します。ただし、ローカルPC内の短い1行を対象とした教材であり、実運用では次も必要です。

- バイナリデータにも対応できる長さヘッダーなどのフレーミング
- 最大サイズ、読取途中、複数メッセージ同時到着への対応
- 切断、部分送信、再接続の設計
- TLSなどによる暗号化と相手の認証
- 複数接続の並行処理

## トラブルシューティング

### 接続できない

サーバー側のTraceに`SERVER_READY`があることを確認してください。先にクライアントを実行すると接続エラーになります。

### `Address already in use`に相当するエラー

以前起動したサーバー側Excelが残っていないか確認してください。既定ポート`60052`を他のアプリが使用している場合は、クライアントとサーバーの`DEFAULT_SERVER_PORT`を同じ未使用ポートへ変更します。

### IPv6アドレスが変換できない

Windowsのネットワーク設定でIPv6が無効化されていないか確認してください。最初のテストでは`::1`を使用します。

### Excelが待機しているように見える

サーバー側Excelの`accept`は接続待ちを行います。これは同期APIの通常動作です。操作と送信は最初のクライアント側Excelで行ってください。

### 別Excelでマクロを実行できない

組織のポリシーやトラストセンター設定により、読み取り専用で自動オープンしたブックのマクロが禁止される場合があります。セキュリティ設定を回避せず、信頼できる場所と管理者の方針を確認してください。

## 参考資料

- [IPv6 guide for Windows Sockets applications](https://learn.microsoft.com/windows/win32/winsock/ip-version-6-2)
- [SOCKADDR_IN6 structure](https://learn.microsoft.com/windows/win32/api/ws2ipdef/ns-ws2ipdef-sockaddr_in6_lh)
- [InetPtonW function](https://learn.microsoft.com/windows/win32/api/ws2tcpip/nf-ws2tcpip-inetptonw)
- [socket function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-socket)
- [send function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-send)
- [recv function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-recv)
- [64-bit VBA overview](https://learn.microsoft.com/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview)

## 実機確認について

コードとAPI宣言は静的に確認していますが、この作成環境にはWindows版Excelがありません。`docs/TESTING.md`に従い、32ビット版または64ビット版Excelでコンパイルとループバック通信を確認してください。
