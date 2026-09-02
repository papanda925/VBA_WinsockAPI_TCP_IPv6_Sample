# 動作確認手順

## 事前確認

- Windows版Excel 2010以降を使用している
- ブックを`.xlsm`または`.xlsb`で保存した
- `src`の3モジュールをインポートした
- 「デバッグ」→「VBAProjectのコンパイル」が成功した
- 接続先が`::1`、ポートが`60052`のままである
- `CheckTcpIPv6StructureSizes`を実行し、`SOCKADDR_IN6=28 bytes`と表示された

## 正常系

1. クライアント側で`ClearTraceLog`を実行する。
2. `StartTcpIPv6ServerInNewExcel`を実行する。
3. サーバー側Traceに`SERVER_READY`が表示されることを確認する。
4. クライアント側で`SendHello`を実行する。
5. HELLO応答が表示されることを確認する。
6. `SendEchoSample`を実行する。
7. 日本語が文字化けせず返信されることを確認する。
8. 両方のTraceで同じTraceIdを検索する。
9. `StopTcpIPv6Server`を実行する。
10. サーバー側Traceに`SHUTDOWN`が表示されることを確認する。
11. `CloseTcpIPv6ServerExcel`を実行し、専用Excelが終了することを確認する。

## 異常系

### サーバーを起動せず送信

- `SendHello`を実行する。
- `CONNECT`の後に`ERROR`が記録されること。
- Excelが異常終了せず、ソケットとWinsockが後始末されること。

### ポート競合

- サーバーを2回起動する。
- 2個目のサーバーが`bind`エラーを表示すること。
- 1個目のサーバーは継続して利用できること。

### サーバー応答停止

- 通信中にサーバー側Excelを閉じる。
- クライアントが無期限に待たず、エラーをTraceへ記録すること。

## 確認結果の記録欄

| 項目 | 環境 | 結果 | 備考 |
| --- | --- | --- | --- |
| VBAコンパイル | Office 32bit | 未確認 | 実機確認待ち |
| VBAコンパイル | Office 64bit | 未確認 | 実機確認待ち |
| HELLO | Windows / Excel | 未確認 | 実機確認待ち |
| 日本語ECHO | Windows / Excel | 未確認 | 実機確認待ち |
| QUIT | Windows / Excel | 未確認 | 実機確認待ち |
