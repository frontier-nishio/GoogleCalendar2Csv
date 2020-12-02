# GoogleCalendar2Csv

![License:Apache-2.0](https://img.shields.io/github/license/frontier-nishio/GoogleCalendar2Csv) ![Release](https://img.shields.io/github/downloads/frontier-nishio/GoogleCalendar2Csv/total)

## このプログラムについて

Googleカレンダーに登録してある内容をCSV形式で出力します。

## 必要なもの

* PowerShellが実行可能なWindows環境
* Googleカレンダーを参照可能なiCal形式のURL（非公開カレンダーの場合は、非公開用のもの）

## 使い方

1. 「GoogleCalendarDownload.ps1」をダウンロードします。
2. 「GoogleCalendarDownload.ps1」をメモ帳などで開きます。
3. 2行目の「(ここにURLを設定します)」をダウンロードしたいカレンダーのiCal形式のURLに変更します。
4. 保存します。
5. 「GoogleCalendarDownload.ps1」を右クリックして「PowerShell で実行」をクリックします。
6. 実行したフォルダに「20201202_030436_export_calendar.csv」といったファイル名のファイルが生成されます。

## 開発・テスト環境

* OS: Windows 10 Pro 64bit (Ver.20H2)
* Editor: Visual Studio Code (Ver.1.51.1)
