# DesktopGPT Launcher

## 概要
このプロジェクトは、Windows 環境で動作するコンソールアプリケーションです。ホットキーやショートカットで呼び出し、あらかじめ設定したプロンプトを ChatGPT (Web版またはデスクトップ版) に送信し、ユーザーの入力に応じて PC 操作を自動化します。

## 前提条件
- .NET SDK (最新LTS 推奨)
- VS Code または Visual Studio
- NuGet パッケージ管理 (dotnet CLI)

## ビルド
1. このリポジトリをクローン
   ```bash
   git clone https://github.com/yourname/MyChatGPTLauncher.git
   cd MyChatGPTLauncher
   ```

2. 必要なパッケージを追加
   ```bash
   dotnet add package Selenium.WebDriver
   dotnet add package Microsoft.Playwraight

## 設定

* `appsettings.json` にて以下の項目を定義できます。
  * `ShortcutKey` : アプリ起動用ショートカット
  * `InitialPrompt` : 起動時に入力される初期プロンプト