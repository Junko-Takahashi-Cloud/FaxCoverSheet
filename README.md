# FaxCoverSheet（C# × Excel Interop FAX送信票自動生成ツール）

このリポジトリは、C#（Windows Forms）と  
**Excel Interop（Microsoft.Office.Interop.Excel）** を使用して作成した  
**FAX送信票（カバーシート）自動生成ツール**です。

内部プロジェクト名（namespace）は **FAX送信票1** です。

Excel の元データを読み込み、企業ごとにデータを抽出・加工し、  
FAX 原紙テンプレート（Excel）へ自動転記して  
**企業別の FAX 送信票を自動生成**します。

---

## ✨ 主な機能

- Excel（元データ.xlsx）の読み込み  
- DataGridView へのデータ表示  
- Col5 以降の列を企業名として扱い、縦持ちデータへ変換  
- 日付（Excel シリアル値）を「日」だけに変換  
- 曜日も OADate → 文字列へ変換  
- 企業ごとにデータをグループ化  
- FAX 原紙テンプレート（FAX原紙.xlsx）へ自動転記  
- A7 / B5 に企業名を自動入力  
- C21〜 に日付、E列に曜日、G列に名前を順次書き込み  
- B19 に元データの値を転記  
- `FAX_企業名.xlsx` として自動保存

---

## 🛠 使用技術

- C# / .NET Framework  
- Windows Forms  
- Microsoft.Office.Interop.Excel  
- DataTable / LINQ  
- Excel テンプレートファイル（FAX原紙.xlsx）

---

## 📁 プロジェクト構成（想定）
/FaxCoverSheet
├─ FAX送信票1.sln
├─ /FAX送信票1
│   ├─ Form1.cs
│   ├─ Program.cs
│   ├─ その他のクラス
│   ├─ 元データ.xlsx
│   └─ FAX原紙.xlsx
└─ README.md

※ 現在は ZIP 形式でアップロードされています。  
　後で展開して配置するとコードが GitHub 上で閲覧できるようになります。

---

## 🚀 使い方

1. アプリを起動  
2. Excel の元データ（元データ.xlsx）を読み込み  
3. 「抽出・生成」ボタンを押す  
4. 企業ごとに FAX 原紙が自動生成される  
5. `FAX_企業名.xlsx` が出力される

---

## 📦 必要環境

- Windows 10 / 11  
- Microsoft Excel（2016 以降推奨）  
- .NET Framework  
- Excel Interop が利用可能な環境

---

## 🔧 今後の改善予定（任意）

- テンプレートの複数選択  
- PDF 直接出力（Excel 依存なし）  
- UI の改善  
- 入力履歴の保存  
- 出力先フォルダの選択機能

---

## 📜 ライセンス

個人利用を目的としています。  
必要に応じて MIT License などを追加してください。
