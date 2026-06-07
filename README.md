# 📠 FaxCoverSheet – FAX送信票自動生成ツール（C# × Excel Interop）

このツールは、C#（Windows Forms）と  
**Microsoft.Office.Interop.Excel** を使用して作成した  
**FAX送信票（カバーシート）自動生成アプリ** です。

Excel の元データを読み込み、企業ごとにデータを抽出・加工し、  
FAX原紙テンプレートへ自動転記して **企業別のFAX送信票を自動生成** します。

---

## ✨ 主な機能
- Excel（元データ.xlsx）の読み込み  
- DataGridView へのデータ表示  
- Col5 以降の列を企業名として扱い、縦持ちデータへ変換  
- 日付（Excelシリアル値）を「日」だけに変換  
- 曜日を OADate → 文字列へ変換  
- 企業ごとにデータをグループ化  
- FAX原紙テンプレート（FAX原紙.xlsx）へ自動転記  
  - A7 / B5 に企業名を自動入力  
  - C21〜 に日付  
  - E列に曜日  
  - G列に名前  
  - B19 に元データの値を転記  
- `FAX_企業名.xlsx` として自動保存

---

## 🛠 使用技術
- C# / .NET Framework  
- Windows Forms  
- Microsoft.Office.Interop.Excel  
- DataTable / LINQ  
- Excel テンプレート（FAX原紙.xlsx）

---

## 📁 プロジェクト構成
/FaxCoverSheet
├─ FAX送信票1.sln
├─ /FAX送信票1
│   ├─ Form1.cs
│   ├─ Program.cs
│   ├─ その他のクラス
│   ├─ 元データ.xlsx
│   └─ FAX原紙.xlsx

---

## 💡 工夫した点
- Excel Interop を使った **実務レベルの自動化処理**  
- DataTable と LINQ によるデータ整形  
- テンプレートファイルへの自動転記処理  
- “現場で使えるツール” を意識した UI と処理フロー  

---

## 🚀 今後の改善予定
- エラー処理の強化  
- テンプレートの複数パターン対応  
- ログ出力機能  
- UI の改善（進捗表示など）

