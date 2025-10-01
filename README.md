# Excel × GUI 自動化ツール

## 概要
Excelに記載した操作手順を読み取り、マウスクリック・文字入力・Enterキー操作などを自動で実行するPythonツールです。  
GUI上で「停止・一時停止・再開・スキップ」を操作でき、Excelの定義と連動した柔軟な自動化を実現します。  

本ツールは、画像生成AIの「待ち時間」を効率化したいという身近な課題から着想を得て制作しました。  
業務改善や自動化の学習成果を形にしたポートフォリオとして公開しています。  

---

## ファイル構成
- `main.py` : エントリーポイント（Excel読み込み → GUI起動）
- `gui_box.py` : GUIと自動処理
- `excel_utils.py` : Excelからテーブルを抽出する関数
- `sound_utils.py` : 効果音再生（終了・停止時）

---

## サンプルExcel
`sample.xlsx` は動作確認用のサンプルです。  

---

## 必要なライブラリ
```bash
pip install openpyxl pyautogui pyperclip
