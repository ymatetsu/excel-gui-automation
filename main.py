from excel_utils import extract_table_from_excel
from gui_box import gui_loop

if __name__ == "__main__":
    file_path = "data/sample.xlsx"
    id_text1 = "ID:VREWで画像を何枚も自動生成したい場合"
    id_text2 = "ID:テーブル1"
    start_text = "特記事項"

    table1 = extract_table_from_excel(file_path, id_text1, start_text)
    table2 = extract_table_from_excel(file_path, id_text2, start_text)

    # GUIの実行開始だけ呼ぶ（入力ダイアログ含め gui_box に任せる）
    gui_loop(table1, table2, id_text2)
