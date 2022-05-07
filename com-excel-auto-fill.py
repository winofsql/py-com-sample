# python -m pip install pywin32
import win32com.client
import traceback
import sys
import os

ExcelApp = win32com.client.Dispatch("Excel.Application")
# デバッグ時は、Excel の本体を表示させて状況が解るようにする
ExcelApp.Visible = True
# UI でチェックさせるようなダイアログを表示せずに実行する
ExcelApp.DisplayAlerts = False

try:
    # ****************************
    # ブック追加
    # ****************************
    Book = ExcelApp.Workbooks.Add()

    # 通常一つのシートが作成されています
    Sheet = Book.Worksheets( 1 )

    # ****************************
    # シート名変更
    # ****************************
    Sheet.Name = "Pythonの処理";

    # ****************************
    # セルに値を直接セット
    # ****************************
    for i in range(1, 11):
        Sheet.Cells(i, 1).Value = f"処理 : {i}"

    # ****************************
    # 1つのセルから
    # AutoFill で値をセット
    # ****************************
    Sheet.Cells(1, 2).Value = "子"
    # 基となるセル範囲
    SourceRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(1,2))
    # オートフィルの範囲(基となるセル範囲を含む )
    FillRange = Sheet.Range(Sheet.Cells(1, 2), Sheet.Cells(10,2))
    SourceRange.AutoFill(FillRange)

    # ****************************
    # 保存
    # ****************************
    Book.SaveAs( os.getcwd() + "\\sample.xlsx" )

except Exception:
    ExcelApp.Quit()
    traceback.print_exc()
    sys.exit( )


ExcelApp.Quit()
print("処理を終了します")