# -*- coding: utf-8 -*-
"""Excelの「詳細データ」シートをTSVにするWindows DnD画面です。"""

import os
import subprocess
import sys
from typing import List

import win32api
import win32con
import win32gui


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
WINDOW_TITLE = "Group Shipment Summary (Drag & Drop)"


def show_message_box(pszMessage: str, pszTitle: str = WINDOW_TITLE) -> None:
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION)


def show_error_message_box(pszMessage: str, pszTitle: str = WINDOW_TITLE) -> None:
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def create_output_path(pszInputExcelPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputExcelPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszInputExcelPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszBaseName + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszBaseName}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def validate_dropped_file(pszInputExcelPath: str) -> str:
    if not os.path.isfile(pszInputExcelPath):
        return "Error: dropped item is not a file. Path = " + pszInputExcelPath
    if os.path.splitext(pszInputExcelPath)[1].lower() not in SUPPORTED_EXTENSIONS:
        return (
            "Error: unsupported Excel extension. Path = "
            + pszInputExcelPath
            + "\n\nSupported extensions: .xlsx, .xlsm"
        )
    return ""


def extract_error_file_path(pszStdErr: str) -> str:
    for pszLine in pszStdErr.splitlines():
        if pszLine.startswith("Error file: "):
            return pszLine[len("Error file: ") :].strip()
    return ""


def extract_warning_file_path(pszStdOut: str) -> str:
    for pszLine in pszStdOut.splitlines():
        if pszLine.startswith("警告ファイル: "):
            return pszLine[len("警告ファイル: ") :].strip()
    return ""


def run_excel_to_tsv(pszInputExcelPath: str) -> None:
    pszScriptPath: str = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "group_shipment_summary_Cmd.py"
    )
    if not os.path.isfile(pszScriptPath):
        show_error_message_box(
            "Error: group_shipment_summary_Cmd.py not found. Path = " + pszScriptPath
        )
        return

    pszValidationError: str = validate_dropped_file(pszInputExcelPath)
    if pszValidationError:
        show_error_message_box(pszValidationError)
        return

    pszOutputTsvPath: str = create_output_path(pszInputExcelPath)
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [sys.executable, pszScriptPath, pszOutputTsvPath, pszInputExcelPath],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as objException:
        show_error_message_box(
            "Error: failed to run converter. Detail = " + str(objException)
        )
        return

    if objCompletedProcess.returncode != 0 or not os.path.isfile(pszOutputTsvPath):
        pszStdErr: str = objCompletedProcess.stderr.strip() or "No stderr output."
        pszErrorFilePath: str = extract_error_file_path(pszStdErr)
        pszMessage: str = (
            "ExcelからTSVへの変換、年月別分割、グループ別集計・品種別集計・総計集計、または"
            "最新5年のグループ別・総計43項目抽出、総計入雛羽数・平均体重結合、総計出荷羽数・出荷重量結合、総計羽数15項目の過去5年表作成とstep0005縦結合、総計廃棄羽数・重量各15項目結合と小計追加、懸鳥前羽数・重量各2項目結合と小計追加、step0006廃棄データ一覧表上段・下段作成、step0007同表上段・下段の合計列追加、step0008総計廃棄データ一覧表作成、step0009総計廃棄データ一覧表加工、step0010総計廃棄データ一覧表月別連結、グループ別4項目結合、または"
            "グループ別年間連結、年間TSVの行列入れ替え、または"
            "年別グループ連結・step0008年間Excel作成・年間合計追加・step0009年間Excel作成・年間合計列追加・step0010年間Excel作成・step0013グループ別控除廃棄年間Excel作成に"
            "失敗しました。\n\nReturn code = "
            + str(objCompletedProcess.returncode)
            + "\n\nstderr:\n"
            + pszStdErr
        )
        if pszErrorFilePath:
            pszMessage += "\n\nエラー詳細:\n" + pszErrorFilePath
        show_error_message_box(pszMessage)
        return

    pszStdOut: str = objCompletedProcess.stdout.strip()
    pszWarningFilePath: str = extract_warning_file_path(pszStdOut)
    pszMessage = (
        (
            "処理が警告付きで完了しました。"
            if pszWarningFilePath
            else "処理が完了しました。"
        )
        + "\n\n入力Excel:\n"
        + pszInputExcelPath
        + "\n\n出力フォルダ:\n"
        + os.path.dirname(pszOutputTsvPath)
        + "\n\n元TSV:\n"
        + pszOutputTsvPath
    )
    if pszWarningFilePath:
        pszMessage += "\n\n警告ファイル:\n" + pszWarningFilePath
    show_message_box(pszMessage)


def draw_instruction_text(iWindowHandle: int) -> None:
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    objClientRect = win32gui.GetClientRect(iWindowHandle)
    iMargin: int = 8
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )
    pszInstructionText: str = (
        "Excelファイルをこのウィンドウに\n"
        "ドラッグ＆ドロップしてください。\n\n"
        "Excelの「詳細データ」シートをTSVへ変換します。\n"
        "処理対象の年月は入力データに含まれる年月から判定します。\n"
        "入力データを年月別のstep0001 TSVに分割します。\n"
        "年月別ファイルは入力データ内にある年月分を作成します。\n"
        "各月について、step0002のグループ別・品種別・総計TSVを作成します。\n"
        "step0001とstep0002の各ファイルは入力Excelと同じフォルダーに作成します。\n"
        "品種別集計は入力のP列（品種No.）と\n"
        "Q列（品種名）を使って行います。\n"
        "品種No.と品種名の組み合わせは次の4種類です。\n"
        "品種No.と品種名は両方とも一致する必要があります。\n"
        "どちらか一方だけが一致する場合もエラーになります。\n"
        "品種No. 5：阿波尾鶏\n"
        "品種No. 6：赤どり\n"
        "品種No. 2：地養鳥\n"
        "品種No. 1：ブロイラー\n"
        "品種別TSVは上記の順で4行を出力します。\n"
        "該当する明細がない品種も行を出力します。\n"
        "明細がない品種の合計値は0を出力します。\n"
        "明細がない品種の平均体重と日齢は空欄です。\n"
        "品種別集計ではグループを区別せずに合算します。\n"
        "品種別TSVのK列とL列には「－」を出力します。\n"
        "品種別集計ではK列・L列を使わず、グループ別集計では対応を検査します。\n"
        "品種別TSVのA列に年、B列に月を出力します。\n"
        "品種別TSVのD列に平均体重を出力します。\n"
        "品種別TSVのP列に品種No.、Q列に品種名を出力します。\n"
        "E列・F列は対象明細の値を合計します。\n"
        "O列には処理年月をYYYYMM形式で出力します。\n"
        "R列は入雛羽数、S列は出荷羽数を合計します。\n"
        "T列には日齢の単純平均を出力します。\n"
        "U列からBL列までは対象明細の値を合計します。\n"
        "上記に記載していない列には全角の「－」を出力します。\n"
        "平均体重は入力D列の平均ではありません。\n"
        "出荷重量合計を出荷羽数合計で割って再計算します。\n"
        "出荷羽数合計が0の場合、平均体重は空欄にします。\n"
        "出荷羽数がすべて空欄の場合も平均体重は空欄です。\n"
        "日齢は数値がある明細の単純平均を計算します。\n"
        "合計対象列は空欄を除いて合計します。\n"
        "合計対象列がすべて空欄の場合は0を出力します。\n"
        "合計値は不要な小数末尾のゼロを除いて出力します。\n"
        "平均体重は小数点以下5桁で四捨五入します。\n"
        "日齢は小数点以下2桁で四捨五入します。\n"
        "数値計算にはDecimalを使用します。\n"
        "出力TSVのヘッダーはstep0001のヘッダーを引き継ぎます。\n"
        "ヘッダーと明細の列数は64列である必要があります。\n"
        "入力に明細行がない場合はエラーになります。\n"
        "1つのstep0001ファイルに複数年月がある場合はエラーです。\n"
        "年・月または処理日付が不正な場合はエラーです。\n"
        "A列・B列の年月とO列の日付の年月が違う場合はエラーです。\n"
        "品種No.と品種名が指定の対応に一致しない場合はエラーです。\n"
        "数値列または日齢に空欄以外の非数値があればエラーです。\n"
        "合計対象列がすべて空欄の場合は警告を記録します。\n"
        "日齢がすべて空欄の場合は警告を記録します。\n"
        "出荷羽数がすべて空欄の場合は警告を記録します。\n"
        "出荷羽数合計が0の場合は警告を記録します。\n"
        "品種の明細がない場合は警告を記録します。\n"
        "警告がある場合、入力Excelと同じフォルダーに\n"
        "警告内容を記載したwarningファイルを作成します。\n"
        "step0002品種別ファイル名には「品種別」を含めます。\n"
        "品種別ファイル名には対象の年と月を含めます。\n"
        "品種別ファイル名の例：詳細データA_step0002_品種別_2026_07.tsv\n"
        "step0002総計ファイルも月ごとに作成します。\n"
        "step0002総計ファイル名には対象の年と月を含めます。\n"
        "品種別・総計の集計対象はそれぞれ同じ月の明細です。\n"
        "同名ファイルがある場合は上書きせず枝番を付けます。\n"
        "TSVはUTF-8（BOMなし）、タブ区切り、CRLF改行です。\n"
        "品種別TSVは追加出力であり、後続のグループ別処理には渡しません。\n"
        "グループ別・総計の後続処理は従来どおり実行します。\n"
        "最新年を含む直近5年間の\n"
        "指定43項目をグループ別・総計の項目別TSVへ抽出し、\n"
        "総計の入雛羽数と平均体重を横結合し、\n"
        "総計の出荷羽数と出荷重量を横結合し、\n"
        "総計の羽数15項目を項目別に直近5年間の月別表へまとめた\n"
        "step0004過去05年TSVを出力します。\n"
        "step0004の15ファイルを指定順に縦結合し、項目間に空白行を入れた\n"
        "step0005総計出荷廃棄羽数過去05年TSVも出力します。\n"
        "step0005総計出荷廃棄羽数過去05年TSVを読み込んだ\n"
        "同名のExcelファイルも出力します。\n"
        "総計の廃棄羽数・重量各15項目を横結合し、\n"
        "各グループの廃棄羽数・廃棄重量各15項目を横結合したTSVも出力します。\n"
        "各グループの廃棄羽数15項目に小計列を追加した\n"
        "step0005グループ別TSVも出力します。\n"
        "各グループの廃棄重量15項目に小計列を追加した\n"
        "step0005グループ別TSVも出力します。\n"
        "step0005グループ別廃棄重量TSVの3グループを縦結合した\n"
        "step0006グループ別廃棄重量TSVを出力します。\n"
        "step0006グループ別廃棄重量TSVから全廃棄重量を抽出した\n"
        "step0007グループ別TSVを出力します。\n"
        "step0005グループ別廃棄羽数TSVの3グループを縦結合した\n"
        "step0006グループ別廃棄羽数TSVを出力します。\n"
        "step0006グループ別廃棄羽数TSVから全廃棄羽数を抽出した\n"
        "step0007グループ別TSVを出力します。\n"
        "step0004グループ別データにstep0007の全廃棄羽数・重量を加えた\n"
        "step0008グループ別控除廃棄TSVを出力します。\n"
        "出力列の8列目は出荷重量です。\n"
        "step0008の全廃棄羽数・重量を出荷分から差し引き、\n"
        "処理羽数・処理重量と平均体重を計算した\n"
        "年・月を含むstep0009グループ別控除廃棄TSVを出力します。\n"
        "step0009契約・直営・フリー各グループの月別行を月順に連結した\n"
        "step0010グループ別控除廃棄年間TSVを出力します。\n"
        "step0010年間表を行列入れ替えしたstep0011グループ別控除廃棄TSVを出力します。\n"
        "step0011の契約・直営・フリーを年別にまとめたstep0012グループ別控除廃棄年間TSVを出力します。\n"
        "育成率を追加したstep0012からstep0013グループ別控除廃棄年間TSVを出力します。\n"
        "step0013グループ別控除廃棄年間TSVを読み込み、同じファイル名のExcelファイルを出力します。\n"
        "総計の懸鳥前羽数・重量各2項目を横結合し、\n"
        "懸鳥前・廃棄データ別の羽数・重量TSVに小計列を追加します。\n"
        "入雛・懸鳥前・廃棄羽数データを横結合した\n"
        "廃棄データ一覧表の上段を出力します。\n"
        "出荷重量・懸鳥前・廃棄重量データを横結合した\n"
        "廃棄データ一覧表の下段を出力します。\n"
        "上段・下段それぞれの小計2列を合算し、\n"
        "step0007廃棄データ一覧表に合計列を追加します。\n"
        "step0007の上段・下段を1つにまとめた\n"
        "step0008総計廃棄データ一覧表を出力します。\n"
        "step0008総計一覧表からグループ列を除き、\n"
        "羽数・重量の区分列を追加したstep0009を出力します。\n"
        "step0009の月別一覧表を年月の新しい順にまとめた\n"
        "step0010総計廃棄データ一覧表を出力し、\n"
        "1行目のH列・X列・Y列の見出しを空白にします。\n"
        "step0010表の右端に前年同月比(羽数)列を追加した\n"
        "step0011総計廃棄データ一覧表を出力します。\n"
        "前年同月の羽数合計との比率をZ列に追加した\n"
        "step0012総計廃棄データ一覧表を出力します。\n"
        "主要4項目をグループ別TSVへ結合し、\n"
        "年・グループ別の年間TSVへ連結し、\n"
        "年間TSVの行と列を入れ替え、\n"
        "契約・直営・フリーを年別に連結します。\n"
        "育成率追加後の年間TSVからExcelを作成し、\n"
        "育成率追加後の年間TSVへ合計を追加します。\n"
        "合計追加後の年間TSVからExcelを作成します。\n"
        "年間TSVの右端に年間合計列を追加します。\n"
        "年間合計列付きTSVからExcelを作成します。\n"
        "その他のシートは処理しません。\n\n"
        "1回につき1ファイルだけ処理できます。\n"
        "出力TSVは入力Excelと同じフォルダに作成します。\n\n"
        "対応形式: .xlsx / .xlsm"
    )
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objClientRect,
        win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK,
    )
    win32gui.EndPaint(iWindowHandle, objPaintStruct)


def window_proc(iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0
    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        try:
            iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)
            arrInputPaths: List[str] = [
                os.path.abspath(win32api.DragQueryFile(iDropHandle, iIndex))
                for iIndex in range(iFileCount)
            ]
        finally:
            win32api.DragFinish(iDropHandle)

        if not arrInputPaths:
            show_error_message_box("Error: no Excel file was dropped.")
            return 0
        if len(arrInputPaths) != 1:
            show_error_message_box(
                "Excelファイルは1回につき1ファイルだけ\n"
                "ドラッグ＆ドロップしてください。\n\n"
                "ドロップされたファイル数: " + str(len(arrInputPaths))
            )
            return 0
        run_excel_to_tsv(arrInputPaths[0])
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def main() -> None:
    pszWindowClassName: str = "GroupShipmentSummaryDndWindowClass"
    iInstanceHandle: int = win32api.GetModuleHandle(None)
    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1
    win32gui.RegisterClass(objWndClass)

    iWindowHandle: int = win32gui.CreateWindowEx(
        win32con.WS_EX_ACCEPTFILES,
        pszWindowClassName,
        WINDOW_TITLE,
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        520,
        320,
        0,
        0,
        iInstanceHandle,
        None,
    )
    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)
    win32gui.SetWindowPos(
        iWindowHandle,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
    )
    win32gui.PumpMessages()


if __name__ == "__main__":
    main()
