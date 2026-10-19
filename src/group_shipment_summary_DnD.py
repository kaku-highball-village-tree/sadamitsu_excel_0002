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
            "ExcelからTSVへの変換、年月別分割、または"
            "step0002品種別・総計ファイルの作成に失敗しました。"
            "\n\nReturn code = "
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
        "各月について、step0002の品種別TSVと総計TSVを作成します。\n"
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
        "入力のK列・L列のグループ対応は検査しません。\n"
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
        "step0002の作成後に処理を終了します。\n"
        "このプログラムではstep0003以降の処理を行いません。\n"
        "後続処理を品種別に行う場合は別途仕様が必要です。\n"
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
