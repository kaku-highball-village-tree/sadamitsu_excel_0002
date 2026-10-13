# -*- coding: utf-8 -*-
"""複数画像を1つのPDFにするWindowsドラッグ＆ドロップ画面です。"""

import os
import subprocess
import sys
from datetime import datetime
from typing import List

import win32api
import win32con
import win32gui


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}
WINDOW_TITLE = "ImagesToPdf (Drag & Drop)"


def show_message_box(pszMessage: str, pszTitle: str = WINDOW_TITLE) -> None:
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONINFORMATION)


def show_error_message_box(pszMessage: str, pszTitle: str = WINDOW_TITLE) -> None:
    win32gui.MessageBox(0, pszMessage, pszTitle, win32con.MB_OK | win32con.MB_ICONERROR)


def create_output_path(pszFirstImagePath: str) -> str:
    pszDirectory: str = os.path.dirname(pszFirstImagePath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszFirstImagePath))[0]
    pszDateTime: str = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    pszStem: str = f"ImagesToPdf_{pszDateTime}_{pszBaseName}"
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".pdf")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.pdf")
        iSuffix += 1
    return pszOutputPath


def validate_dropped_files(arrImagePaths: List[str]) -> str:
    for pszImagePath in arrImagePaths:
        if not os.path.isfile(pszImagePath):
            return "Error: dropped item is not a file. Path = " + pszImagePath
        if os.path.splitext(pszImagePath)[1].lower() not in SUPPORTED_EXTENSIONS:
            return (
                "Error: unsupported image extension. Path = "
                + pszImagePath
                + "\n\nSupported extensions: .jpg, .jpeg, .png, .bmp, .gif"
            )
    return ""


def run_images_to_pdf(arrImagePaths: List[str]) -> None:
    pszScriptPath: str = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "ImagesToPdf_Cmd.py"
    )
    if not os.path.isfile(pszScriptPath):
        show_error_message_box("Error: ImagesToPdf_Cmd.py not found. Path = " + pszScriptPath)
        return

    pszValidationError: str = validate_dropped_files(arrImagePaths)
    if pszValidationError:
        show_error_message_box(pszValidationError)
        return

    pszOutputPdfPath: str = create_output_path(arrImagePaths[0])
    try:
        objCompletedProcess: subprocess.CompletedProcess[str] = subprocess.run(
            [sys.executable, pszScriptPath, pszOutputPdfPath, *arrImagePaths],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as objException:
        show_error_message_box("Error: failed to run converter. Detail = " + str(objException))
        return

    if objCompletedProcess.returncode != 0 or not os.path.isfile(pszOutputPdfPath):
        pszStdErr: str = objCompletedProcess.stderr.strip() or "No stderr output."
        show_error_message_box(
            "画像からPDFへの変換に失敗しました。\n\nReturn code = "
            + str(objCompletedProcess.returncode)
            + "\n\nstderr:\n"
            + pszStdErr
        )
        return

    show_message_box(
        "画像をPDFへ変換しました。\n\n画像数: "
        + str(len(arrImagePaths))
        + "\n出力PDF:\n"
        + pszOutputPdfPath
    )


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
        "JPEG、PNG、BMP、GIF画像をこのウィンドウに\n"
        "ドラッグ＆ドロップしてください。\n\n"
        "複数の画像を、受け取った順番で1つのPDFにまとめます。\n"
        "出力PDFは先頭画像と同じフォルダに作成します。\n\n"
        "対応形式: .jpg / .jpeg / .png / .bmp / .gif"
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
            arrImagePaths: List[str] = [
                os.path.abspath(win32api.DragQueryFile(iDropHandle, iIndex))
                for iIndex in range(iFileCount)
            ]
        finally:
            win32api.DragFinish(iDropHandle)
        if not arrImagePaths:
            show_error_message_box("Error: no image files were dropped.")
            return 0
        run_images_to_pdf(arrImagePaths)
        return 0
    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0
    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0
    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def main() -> None:
    pszWindowClassName: str = "ImagesToPdfDndWindowClass"
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
        win32con.WS_OVERLAPPED | win32con.WS_CAPTION | win32con.WS_SYSMENU | win32con.WS_MINIMIZEBOX,
        win32con.CW_USEDEFAULT,
        win32con.CW_USEDEFAULT,
        520,
        280,
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
