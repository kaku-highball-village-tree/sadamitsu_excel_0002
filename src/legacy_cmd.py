# -*- coding: utf-8 -*-
"""複数の画像ファイルを、指定された順番で1つのPDFにまとめます。"""

import os
import sys
import tempfile
from typing import List

import fitz  # PyMuPDF


SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}


def validate_paths(pszOutputPdfPath: str, arrImagePaths: List[str]) -> None:
    if os.path.splitext(pszOutputPdfPath)[1].lower() != ".pdf":
        raise ValueError("output file extension must be .pdf.")

    pszOutputDirectory: str = os.path.dirname(pszOutputPdfPath)
    if not os.path.isdir(pszOutputDirectory):
        raise ValueError(
            "output directory was not found. Path = " + pszOutputDirectory
        )
    if os.path.isdir(pszOutputPdfPath):
        raise ValueError("output path points to a directory. Path = " + pszOutputPdfPath)

    pszNormalizedOutputPath: str = os.path.normcase(pszOutputPdfPath)
    for pszImagePath in arrImagePaths:
        if not os.path.isfile(pszImagePath):
            raise ValueError("input image file was not found. Path = " + pszImagePath)
        if os.path.splitext(pszImagePath)[1].lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                "unsupported image extension. Path = "
                + pszImagePath
                + "\nSupported extensions: .jpg, .jpeg, .png, .bmp, .gif"
            )
        if os.path.normcase(pszImagePath) == pszNormalizedOutputPath:
            raise ValueError("input and output paths must be different. Path = " + pszImagePath)


def convert_images_to_pdf(pszOutputPdfPath: str, arrImagePaths: List[str]) -> None:
    objOutputDocument: fitz.Document = fitz.open()
    pszTemporaryPdfPath: str = ""

    try:
        iTotalImages: int = len(arrImagePaths)
        for iImageIndex, pszImagePath in enumerate(arrImagePaths):
            objImageDocument: fitz.Document = fitz.open(pszImagePath)
            try:
                bImagePdf: bytes = objImageDocument.convert_to_pdf()
                objImagePdfDocument: fitz.Document = fitz.open("pdf", bImagePdf)
                try:
                    # GIFも1画像につき先頭の1ページだけを追加します。
                    objOutputDocument.insert_pdf(objImagePdfDocument, from_page=0, to_page=0)
                finally:
                    objImagePdfDocument.close()
            finally:
                objImageDocument.close()

            print(f"画像 {iImageIndex + 1}/{iTotalImages} → {pszImagePath}")

        iTemporaryFileDescriptor, pszTemporaryPdfPath = tempfile.mkstemp(
            prefix="ImagesToPdf_",
            suffix=".pdf",
            dir=os.path.dirname(pszOutputPdfPath),
        )
        os.close(iTemporaryFileDescriptor)
        objOutputDocument.save(pszTemporaryPdfPath)
        objOutputDocument.close()
        os.replace(pszTemporaryPdfPath, pszOutputPdfPath)
        pszTemporaryPdfPath = ""
    finally:
        if not objOutputDocument.is_closed:
            objOutputDocument.close()
        if pszTemporaryPdfPath and os.path.exists(pszTemporaryPdfPath):
            os.remove(pszTemporaryPdfPath)


def main() -> int:
    if len(sys.argv) < 3:
        print(
            "Usage: python ImagesToPdf_Cmd.py <output_pdf_path> "
            "<image_path_1> [image_path_2 ...]",
            file=sys.stderr,
        )
        return 1

    pszOutputPdfPath: str = os.path.abspath(sys.argv[1])
    arrImagePaths: List[str] = [os.path.abspath(pszPath) for pszPath in sys.argv[2:]]

    try:
        validate_paths(pszOutputPdfPath, arrImagePaths)
        convert_images_to_pdf(pszOutputPdfPath, arrImagePaths)
    except Exception as objException:
        print("Error: failed to create PDF. Detail = " + str(objException), file=sys.stderr)
        return 1

    print("◇ 変換完了 ◇")
    print("出力PDF: " + pszOutputPdfPath)
    return 0


if __name__ == "__main__":
    sys.exit(main())
