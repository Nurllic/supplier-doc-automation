#!/usr/bin/env python3
# -*- coding: utf-8 -*-
""".doc 老格式 → .docx 转换：Linux 容器用 LibreOffice headless，Windows 用 Word COM。"""

import os
import shutil
import subprocess


def convert_doc_to_docx(src, out_dir=None):
    """返回转换后的 .docx 路径；失败返回空字符串"""
    if not src or not os.path.exists(src) or os.path.getsize(src) <= 0:
        return ""
    out_dir = out_dir or os.path.dirname(src)
    dst = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + ".docx")
    if os.path.exists(dst):
        return dst
    if os.name != "nt":
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            return ""
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "docx", "--outdir", out_dir, src],
                capture_output=True, text=True, timeout=180,
            )
        except Exception:
            return ""
        return dst if os.path.exists(dst) else ""
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
        except Exception:
            pass
        d = word.Documents.Open(src, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)
        d.SaveAs(dst, FileFormat=12)
        d.Close(False)
        try:
            word.Quit()
        except Exception:
            pass
        return dst if os.path.exists(dst) else ""
    except Exception:
        return ""


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python doc_converter.py <文件.doc> [输出目录]")
        sys.exit(1)
    out = convert_doc_to_docx(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(out or "转换失败")
