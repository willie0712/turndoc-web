import os
import sys
import uuid
import json
import subprocess
import shutil
import re
import tempfile
import glob
from pathlib import Path

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from PIL import Image

# =========================================================
# HEIC 支援
# =========================================================

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

# =========================================================
# PDF → Word
# =========================================================

try:
    from pdf2docx import Converter
except ImportError:
    Converter = None


class TurnDocConverter:

    def __init__(self, output_dir="outputs"):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # -------------------------------------------------
        # 轉換引擎
        # -------------------------------------------------

        # Microsoft Office：只在 Windows 使用 COM
        self.office_available = self._check_office()

        # LibreOffice：跨平台備援（Windows / Linux / Render / Docker）
        self.libreoffice_path = self._find_libreoffice_path()
        self.libreoffice_available = self.libreoffice_path is not None

        print(f"[TurnDoc] Microsoft Office: {self.office_available}")
        print(f"[TurnDoc] LibreOffice: {self.libreoffice_available}")
        print(f"[TurnDoc] LibreOffice path: {self.libreoffice_path}")

        if not self.office_available and not self.libreoffice_available:
            print(
                "[TurnDoc] 警告：目前環境沒有 Microsoft Office 也沒有 LibreOffice，"
                "Word/PPT/Excel → PDF 轉換將無法使用。"
            )

    # =========================================================
    # 基本工具
    # =========================================================

    def _get_output_path(self, input_path, suffix):
        input_path = Path(input_path)
        return self.output_dir / f"{input_path.stem}.{suffix}"

    def _sanitize_filename(self, filename):
        filename = os.path.basename(str(filename))

        # 只處理真正的 \uXXXX escape
        if re.search(r"\\u[0-9a-fA-F]{4}", filename):
            try:
                filename = filename.encode("utf-8").decode("unicode_escape")
            except Exception:
                pass

        filename = re.sub(r'[<>:"/\\|?*]', "_", filename)
        filename = filename.strip()

        if not filename:
            filename = f"file_{uuid.uuid4().hex[:8]}"

        return filename

    def get_engine_status(self):
        """方便網頁前端顯示目前可用引擎"""
        return {
            "microsoft_office": self.office_available,
            "libreoffice": self.libreoffice_available,
            "libreoffice_path": self.libreoffice_path,
        }

    # =========================================================
    # Microsoft Office
    # =========================================================

    def _check_office(self):
        """
        Microsoft Office 只在 Windows 使用 COM。
        Linux / Render 一律回傳 False，避免嘗試啟動。
        """
        if os.name != "nt":
            return False

        try:
            import win32com.client
            import pythoncom
        except ImportError:
            print("[TurnDoc] Office detection: win32com 未安裝")
            return False

        word = None
        try:
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            return True

        except Exception as error:
            print(f"[TurnDoc] Office detection failed: {error}")
            return False

        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # =========================================================
    # LibreOffice
    # =========================================================

    def _find_libreoffice_path(self):
        """
        依序尋找：
        1. PATH 裡的 soffice / libreoffice
        2. Windows 常見安裝位置
        3. Linux / Render / Docker 常見位置
        4. Snap / Flatpak / 自訂路徑
        """
        candidates = []

        # 1. PATH
        for command in ("soffice", "libreoffice"):
            found = shutil.which(command)
            if found:
                candidates.append(found)

        # 2. Windows
        if os.name == "nt":
            candidates.extend([
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                r"C:\Program Files\LibreOffice*\program\soffice.exe",
            ])

        # 3. Linux / Render / Docker 常見位置
        candidates.extend([
            "/usr/bin/soffice",
            "/usr/bin/libreoffice",
            "/usr/local/bin/soffice",
            "/usr/local/bin/libreoffice",
            "/opt/libreoffice/program/soffice",
            "/opt/libreoffice/program/soffice.bin",
            "/opt/libreoffice*/program/soffice",
            "/usr/lib/libreoffice/program/soffice",
            "/usr/lib/libreoffice/program/soffice.bin",
            "/usr/lib/libreoffice/program/soffice.com",
            "/snap/bin/libreoffice",
            "/var/lib/flatpak/exports/bin/org.libreoffice.LibreOffice",
            "/app/libreoffice/program/soffice",          # 某些 Docker image
            "/opt/libreoffice25.2/program/soffice",     # 版本目錄常見寫法
            "/opt/libreoffice24.8/program/soffice",
            "/opt/libreoffice24.2/program/soffice",
        ])

        checked = set()

        for candidate in candidates:
            if not candidate:
                continue

            # 處理含 * 的路徑（Windows / Linux 都支援）
            if "*" in candidate:
                matches = glob.glob(candidate)
                for m in matches:
                    if m not in checked:
                        checked.add(m)
                        path = self._test_libreoffice_binary(m)
                        if path:
                            return path
                continue

            candidate = str(candidate)
            if candidate in checked:
                continue
            checked.add(candidate)

            path = self._test_libreoffice_binary(candidate)
            if path:
                return path

        return None

    def _test_libreoffice_binary(self, candidate):
        """實際測試一個候選路徑是否可用"""
        try:
            # 如果是檔案
            if os.path.isfile(candidate):
                if not os.access(candidate, os.X_OK):
                    return None

                result = subprocess.run(
                    [candidate, "--version"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    timeout=10,
                )
                if result.returncode == 0:
                    return candidate

            # 如果是指令名稱
            else:
                found = shutil.which(candidate)
                if found:
                    return found

        except Exception as error:
            print(f"[TurnDoc] LibreOffice check failed ({candidate}): {error}")

        return None

    def _check_libreoffice(self):
        return self._find_libreoffice_path() is not None

    # =========================================================
    # LibreOffice 轉換
    # =========================================================

    def _convert_with_libreoffice(self, input_path, output_format):
        if not self.libreoffice_path:
            raise RuntimeError("目前伺服器沒有 LibreOffice。")

        input_path = Path(input_path)
        temp_dir = Path(tempfile.mkdtemp(prefix="turndoc_", dir=str(self.output_dir)))

        # 每個轉換使用獨立 UserInstallation，避免多請求衝突
        user_install = temp_dir / "lo_user"
        user_install.mkdir(exist_ok=True)

        try:
            filters = {
                "pdf": "pdf:writer_pdf_Export",
                "docx": "docx:Office Open XML Text",
                "pptx": "pptx:Impress MS PowerPoint 2007 XML",
                "xlsx": "xlsx:Calc MS Excel 2007 XML",
            }

            convert_filter = filters.get(output_format, output_format)

            command = [
                self.libreoffice_path,
                "--headless",
                "--nologo",
                "--nodefault",
                "--nofirststartwizard",
                "--norestore",
                f"-env:UserInstallation=file://{user_install.as_posix()}",
                "--convert-to",
                convert_filter,
                "--outdir",
                str(temp_dir),
                str(input_path),
            ]

            print("[TurnDoc] LibreOffice command:", " ".join(command))

            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=180,
            )

            stdout = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""

            print(f"[TurnDoc] LibreOffice stdout: {stdout}")
            if stderr:
                print(f"[TurnDoc] LibreOffice stderr: {stderr}")

            if result.returncode != 0:
                message = stderr or stdout or "LibreOffice 執行失敗。"
                raise RuntimeError(f"LibreOffice 轉換失敗：\n{message}")

            # 找輸出檔
            expected = temp_dir / f"{input_path.stem}.{output_format}"
            if expected.exists():
                generated_file = expected
            else:
                matches = list(temp_dir.glob(f"{input_path.stem}.*"))
                # 排除 user profile 目錄
                matches = [m for m in matches if m.is_file()]
                if not matches:
                    raise RuntimeError("LibreOffice 已執行，但找不到輸出檔案。")
                generated_file = matches[0]

            final_path = self.output_dir / f"{input_path.stem}.{output_format}"
            if final_path.exists():
                final_path.unlink()

            shutil.move(str(generated_file), str(final_path))

            if not final_path.exists():
                raise RuntimeError("輸出檔案建立失敗。")

            return str(final_path)

        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    # =========================================================
    # Microsoft Office 轉換
    # =========================================================

    def _convert_with_office(self, input_path, output_format):
        if not self.office_available:
            raise RuntimeError("Microsoft Office 不可用。")

        import win32com.client
        import pythoncom

        input_path = Path(input_path).resolve()
        output_path = self._get_output_path(input_path, output_format)

        if output_path.exists():
            output_path.unlink()

        pythoncom.CoInitialize()

        word = None
        powerpoint = None
        excel = None

        try:
            extension = input_path.suffix.lower()

            # Word → PDF
            if extension in (".doc", ".docx") and output_format == "pdf":
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0

                document = word.Documents.Open(str(input_path))
                try:
                    document.ExportAsFixedFormat(str(output_path), 17)  # 17 = PDF
                finally:
                    document.Close(False)

            # PowerPoint → PDF
            elif extension in (".ppt", ".pptx") and output_format == "pdf":
                powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
                presentation = powerpoint.Presentations.Open(
                    str(input_path), WithWindow=False
                )
                try:
                    presentation.SaveAs(str(output_path), 32)  # 32 = PDF
                finally:
                    presentation.Close()

            # Excel → PDF
            elif extension in (".xls", ".xlsx") and output_format == "pdf":
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                workbook = excel.Workbooks.Open(str(input_path))
                try:
                    workbook.ExportAsFixedFormat(0, str(output_path))  # 0 = PDF
                finally:
                    workbook.Close(False)

            else:
                raise RuntimeError("目前 Microsoft Office 不支援這個轉換組合。")

            if not output_path.exists():
                raise RuntimeError("Microsoft Office 沒有產生輸出檔案。")

            return str(output_path)

        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            if powerpoint is not None:
                try:
                    powerpoint.Quit()
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    # =========================================================
    # 通用文件轉換（優先 Office → LibreOffice）
    # =========================================================

    def _convert_file(self, input_path, output_format):
        input_path = Path(input_path)

        # 1. Microsoft Office 優先（僅 Windows）
        if os.name == "nt" and self.office_available:
            try:
                print("[TurnDoc] Using Microsoft Office...")
                return self._convert_with_office(input_path, output_format)
            except Exception as error:
                print(f"[TurnDoc] Microsoft Office failed: {error}")

        # 2. LibreOffice 備援
        if self.libreoffice_available:
            try:
                print("[TurnDoc] Using LibreOffice...")
                return self._convert_with_libreoffice(input_path, output_format)
            except Exception as error:
                print(f"[TurnDoc] LibreOffice failed: {error}")
                raise RuntimeError(f"文件轉換失敗。\n\n{error}")

        # 3. 都沒有
        raise RuntimeError(
            "找不到可用的轉換引擎。\n\n"
            "目前環境沒有 Microsoft Office 或 LibreOffice。\n"
            "請確認伺服器已安裝其中一個。"
        )

    # =========================================================
    # Word → PDF
    # =========================================================

    def word_to_pdf(self, input_path):
        return self._convert_file(input_path, "pdf")

    # =========================================================
    # PDF → Word
    # =========================================================

    def pdf_to_word(self, input_path):
        if Converter is None:
            raise RuntimeError(
                "找不到 pdf2docx 套件，請確認 requirements.txt。"
            )

        input_path = Path(input_path)
        output_path = self._get_output_path(input_path, "docx")

        if output_path.exists():
            output_path.unlink()

        converter = None
        try:
            converter = Converter(str(input_path))
            converter.convert(str(output_path))
        finally:
            if converter:
                try:
                    converter.close()
                except Exception:
                    pass

        if not output_path.exists():
            raise RuntimeError("PDF → Word 轉換失敗。")

        return str(output_path)

    # =========================================================
    # PowerPoint → PDF
    # =========================================================

    def ppt_to_pdf(self, input_path):
        return self._convert_file(input_path, "pdf")

    # =========================================================
    # Excel → PDF
    # =========================================================

    def excel_to_pdf(self, input_path):
        return self._convert_file(input_path, "pdf")

    # =========================================================
    # PDF → PowerPoint
    # =========================================================

    def pdf_to_ppt(self, input_path):
        input_path = Path(input_path)
        output_path = self._get_output_path(input_path, "pptx")

        try:
            from pptx import Presentation
            from pptx.util import Inches
        except ImportError:
            raise RuntimeError(
                "缺少 python-pptx，請確認 requirements.txt。"
            )

        images = self._pdf_to_pil_images(input_path)
        if not images:
            raise RuntimeError("PDF 沒有可轉換的頁面。")

        presentation = Presentation()
        presentation.slide_width = Inches(13.333)
        presentation.slide_height = Inches(7.5)

        try:
            for image in images:
                temp_image = self.output_dir / f"ppt_page_{uuid.uuid4().hex}.png"
                try:
                    image.save(temp_image, "PNG")
                    slide = presentation.slides.add_slide(
                        presentation.slide_layouts[6]
                    )
                    slide.shapes.add_picture(
                        str(temp_image),
                        0, 0,
                        width=presentation.slide_width,
                        height=presentation.slide_height,
                    )
                finally:
                    try:
                        temp_image.unlink()
                    except Exception:
                        pass

            presentation.save(str(output_path))
        finally:
            for image in images:
                try:
                    image.close()
                except Exception:
                    pass

        return str(output_path)

    # =========================================================
    # PDF → Excel
    # =========================================================

    def pdf_to_excel(self, input_path):
        try:
            import pdfplumber
        except ImportError:
            raise RuntimeError("缺少 pdfplumber，請確認 requirements.txt。")

        try:
            from openpyxl import Workbook
        except ImportError:
            raise RuntimeError("缺少 openpyxl，請確認 requirements.txt。")

        input_path = Path(input_path)
        output_path = self._get_output_path(input_path, "xlsx")

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "PDF"
        row_number = 1

        with pdfplumber.open(str(input_path)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            if row:
                                for column_number, value in enumerate(row, start=1):
                                    worksheet.cell(
                                        row=row_number,
                                        column=column_number,
                                        value=value
                                    )
                                row_number += 1
                        row_number += 1
                else:
                    text = page.extract_text()
                    if text:
                        for line in text.splitlines():
                            worksheet.cell(row=row_number, column=1, value=line)
                            row_number += 1

        workbook.save(str(output_path))
        return str(output_path)

    # =========================================================
    # PDF 合併
    # =========================================================

    def merge_pdfs(self, input_paths):
        if not input_paths:
            raise RuntimeError("沒有 PDF 可以合併。")

        output_path = self.output_dir / f"merged_{uuid.uuid4().hex[:8]}.pdf"
        merger = PdfMerger()

        try:
            for path in input_paths:
                merger.append(str(path))
            merger.write(str(output_path))
        finally:
            merger.close()

        return str(output_path)

    # =========================================================
    # PDF 壓縮
    # =========================================================

    def compress_pdf(self, input_path, quality="medium"):
        input_path = Path(input_path)
        output_path = self.output_dir / f"{input_path.stem}.compressed.pdf"

        ghostscript = self._find_ghostscript()
        if ghostscript:
            quality_map = {
                "low": "/screen",
                "medium": "/ebook",
                "high": "/printer",
            }
            pdf_quality = quality_map.get(quality, "/ebook")

            command = [
                ghostscript,
                "-sDEVICE=pdfwrite",
                "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={pdf_quality}",
                "-dNOPAUSE",
                "-dQUIET",
                "-dBATCH",
                f"-sOutputFile={output_path}",
                str(input_path),
            ]

            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=180,
            )

            if result.returncode == 0 and output_path.exists():
                return str(output_path)

        # 沒有 Ghostscript → 使用 PyPDF2 重寫
        reader = PdfReader(str(input_path))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        with open(output_path, "wb") as file:
            writer.write(file)

        return str(output_path)

    def _find_ghostscript(self):
        commands = ["gs", "gswin64c", "gswin32c"]
        for command in commands:
            found = shutil.which(command)
            if found:
                return found

        windows_paths = [
            r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
            r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
        ]
        for pattern in windows_paths:
            matches = glob.glob(pattern)
            if matches:
                return matches[-1]

        return None

    # =========================================================
    # Poppler
    # =========================================================

    def _find_poppler_path(self):
        if os.name != "nt":
            # Linux 直接使用 PATH
            if shutil.which("pdftoppm"):
                return None
            return None

        # Windows
        base_dir = Path(__file__).resolve().parent
        candidates = [
            base_dir / "poppler" / "bin",
            base_dir.parent / "poppler" / "bin",
            Path.cwd() / "poppler" / "bin",
            Path(r"C:\poppler\bin"),
        ]

        for path in candidates:
            if (path / "pdftoppm.exe").exists():
                return str(path)

        return None

    # =========================================================
    # PDF → PIL Image
    # =========================================================

    def _pdf_to_pil_images(self, input_path):
        try:
            from pdf2image import convert_from_path
        except ImportError:
            raise RuntimeError("缺少 pdf2image，請確認 requirements.txt。")

        poppler_path = self._find_poppler_path()

        try:
            if poppler_path:
                return convert_from_path(str(input_path), poppler_path=poppler_path)
            return convert_from_path(str(input_path))
        except Exception as error:
            raise RuntimeError(
                "PDF → 圖片失敗。\n\n系統需要 Poppler / pdftoppm。\n"
                f"{error}"
            )

    # =========================================================
    # PDF → 圖片
    # =========================================================

    def pdf_to_images(self, input_path, output_format="png"):
        input_path = Path(input_path)
        images = self._pdf_to_pil_images(input_path)
        output_files = []

        output_format = output_format.lower()
        if output_format not in ("png", "jpg", "jpeg", "webp"):
            output_format = "png"

        for index, image in enumerate(images, start=1):
            extension = "jpg" if output_format == "jpeg" else output_format
            output_path = self.output_dir / f"{input_path.stem}_{index}.{extension}"

            if extension in ("jpg", "jpeg"):
                if image.mode in ("RGBA", "LA", "P"):
                    image = image.convert("RGB")
                image.save(output_path, "JPEG", quality=95)
            else:
                image.save(output_path, output_format.upper())

            output_files.append(str(output_path))

        return output_files

    # =========================================================
    # 圖片 → PDF
    # =========================================================

    def images_to_pdf(self, input_paths):
        if not input_paths:
            raise RuntimeError("沒有圖片可以轉換。")

        output_path = self.output_dir / f"images_{uuid.uuid4().hex[:8]}.pdf"
        images = []

        try:
            for path in input_paths:
                image = Image.open(path)
                if image.mode != "RGB":
                    image = image.convert("RGB")
                images.append(image)

            if not images:
                raise RuntimeError("沒有有效圖片。")

            first = images[0]
            remaining = images[1:]
            first.save(
                output_path,
                "PDF",
                resolution=100.0,
                save_all=True,
                append_images=remaining,
            )
        finally:
            for image in images:
                try:
                    image.close()
                except Exception:
                    pass

        return str(output_path)

    # =========================================================
    # 圖片格式轉換
    # =========================================================

    def convert_image(self, input_path, output_format):
        input_path = Path(input_path)
        output_format = output_format.lower()
        if output_format == "jpeg":
            output_format = "jpg"

        supported = ["png", "jpg", "jpeg", "webp", "bmp", "tiff", "gif"]
        if output_format not in supported:
            raise RuntimeError(f"不支援的圖片格式：{output_format}")

        output_path = self._get_output_path(input_path, output_format)
        image = Image.open(input_path)

        try:
            if output_format in ("jpg", "jpeg"):
                if image.mode in ("RGBA", "LA", "P"):
                    image = image.convert("RGB")
                image.save(output_path, "JPEG", quality=95)
            else:
                image.save(output_path, output_format.upper())
        finally:
            image.close()

        return str(output_path)

    # =========================================================
    # 圖片壓縮
    # =========================================================

    def compress_image(self, input_path, quality="medium", output_format=None):
        input_path = Path(input_path)

        quality_map = {"low": 45, "medium": 70, "high": 85}
        image_quality = quality_map.get(quality, 70)

        if not output_format or output_format == "original":
            extension = input_path.suffix.lower().lstrip(".")
            if extension == "jpeg":
                extension = "jpg"
        else:
            extension = output_format.lower()

        if extension == "jpeg":
            extension = "jpg"

        output_path = self.output_dir / f"{input_path.stem}.compressed.{extension}"
        image = Image.open(input_path)

        try:
            if extension in ("jpg", "jpeg"):
                if image.mode in ("RGBA", "LA", "P"):
                    image = image.convert("RGB")
                image.save(output_path, "JPEG", quality=image_quality, optimize=True)
            elif extension == "webp":
                image.save(output_path, "WEBP", quality=image_quality, method=6)
            elif extension == "png":
                image.save(output_path, "PNG", optimize=True)
            else:
                image.save(output_path)
        finally:
            image.close()

        return str(output_path)

    # =========================================================
    # GIF 分割
    # =========================================================

    def split_gif(self, input_path, output_format="png"):
        input_path = Path(input_path)
        image = Image.open(input_path)
        output_files = []

        try:
            frame_count = getattr(image, "n_frames", 1)
            output_format = output_format.lower()
            if output_format == "jpeg":
                output_format = "jpg"

            for index in range(frame_count):
                image.seek(index)
                frame = image.convert("RGBA")
                extension = output_format

                output_path = (
                    self.output_dir
                    / f"{input_path.stem}_frame_{index + 1}.{extension}"
                )

                try:
                    if extension in ("jpg", "jpeg"):
                        frame = frame.convert("RGB")
                        frame.save(output_path, "JPEG", quality=95)
                    else:
                        frame.save(output_path, output_format.upper())
                    output_files.append(str(output_path))
                finally:
                    frame.close()
        finally:
            image.close()

        return output_files