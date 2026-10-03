import os
import sys
import uuid
import json
import subprocess
import shutil
import re

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from PIL import Image


# ============================================================
# 支援 HEIC 格式
# ============================================================

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass


# ============================================================
# PDF 轉 Word
# ============================================================

try:
    from pdf2docx import Converter
except ImportError:
    pass


# ============================================================
# 強制 UTF-8 編碼
# ============================================================

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass


# ============================================================
# TurnDoc 轉換核心 - v1.0.3
# ============================================================

class TurnDocConverter:

    def __init__(self, output_dir='./outputs'):

        self.output_dir = os.path.abspath(
            output_dir
        )

        os.makedirs(
            self.output_dir,
            exist_ok=True
        )

        self.office_available = (
            self._check_office()
        )

        self.libreoffice_available = (
            self._check_libreoffice()
        )

        print(
            f'[INFO] Microsoft Office: '
            f'{"OK" if self.office_available else "NO"}',
            file=sys.stderr
        )

        print(
            f'[INFO] LibreOffice: '
            f'{"OK" if self.libreoffice_available else "NO"}',
            file=sys.stderr
        )

        if (
            not self.office_available
            and not self.libreoffice_available
        ):

            print(
                '[WARN] 找不到任何轉換引擎！',
                file=sys.stderr
            )

    # ============================================================
    # Microsoft Office
    # ============================================================

    def _check_office(self):

        try:

            import win32com.client
            import pythoncom

            pythoncom.CoInitialize()

            word = win32com.client.Dispatch(
                "Word.Application"
            )

            word.Quit()

            return True

        except Exception:

            return False

    # ============================================================
    # LibreOffice
    # ============================================================

    def _check_libreoffice(self):

        common_paths = [

            r'C:\Program Files\LibreOffice\program\soffice.exe',

            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'

        ]

        for path in common_paths:

            if os.path.exists(path):

                return True

        try:

            subprocess.run(
                ['soffice', '--version'],
                capture_output=True,
                check=True
            )

            return True

        except Exception:
            pass

        try:

            subprocess.run(
                ['libreoffice', '--version'],
                capture_output=True,
                check=True
            )

            return True

        except Exception:
            pass

        return False

    # ============================================================
    # LibreOffice 路徑
    # ============================================================

    def _find_libreoffice_path(self):

        common_paths = [

            r'C:\Program Files\LibreOffice\program\soffice.exe',

            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'

        ]

        for path in common_paths:

            if os.path.exists(path):

                return path

        return 'soffice'

    # ============================================================
    # 檔名安全處理
    # ============================================================

    def _sanitize_filename(
        self,
        filename
    ):

        if not filename:

            return 'TurnDoc_output'

        try:

            if (
                '\\u' in filename
                or 'u' in filename
            ):

                filename = re.sub(
                    r'(?<!\\)u([0-9a-fA-F]{4})',
                    r'\\u\1',
                    filename
                )

                filename = filename.encode(
                    'utf-8'
                ).decode(
                    'unicode_escape'
                )

        except Exception:

            pass

        filename = re.sub(
            r'[<>:"/\\|?*]',
            '',
            filename
        )

        filename = filename.strip(
            '. '
        )

        if not filename:

            filename = 'TurnDoc_output'

        return filename

    # ============================================================
    # 取得輸出路徑
    # ============================================================

    def _get_output_path(
        self,
        input_path,
        ext
    ):

        base_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        base_name = self._sanitize_filename(
            base_name
        )

        output_filename = (
            f'{base_name}.{ext}'
        )

        output_path = os.path.join(
            self.output_dir,
            output_filename
        )

        counter = 1

        while os.path.exists(
            output_path
        ):

            output_filename = (
                f'{base_name}_{counter}.{ext}'
            )

            output_path = os.path.join(
                self.output_dir,
                output_filename
            )

            counter += 1

        return output_path

    # ============================================================
    # Microsoft Office 轉換
    # ============================================================

    def _convert_with_office(
        self,
        input_path,
        output_format='pdf'
    ):

        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()

        ext = os.path.splitext(
            input_path
        )[1].lower()

        out_ext = output_format.lower()

        output_path = self._get_output_path(
            input_path,
            out_ext
        )

        try:

            # ----------------------------------------------------
            # Word
            # ----------------------------------------------------

            if ext in [
                '.docx',
                '.doc'
            ]:

                word = win32com.client.Dispatch(
                    "Word.Application"
                )

                word.Visible = False

                doc = word.Documents.Open(
                    os.path.abspath(
                        input_path
                    )
                )

                doc.SaveAs(
                    os.path.abspath(
                        output_path
                    ),
                    FileFormat=(
                        17
                        if out_ext == 'pdf'
                        else 16
                    )
                )

                doc.Close()

                word.Quit()

            # ----------------------------------------------------
            # Excel
            # ----------------------------------------------------

            elif ext in [
                '.xlsx',
                '.xls'
            ]:

                excel = win32com.client.Dispatch(
                    "Excel.Application"
                )

                excel.Visible = False

                wb = excel.Workbooks.Open(
                    os.path.abspath(
                        input_path
                    )
                )

                if out_ext == 'pdf':

                    wb.ExportAsFixedFormat(
                        0,
                        os.path.abspath(
                            output_path
                        )
                    )

                else:

                    wb.SaveAs(
                        os.path.abspath(
                            output_path
                        ),
                        FileFormat=51
                    )

                wb.Close()

                excel.Quit()

            # ----------------------------------------------------
            # PowerPoint
            # ----------------------------------------------------

            elif ext in [
                '.pptx',
                '.ppt'
            ]:

                powerpoint = (
                    win32com.client.Dispatch(
                        "PowerPoint.Application"
                    )
                )

                powerpoint.Visible = True

                ppt = (
                    powerpoint.Presentations.Open(
                        os.path.abspath(
                            input_path
                        ),
                        WithWindow=False
                    )
                )

                if out_ext == 'pdf':

                    ppt.SaveAs(
                        os.path.abspath(
                            output_path
                        ),
                        FileFormat=32
                    )

                else:

                    ppt.SaveAs(
                        os.path.abspath(
                            output_path
                        ),
                        FileFormat=24
                    )

                ppt.Close()

                powerpoint.Quit()

            else:

                raise Exception(
                    f'不支援的格式: {ext}'
                )

            return output_path

        except Exception as e:

            raise Exception(
                f'Office 轉換失敗: {str(e)}'
            )

    # ============================================================
    # LibreOffice 轉換
    # ============================================================

    def _convert_with_libreoffice(
        self,
        input_path,
        output_format='pdf'
    ):

        soffice_path = (
            self._find_libreoffice_path()
        )

        out_ext = (
            output_format.lower()
        )

        temp_dir = os.path.join(
            self.output_dir,
            'temp_libreoffice'
        )

        os.makedirs(
            temp_dir,
            exist_ok=True
        )

        cmd = [

            soffice_path,

            '--headless',

            '--convert-to',

            out_ext,

            '--outdir',

            os.path.abspath(
                temp_dir
            ),

            os.path.abspath(
                input_path
            )

        ]

        try:

            subprocess.run(
                cmd,
                check=True,
                capture_output=True,
                text=True,
                encoding='utf-8'
            )

        except subprocess.CalledProcessError as e:

            raise Exception(
                f'LibreOffice 轉換失敗: {e.stderr}'
            )

        base_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        generated = os.path.join(
            temp_dir,
            f'{base_name}.{out_ext}'
        )

        if not os.path.exists(
            generated
        ):

            files = [

                f

                for f in os.listdir(
                    temp_dir
                )

                if f.endswith(
                    f'.{out_ext}'
                )

            ]

            if files:

                generated = os.path.join(
                    temp_dir,
                    files[0]
                )

            else:

                raise Exception(
                    '找不到 LibreOffice 輸出檔案'
                )

        output_path = (
            self._get_output_path(
                input_path,
                out_ext
            )
        )

        shutil.copy2(
            generated,
            output_path
        )

        try:

            shutil.rmtree(
                temp_dir
            )

        except Exception:

            pass

        return output_path

    # ============================================================
    # 一般檔案轉換
    # ============================================================

    def _convert_file(
        self,
        input_path,
        output_format='pdf'
    ):

        if self.office_available:

            try:

                return self._convert_with_office(
                    input_path,
                    output_format
                )

            except Exception as e:

                print(
                    f'[WARN] Office 轉換失敗: {e}，'
                    f'嘗試切換 LibreOffice...',
                    file=sys.stderr
                )

        if self.libreoffice_available:

            return self._convert_with_libreoffice(
                input_path,
                output_format
            )

        raise Exception(
            '找不到可用的轉換引擎！'
            '請確認是否已安裝 Microsoft Office 或 LibreOffice。'
        )

    # ============================================================
    # Word → PDF
    # ============================================================

    def word_to_pdf(
        self,
        input_path
    ):

        return self._convert_file(
            input_path,
            'pdf'
        )

    # ============================================================
    # PDF → Word
    # ============================================================

    def pdf_to_word(
        self,
        input_path
    ):

        try:

            output_path = self._get_output_path(
                input_path,
                'docx'
            )

            cv = Converter(
                input_path
            )

            cv.convert(
                output_path,
                start=0,
                end=None
            )

            cv.close()

            return output_path

        except Exception as e:

            raise Exception(
                f'PDF 轉 Word 失敗: {str(e)}'
            )

    # ============================================================
    # PowerPoint → PDF
    # ============================================================

    def ppt_to_pdf(
        self,
        input_path
    ):

        return self._convert_file(
            input_path,
            'pdf'
        )

    # ============================================================
    # Excel → PDF
    # ============================================================

    def excel_to_pdf(
        self,
        input_path
    ):

        return self._convert_file(
            input_path,
            'pdf'
        )

    # ============================================================
    # PDF → PowerPoint
    # ============================================================

    def pdf_to_ppt(
        self,
        input_path
    ):

        return self._convert_file(
            input_path,
            'pptx'
        )

    # ============================================================
    # PDF → Excel
    # ============================================================

    def pdf_to_excel(
        self,
        input_path
    ):

        return self._convert_file(
            input_path,
            'xlsx'
        )

    # ============================================================
    # 合併 PDF
    # ============================================================

    def merge_pdfs(
        self,
        input_paths
    ):

        if not input_paths:

            raise Exception(
                '合併列表不能為空'
            )

        output_path = (
            self._get_output_path(
                input_paths[0],
                'pdf'
            )
        )

        if not output_path.endswith(
            '_merged.pdf'
        ):

            output_path = (
                output_path.replace(
                    '.pdf',
                    '_merged.pdf'
                )
            )

        merger = PdfMerger()

        try:

            for path in input_paths:

                merger.append(
                    path
                )

            merger.write(
                output_path
            )

        finally:

            merger.close()

        return output_path

    # ============================================================
    # PDF 壓縮
    #
    # 核心規則：
    #
    # 1. 嘗試 PyPDF2
    # 2. 嘗試 Ghostscript
    # 3. 比較所有結果
    # 4. 只有比原始檔更小才使用
    # 5. 如果全部都變大 → 保留原始 PDF
    # ============================================================

    def compress_pdf(
        self,
        input_path,
        quality=80
    ):

        input_path = os.path.abspath(
            input_path
        )

        if not os.path.exists(
            input_path
        ):

            raise Exception(
                f'找不到輸入 PDF: {input_path}'
            )

        os.makedirs(
            self.output_dir,
            exist_ok=True
        )

        base_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        base_name = self._sanitize_filename(
            base_name
        )

        # --------------------------------------------------------
        # 原始大小
        # --------------------------------------------------------

        original_size = os.path.getsize(
            input_path
        )

        # --------------------------------------------------------
        # 最終輸出檔
        # --------------------------------------------------------

        output_path = os.path.join(
            self.output_dir,
            f'{base_name}_compressed.pdf'
        )

        counter = 1

        while os.path.exists(
            output_path
        ):

            output_path = os.path.join(
                self.output_dir,
                f'{base_name}_compressed_{counter}.pdf'
            )

            counter += 1

        # --------------------------------------------------------
        # 暫存檔
        # --------------------------------------------------------

        temp_pypdf = os.path.join(
            self.output_dir,
            f'.{uuid.uuid4().hex}_pypdf.pdf'
        )

        temp_gs = os.path.join(
            self.output_dir,
            f'.{uuid.uuid4().hex}_gs.pdf'
        )

        candidates = []

        try:

            # ====================================================
            # 1. PyPDF2
            # ====================================================

            try:

                reader = PdfReader(
                    input_path
                )

                writer = PdfWriter()

                for page in reader.pages:

                    try:

                        page.compress_content_streams()

                    except Exception:

                        pass

                    writer.add_page(
                        page
                    )

                try:

                    writer._compress = True

                except Exception:

                    pass

                with open(
                    temp_pypdf,
                    'wb'
                ) as output_file:

                    writer.write(
                        output_file
                    )

                if os.path.exists(
                    temp_pypdf
                ):

                    pypdf_size = os.path.getsize(
                        temp_pypdf
                    )

                    if pypdf_size > 0:

                        candidates.append(
                            (
                                pypdf_size,
                                temp_pypdf,
                                'PyPDF2'
                            )
                        )

            except Exception as e:

                print(
                    f'[WARN] PyPDF2 壓縮失敗: {e}',
                    file=sys.stderr
                )

            # ====================================================
            # 2. Ghostscript
            # ====================================================

            try:

                quality = int(
                    quality
                )

            except Exception:

                quality = 80

            quality = max(
                1,
                min(
                    100,
                    quality
                )
            )

            gs_path = None

            possible_commands = [

                'gswin64c',

                'gswin32c',

                'gs'

            ]

            for command in possible_commands:

                found = shutil.which(
                    command
                )

                if found:

                    gs_path = found

                    break

            # ----------------------------------------------------
            # Windows Ghostscript 常見路徑
            # ----------------------------------------------------

            if (
                not gs_path
                and os.name == 'nt'
            ):

                import glob

                possible_patterns = [

                    r'C:\Program Files\gs\gs*\bin\gswin64c.exe',

                    r'C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe'

                ]

                for pattern in possible_patterns:

                    matches = glob.glob(
                        pattern
                    )

                    if matches:

                        matches.sort(
                            reverse=True
                        )

                        gs_path = matches[0]

                        break

            # ----------------------------------------------------
            # Ghostscript 找得到
            # ----------------------------------------------------

            if gs_path:

                if quality < 30:

                    pdf_settings = (
                        '/screen'
                    )

                elif quality < 60:

                    pdf_settings = (
                        '/ebook'
                    )

                else:

                    pdf_settings = (
                        '/prepress'
                    )

                gs_cmd = [

                    gs_path,

                    '-sDEVICE=pdfwrite',

                    '-dCompatibilityLevel=1.4',

                    f'-dPDFSETTINGS={pdf_settings}',

                    '-dNOPAUSE',

                    '-dQUIET',

                    '-dBATCH',

                    f'-sOutputFile={temp_gs}',

                    input_path

                ]

                result = subprocess.run(

                    gs_cmd,

                    check=False,

                    capture_output=True,

                    text=True,

                    encoding='utf-8',

                    errors='replace'

                )

                if (
                    result.returncode == 0
                    and os.path.exists(
                        temp_gs
                    )
                ):

                    gs_size = os.path.getsize(
                        temp_gs
                    )

                    if gs_size > 0:

                        candidates.append(
                            (
                                gs_size,
                                temp_gs,
                                'Ghostscript'
                            )
                        )

                else:

                    print(
                        '[WARN] Ghostscript 壓縮失敗',
                        file=sys.stderr
                    )

            else:

                print(
                    '[INFO] 找不到 Ghostscript，'
                    '只使用 PyPDF2。',
                    file=sys.stderr
                )

            # ====================================================
            # 3. 找出比原始檔更小的結果
            # ====================================================

            smaller_candidates = [

                candidate

                for candidate in candidates

                if candidate[0] < original_size

            ]

            # ====================================================
            # 4. 沒有任何版本更小
            # ====================================================

            if not smaller_candidates:

                shutil.copy2(
                    input_path,
                    output_path
                )

                print(
                    '[INFO] 壓縮後沒有更小，'
                    '保留原始 PDF。',
                    file=sys.stderr
                )

                return os.path.abspath(
                    output_path
                )

            # ====================================================
            # 5. 選最小的版本
            # ====================================================

            best_size, best_path, best_method = min(

                smaller_candidates,

                key=lambda x: x[0]

            )

            shutil.copy2(
                best_path,
                output_path
            )

            # ====================================================
            # 6. 顯示結果
            # ====================================================

            original_mb = (
                original_size
                / 1024
                / 1024
            )

            final_mb = (
                best_size
                / 1024
                / 1024
            )

            saved_percent = (

                (
                    original_size
                    - best_size
                )

                / original_size

                * 100

            )

            print(
                f'[INFO] PDF 壓縮完成 '
                f'({best_method})',
                file=sys.stderr
            )

            print(
                f'[INFO] 原始大小: '
                f'{original_mb:.2f} MB',
                file=sys.stderr
            )

            print(
                f'[INFO] 壓縮後: '
                f'{final_mb:.2f} MB',
                file=sys.stderr
            )

            print(
                f'[INFO] 減少: '
                f'{saved_percent:.1f}%',
                file=sys.stderr
            )

            return os.path.abspath(
                output_path
            )

        except Exception as e:

            if os.path.exists(
                output_path
            ):

                try:

                    os.remove(
                        output_path
                    )

                except Exception:

                    pass

            raise Exception(
                f'PDF 壓縮失敗: {str(e)}'
            )

        finally:

            # ----------------------------------------------------
            # 清理暫存檔
            # ----------------------------------------------------

            for temp_file in [

                temp_pypdf,

                temp_gs

            ]:

                try:

                    if os.path.exists(
                        temp_file
                    ):

                        os.remove(
                            temp_file
                        )

                except Exception:

                    pass

    # ============================================================
    # PDF → 圖片
    # ============================================================

    def pdf_to_images(
        self,
        input_path,
        output_format='PNG'
    ):

        from pdf2image import (
            convert_from_path
        )

        base_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        base_name = self._sanitize_filename(
            base_name
        )

        # --------------------------------------------------------
        # 找程式所在位置
        # --------------------------------------------------------

        if getattr(
            sys,
            'frozen',
            False
        ):

            base_dir = os.path.dirname(
                os.path.abspath(
                    sys.executable
                )
            )

        else:

            base_dir = os.path.dirname(
                os.path.abspath(
                    __file__
                )
            )

        # --------------------------------------------------------
        # Poppler 搜尋路徑
        # --------------------------------------------------------

        poppler_possible_paths = [

            # api/poppler/bin
            os.path.join(
                base_dir,
                'poppler',
                'bin'
            ),

            # 專案根目錄/poppler/bin
            os.path.join(
                os.path.dirname(
                    base_dir
                ),
                'poppler',
                'bin'
            ),

            # 目前工作目錄/poppler/bin
            os.path.join(
                os.getcwd(),
                'poppler',
                'bin'
            ),

            # 備用
            r'C:\poppler\bin'

        ]

        poppler_path = None

        for p in poppler_possible_paths:

            exe_path = os.path.join(
                p,
                'pdftoppm.exe'
            )

            if os.path.exists(
                exe_path
            ):

                poppler_path = p

                break

        fmt = output_format.upper()

        if fmt == 'JPG':

            fmt = 'JPEG'

        try:

            if poppler_path:

                images = convert_from_path(

                    input_path,

                    poppler_path=poppler_path

                )

            else:

                searched_str = '\n'.join(
                    poppler_possible_paths
                )

                raise Exception(

                    '找不到 poppler/bin/pdftoppm.exe\n'

                    f'已搜尋路徑:\n{searched_str}'

                )

        except Exception as e:

            raise Exception(
                f'無法讀取 PDF: {str(e)}'
            )

        output_paths = []

        for i, img in enumerate(
            images
        ):

            out_ext = (

                'jpg'

                if fmt == 'JPEG'

                else fmt.lower()

            )

            output_filename = (

                f'{base_name}_page_{i + 1}.{out_ext}'

            )

            output_path = os.path.join(

                self.output_dir,

                output_filename

            )

            img.save(
                output_path,
                fmt
            )

            output_paths.append(
                output_path
            )

        return output_paths

    # ============================================================
    # 圖片 → PDF
    # ============================================================

    def images_to_pdf(
        self,
        image_paths
    ):

        if not image_paths:

            raise Exception(
                '圖片列表不能為空'
            )

        output_path = (
            self._get_output_path(
                image_paths[0],
                'pdf'
            )
        )

        images = []

        for path in image_paths:

            img = Image.open(
                path
            )

            if img.mode in (
                'RGBA',
                'LA',
                'P'
            ):

                img = img.convert(
                    'RGB'
                )

            images.append(
                img
            )

        images[0].save(

            output_path,

            save_all=True,

            append_images=images[1:]

        )

        return output_path

    # ============================================================
    # 圖片轉換
    # ============================================================

    def convert_image(
        self,
        input_path,
        output_format='PNG'
    ):

        fmt = output_format.upper()

        if fmt == 'JPG':

            fmt = 'JPEG'

        out_ext = (

            'jpg'

            if fmt == 'JPEG'

            else fmt.lower()

        )

        output_path = (
            self._get_output_path(
                input_path,
                out_ext
            )
        )

        img = Image.open(
            input_path
        )

        if (
            fmt == 'JPEG'
            and img.mode in (
                'RGBA',
                'LA',
                'P'
            )
        ):

            background = Image.new(

                'RGB',

                img.size,

                (255, 255, 255)

            )

            if img.mode == 'P':

                img = img.convert(
                    'RGBA'
                )

            if img.mode == 'RGBA':

                background.paste(

                    img,

                    mask=img.split()[-1]

                )

            else:

                background.paste(
                    img
                )

            img = background

        img.save(

            output_path,

            format=fmt

        )

        return output_path

    # ============================================================
    # 圖片壓縮
    # ============================================================

    def compress_image(
        self,
        input_path,
        quality=80,
        output_format=None
    ):

        img = Image.open(
            input_path
        )

        orig_ext = os.path.splitext(
            input_path
        )[1].lower()

        if orig_ext in [
            '.heic',
            '.heif'
        ]:

            fmt = 'JPEG'

        else:

            fmt = (

                output_format.upper()

                if output_format

                else (

                    img.format

                    or 'JPEG'

                )

            )

            if fmt == 'JPG':

                fmt = 'JPEG'

        png_compress = max(

            0,

            min(

                9,

                int(
                    (100 - quality) / 11
                )

            )

        )

        out_ext = (

            'jpg'

            if fmt == 'JPEG'

            else fmt.lower()

        )

        output_path = (
            self._get_output_path(
                input_path,
                f'compressed.{out_ext}'
            )
        )

        if (
            fmt == 'JPEG'
            and img.mode in (
                'RGBA',
                'LA',
                'P'
            )
        ):

            background = Image.new(

                'RGB',

                img.size,

                (255, 255, 255)

            )

            if img.mode == 'P':

                img = img.convert(
                    'RGBA'
                )

            if img.mode == 'RGBA':

                background.paste(

                    img,

                    mask=img.split()[-1]

                )

            else:

                background.paste(
                    img
                )

            img = background

        if fmt == 'JPEG':

            img.save(

                output_path,

                format='JPEG',

                quality=quality,

                optimize=True

            )

        elif fmt == 'PNG':

            img.save(

                output_path,

                format='PNG',

                compress_level=png_compress,

                optimize=True

            )

        elif fmt == 'WEBP':

            img.save(

                output_path,

                format='WEBP',

                quality=quality

            )

        else:

            img.save(

                output_path,

                format=fmt,

                quality=quality

            )

        return output_path

    # ============================================================
    # GIF 分割
    # ============================================================

    def split_gif(
        self,
        input_path,
        output_format='PNG'
    ):

        from PIL import ImageSequence

        fmt = output_format.upper()

        if fmt == 'JPG':

            fmt = 'JPEG'

        out_ext = (

            'jpg'

            if fmt == 'JPEG'

            else fmt.lower()

        )

        base_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        base_name = self._sanitize_filename(
            base_name
        )

        img = Image.open(
            input_path
        )

        output_paths = []

        for i, frame in enumerate(
            ImageSequence.Iterator(img)
        ):

            output_path = os.path.join(

                self.output_dir,

                f'{base_name}_frame_{i + 1}.{out_ext}'

            )

            frame.save(

                output_path,

                format=fmt

            )

            output_paths.append(
                output_path
            )

        return output_paths


# ============================================================
# CLI 入口
# ============================================================

if __name__ == '__main__':

    import argparse

    parser = argparse.ArgumentParser()

    parser.add_argument(
        '--tool',
        required=True
    )

    parser.add_argument(
        '--input',
        nargs='+',
        required=True
    )

    parser.add_argument(
        '--output',
        default='./outputs'
    )

    parser.add_argument(
        '--format',
        default='PNG'
    )

    parser.add_argument(
        '--quality',
        type=int,
        default=80,
        help='壓縮品質 (1-100)'
    )

    args = parser.parse_args()

    converter = TurnDocConverter(
        output_dir=args.output
    )

    try:

        # --------------------------------------------------------
        # Word → PDF
        # --------------------------------------------------------

        if args.tool == 'word-to-pdf':

            out = converter.word_to_pdf(
                args.input[0]
            )

        # --------------------------------------------------------
        # PDF → Word
        # --------------------------------------------------------

        elif args.tool == 'pdf-to-word':

            out = converter.pdf_to_word(
                args.input[0]
            )

        # --------------------------------------------------------
        # PowerPoint → PDF
        # --------------------------------------------------------

        elif args.tool == 'ppt-to-pdf':

            out = converter.ppt_to_pdf(
                args.input[0]
            )

        # --------------------------------------------------------
        # Excel → PDF
        # --------------------------------------------------------

        elif args.tool == 'excel-to-pdf':

            out = converter.excel_to_pdf(
                args.input[0]
            )

        # --------------------------------------------------------
        # PDF → PowerPoint
        # --------------------------------------------------------

        elif args.tool == 'pdf-to-ppt':

            out = converter.pdf_to_ppt(
                args.input[0]
            )

        # --------------------------------------------------------
        # PDF → Excel
        # --------------------------------------------------------

        elif args.tool == 'pdf-to-excel':

            out = converter.pdf_to_excel(
                args.input[0]
            )

        # --------------------------------------------------------
        # 合併 PDF
        # --------------------------------------------------------

        elif args.tool == 'merge':

            out = converter.merge_pdfs(
                args.input
            )

        # --------------------------------------------------------
        # PDF 壓縮
        # --------------------------------------------------------

        elif args.tool == 'compress':

            out = converter.compress_pdf(

                args.input[0],

                args.quality

            )

        # --------------------------------------------------------
        # PDF → 圖片
        # --------------------------------------------------------

        elif args.tool == 'pdf-to-image':

            out = converter.pdf_to_images(

                args.input[0],

                args.format

            )

        # --------------------------------------------------------
        # 圖片 → PDF
        # --------------------------------------------------------

        elif args.tool == 'image-to-pdf':

            out = converter.images_to_pdf(
                args.input
            )

        # --------------------------------------------------------
        # 圖片轉換
        # --------------------------------------------------------

        elif args.tool == 'image-convert':

            out = converter.convert_image(

                args.input[0],

                args.format

            )

        # --------------------------------------------------------
        # 圖片壓縮
        # --------------------------------------------------------

        elif args.tool == 'image-compress':

            out = converter.compress_image(

                args.input[0],

                args.quality,

                (
                    args.format

                    if args.format != 'original'

                    else None
                )

            )

        # --------------------------------------------------------
        # GIF 分割
        # --------------------------------------------------------

        elif args.tool == 'gif-split':

            out = converter.split_gif(

                args.input[0],

                args.format

            )

        # --------------------------------------------------------
        # 不支援
        # --------------------------------------------------------

        else:

            print(

                json.dumps(

                    {

                        'error':
                        f'不支援的工具: {args.tool}'

                    },

                    ensure_ascii=False

                ),

                file=sys.stderr

            )

            sys.exit(1)

        # --------------------------------------------------------
        # 成功輸出
        # --------------------------------------------------------

        print(

            json.dumps(

                out

                if isinstance(
                    out,
                    list
                )

                else [
                    out
                ],

                ensure_ascii=False

            ),

            end=''

        )

    except Exception as e:

        error_msg = str(e)

        print(

            json.dumps(

                {

                    'error':
                    error_msg

                },

                ensure_ascii=False

            ),

            file=sys.stderr

        )

        sys.exit(1)