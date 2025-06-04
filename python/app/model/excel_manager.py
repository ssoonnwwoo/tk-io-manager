from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import os
import re
import pyseq
import subprocess
import json

class ExcelManager:
    def __init__(self, iomanager_instance):
        self.date_directory_path = ""
        self.iomanager = iomanager_instance
    
    def set_date_path(self, date_directory_path):
        self.date_directory_path = date_directory_path

    def get_latest_xlsx(self):
        date_prefix = os.path.basename(self.date_directory_path)
        # ex ) date_prefix: 20250529
        
        pattern = re.compile(rf"{date_prefix}_list_v(\d{{3}})\.xlsx$")

        versioned_files = []
        for fname in os.listdir(self.date_directory_path):
            match = pattern.match(fname)
            if match:
                version = int(match.group(1))
                versioned_files.append((version, fname))
        # ex ) versioned_files: [(1, 20250529_list_v001.xlsx), (2, 20250529_list_v002.xlsx), ..., (14, 20250529_list_v014.xlsx)]

        meta_data_list = self.export_metadata(self.date_directory_path)
        
        if not versioned_files:
            # If there is no xlsx file. Make _v001.xlsx
            print("[INFO] No versioned file found. v001 will be created")
            new_xlsx_path = self.save_as_xlsx(meta_data_list)
            latest_xlsx_path = new_xlsx_path
        else:
            # return latest version filename
            latest_version_tuple = max(versioned_files)
            latest_xlsx_path = os.path.join(self.date_directory_path, latest_version_tuple[1])
        return latest_xlsx_path
    

    def export_metadata(self, date_path):
        # meta data list: one of xlsx rows 
        meta_list = []
        thumbnails_dir = os.path.join(date_path, "thumbnails")
        os.makedirs(thumbnails_dir, exist_ok=True)

        # .../scan/{date} 순회 하면서 Sequence 객체 get
        for root, dirs, seqs in pyseq.walk(date_path):
            dirs[:] = [d for d in dirs if d != "thumbnails"]
            first_frame_path = "" #exr의 첫 프레임 path / mov path
            for seq in seqs:
                _, ext = os.path.splitext(seq.name)
                # EXR sequnece 폴더의 첫 프레임을 thumnail(JPG)로 변환
                if ext == ".exr":
                    first_frame_path = seq[0].path
                    thumb_name = seq.head().strip(".") + ".jpg"
                    thumb_path = os.path.join(thumbnails_dir, thumb_name)
                    # If thumbnail not exist, make thumbnail
                    if not os.path.exists(thumb_path):
                        make_excel_thumbnail_result = self.make_excel_thumbnail(first_frame_path, thumb_path)
                        if not make_excel_thumbnail_result:
                            QMessageBox.warning(
                                self.iomanager,
                                "Error",
                                f"Error occurred while attempting to convert thumbnail.\nIO Manager skips the thumbnail creation process.",
                                QMessageBox.Ok
                            )
                            thumb_path = ""
                    # Skip making thumbnail if thumbnail exist
                    else:
                        print(f"[SKIP] Thumbnail already exist: {thumb_path}")
                # MOV 폴더의 첫 프레임 thumbnial(JPG) 추출
                elif ext == ".mov":
                    # mov를 shot 단위의 exr로 만들고 이후 과정은 똑같이
                    # mov_path = os.path.join(root, seq.name)
                    # scan_data_path = mov_path # for export meta data
                    # thumb_name = os.path.splitext(seq.name)[0] + ".jpg"
                    # thumb_path = os.path.join(thumbnails_dir, thumb_name)
                    # if not os.path.exists(thumb_path):
                    #     if mov_to_jpg(scan_data_path, thumb_path):
                    #         print(f"[OK] Thumbnail created: {thumb_path}")
                    #     else:
                    #         print(f"[FAIL] Thumbnail create failed: {scan_data_path}")
                    #         thumb_path = ""

                    # else:
                    #     print(f"[SKIP] Thumbnail already exist: {thumb_path}")
                    continue
                else:
                    print(f"[SKIP] {seq} is not EXR OR MOV")
                    continue

                # Exporting metadata using EXIF tool 
                result = subprocess.run(
                    ["exiftool", "-json", first_frame_path],
                    capture_output=True, 
                    text=True
                )
                # meta data json 파싱
                # ex) result(string) -> meta(json) 
                meta = json.loads(result.stdout)[0]
                meta["thumbnail_path"] = thumb_path
                meta_list.append(meta)
        return meta_list

    def make_excel_thumbnail(self, input_exr_path, output_jpg_path):        
        cmd = [
            "oiiotool",
            input_exr_path,
            "--colorconvert", "ACES", "sRGB",
            "-o", output_jpg_path
        ]

        result = subprocess.run(cmd, check=False)
        if result.returncode != 0:
            print(f"[ERROR] Error occurred while attempting to convert thumbnail")
            return False
        print(f"[COMPLETE] {input_exr_path} >>>>> {output_jpg_path}")
        return True



    def save_as_xlsx(self, meta_data_list, xl_name = None):
        """
        Save meta datas as xlsx using openpyxl

        meta_data : data, one row of xlsx rows
        meta_data_list = [meta_data1, meta_data2, ... , meta_dataN]
        """
        if not meta_data_list:
            QMessageBox.warning(self.iomanager, "No shot error", "No shot for exporting excel file")
            return None
        if not xl_name:
            xl_name = f"{os.path.basename(self.date_directory_path)}_list_v001.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Metadata"

        default_fields = ["thumbnail", "thumbnail_path", "shot_name", "seq_name"]
        all_keys_set = set()
        for meta_data in meta_data_list:
            for key in meta_data.keys():
                all_keys_set.add(key)

        all_fields = default_fields + list(all_keys_set) # headers
        ws.append(all_fields)

        # Insert meta data to work sheet
        for row_idx, meta_data in enumerate(meta_data_list, start=2): # Start at row2(Row1 : header)
            for col_idx, field in enumerate(all_fields, start=1):
                value = meta_data.get(field, "")
                # if value's type is list
                if isinstance(value, list):
                    value = ", ".join(str(v) for v in value)
                ws.cell(row=row_idx, column=col_idx, value=value)

            # Insert thumbnail to cell
            thumb_path = meta_data.get("thumbnail_path")
            if thumb_path and os.path.exists(thumb_path):
                col_letter = get_column_letter(all_fields.index("thumbnail") + 1)
                cell_ref = f"{col_letter}{row_idx}"
                img = Image(thumb_path)
                img.width = 192
                img.height = 108
                ws.add_image(img, cell_ref)
        xl_path = os.path.join(self.date_directory_path, xl_name)
        wb.save(xl_path)

        print(f"[COMPLETE] Metadata exported to:\n{xl_path}")
        return xl_path