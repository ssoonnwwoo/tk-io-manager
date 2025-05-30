import os
import re
import pyseq
import subprocess
import json

class ExcelManager:
    def __init__(self):
        self.date_directory_path = ""
    
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

        if not versioned_files:
            print("No versioned file found")
            latest_xlsx_path = None
        else :
            # return latest version filename
            latest_version_tuple = max(versioned_files)
            latest_xlsx_path = os.path.join(self.date_directory_path, latest_version_tuple[1])

        if not latest_xlsx_path:
            meta_data_list = self.export_metadata(self.date_directory_path)


        return os.path.join(self.date_directory_path, latest_version_tuple[1])
    
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
                        if not self.make_excel_thumbnail(first_frame_path, thumb_path):
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
                    pass
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
        import os
        print("현재 PATH:", os.environ.get("PATH"))



        if not os.path.isfile(input_exr_path):
            print(f"[ERROR] EXR does not exist: {input_exr_path}")
            return False
        
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
