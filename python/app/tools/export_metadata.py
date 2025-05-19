import subprocess
import json
import os
from .convert import exr_to_jpg, mov_to_jpg
import pyseq

def export_metadata(date_path):
    # meta data 리스트, xlsx 파일의 한 행이 됌
    meta_list = []

    # 썸네일 저장 폴더 생성
    thumbnails_dir = os.path.join(date_path, "thumbnails")
    os.makedirs(thumbnails_dir, exist_ok=True)

    # .../scan/{date_path} 순회 하면서 Sequence 객체 get
    for root, dirs, seqs in pyseq.walk(date_path):
        dirs[:] = [d for d in dirs if d != "thumbnails"]
        scan_data_path = "" #exr의 첫 프레임 path / mov path
        for seq in seqs:
            _, ext = os.path.splitext(seq.name)
            # EXR sequnece 폴더의 첫 프레임으로 thumnail(JPG) 추출
            if ext == ".exr":
                scan_data_path = seq[0].path
                thumb_name = seq.head().strip(".") + ".jpg"
                thumb_path = os.path.join(thumbnails_dir, thumb_name)
                if not os.path.exists(thumb_path):
                    if exr_to_jpg(scan_data_path, thumb_path):
                        print(f"[OK] Thumbnail created: {thumb_path}")
                    else:
                        print(f"[FAIL] Thumbnail create failed: {scan_data_path}")
                        thumb_path = ""
                else:
                    print(f"[SKIP] Thumbnail already exist: {thumb_path}")
            # MOV 폴더의 첫 프레임 thumbnial(JPG) 추출
            elif ext == ".mov":
                mov_path = os.path.join(root, seq.name)
                scan_data_path = mov_path # for export meta data
                thumb_name = os.path.splitext(seq.name)[0] + ".jpg"
                thumb_path = os.path.join(thumbnails_dir, thumb_name)
                if not os.path.exists(thumb_path):
                    if mov_to_jpg(scan_data_path, thumb_path):
                        print(f"[OK] Thumbnail created: {thumb_path}")
                    else:
                        print(f"[FAIL] Thumbnail create failed: {scan_data_path}")
                        thumb_path = ""

                else:
                    print(f"[SKIP] Thumbnail already exist: {thumb_path}")

            else:
                print(f"[SKIP] {seq} is not EXR OR JPG")
                continue

            # 메타데이터 추출
            result = subprocess.run(
                ["exiftool", "-json", scan_data_path],
                capture_output=True, 
                text=True
            )

            # meta data json 파싱
            meta = json.loads(result.stdout)[0]
            meta["thumbnail_path"] = thumb_path
            meta_list.append(meta)
    return meta_list