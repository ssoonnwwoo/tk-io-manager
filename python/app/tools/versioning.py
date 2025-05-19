import os
import re
def get_latest_version_file(date_path):
    # 마지막 폴더 이름 → 파일 접두어로 사용
    date_prefix = os.path.basename(os.path.normpath(date_path))
    
    # 정규식: 20241226_2_list_v001.xlsx, v002.xlsx ...
    pattern = re.compile(rf"{date_prefix}_list_v(\d{{3}})\.xlsx$")

    versioned_files = []
    for fname in os.listdir(date_path):
        match = pattern.match(fname)
        if match:
            version = int(match.group(1))
            versioned_files.append((version, fname))

    if not versioned_files:
        return None

    # 가장 높은 버전 반환
    latest = max(versioned_files)
    return os.path.join(date_path, latest[1])