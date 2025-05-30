import os
import re

def get_latest_version_file(date_path):
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
        print("No versioned file found")
        return None

    # retrurn latest version file
    latest = max(versioned_files)
    return os.path.join(date_path, latest[1])

