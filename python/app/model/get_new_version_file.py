import os
import re

def get_new_version_name(date_path):
    date_prefix = os.path.basename(os.path.normpath(date_path))
    
    pattern = re.compile(rf"{date_prefix}_list_v(\d{{3}})\.xlsx$")
    versions = []

    for fname in os.listdir(date_path):
        match = pattern.match(fname)
        if match:
            version = int(match.group(1))
            versions.append(version)

    next_version = max(versions) + 1
    file_name = f"{date_prefix}_list_v{next_version:03d}.xlsx"
    return os.path.join(date_path, file_name)
