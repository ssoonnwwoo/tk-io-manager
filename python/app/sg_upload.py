import os
from openpyxl import load_workbook
import sgtk
import re

logger = sgtk.platform.get_logger(__name__)

def get_publish_list(app_instance, xl_path):
    if os.path.exists(xl_path):
        wb = load_workbook(xl_path)
        ws = wb.active    

        # Make headers
        #headers ex : {'seq name': 0, 'shot name': 1, 'type': 2, 'Directory': 3}
        headers = {}
        for col_idx, cell in enumerate(ws[1]):
            headers[cell.value] = col_idx

        data = []
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start = 2):
            checked = row[headers["check"]]
            if not checked:
                continue
            seq = row[headers["seq name"]]
            shot = row[headers["shot name"]]
            # scan = row[headers["scan name"]]
            directory = row[headers["Directory"]]
            
            if not (seq and shot):
                logger.info(f"Not enough information in excel file. Row {row_idx} skipped.")
                continue
            shot_dir = os.path.join(app_instance.project_path, "seq", seq, shot, "plate/org")
            if not os.path.exists(shot_dir):
                versions = []
            else:
                versions = os.listdir(shot_dir)
            new_version = get_new_plate_version(versions)

            data.append({
                "seq": str(seq),
                "shot": str(shot),
                # "scan": str(scan),
                "version": new_version,
                "directory": str(directory)
            })
        return data

    else:
        data = []
        logger.warning(f"{xl_path} not exist")
        return data
    
def get_new_plate_version(dir_list):
    version_nums = []
    for name in dir_list:
        match = re.match(r"v(\d{3})", name)
        if match:
            version_nums.append(int(match.group(1)))
    if version_nums:
        latest_num = max(version_nums) + 1
        return f"v{latest_num:03d}"
    else:
        return "v001"
