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
            fps = row[headers["FramesPerSecond"]] or row[headers["CaptureRate"]]
            fps = round(float(fps), 4)
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
                "fps": fps,
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

def upload_shotgrid(app_instance, paths, seq, shot, ver):
    if all(os.path.exists(path) for path in paths.values()):
        mp4_path = paths["mp4"]
        webm_path = paths["webm"]
        filmstrip_path = paths["montage"]
        thumbnail_path = paths["thumbnail"]
        project_id = app_instance.sg.find_one("Project", [["name", "is", app_instance.project_name]], ["id"])
        # Create Sequence entity
        sequence_name = seq
        sequence = app_instance.sg.find_one("Sequence", [
            ["project", "is", project_id],
            ["code", "is", sequence_name]
        ], ["code"])

        if not sequence:
            sequence_data = {
                "project": project_id,
                "code": sequence_name
            }
            sequence = app_instance.sg.create("Sequence", sequence_data)
            print(f"Sequence created: {sequence['code']}")
        

        # Create Shot entity
        shot_name = f"{shot}_{ver}"
        shot_data = {
            "project": project_id,
            "code": shot_name,
            "sg_sequence": sequence
        }
        created_shot = app_instance.sg.create("Shot", shot_data)
        print(f"Shot created: {created_shot['code']}")

        # Thumbnail upload to shot
        app_instance.sg.upload_thumbnail("Shot", created_shot["id"], thumbnail_path)
        print(f"Thumbnail uploaded sucessfully")

        # Flimstrip upload to shot
        app_instance.sg.upload_filmstrip_thumbnail("Shot", created_shot["id"], filmstrip_path)
        print(f"Filmstrip uploaded sucessfully")

        # Create version entity and link to shot
        version_name = f"{shot}_plate_{ver}"
        version_data = {
            "project": project_id,
            "code": version_name,
            "entity": created_shot,
            "description": f"Description for {shot}"
        }
        version = app_instance.sg.create("Version", version_data)
        print(f"Version created: {version['code']}")

        app_instance.sg.upload("Version", version["id"], mp4_path, field_name="sg_uploaded_movie")
        print(f"MP4 uploaded successfully")

        app_instance.sg.upload("Version", version["id"], webm_path, field_name="sg_uploaded_movie_webm")
        print(f"WEBM uploaded successfully")

        # jpg, filmstrip upload to Version
        app_instance.sg.upload_thumbnail("Version", version["id"], thumbnail_path)
        app_instance.sg.upload_filmstrip_thumbnail("Version", version["id"], filmstrip_path)
        logger.info(f"{shot_name} publish completed")
        return True
            
    else:
        app_instance.show_error_dialog("Publish failed.")
        for path in [mp4_path, webm_path, filmstrip_path, thumbnail_path]:
            error_msg = f"{shot_name} publish failed.\nMissing: {path}"
            logger.error(error_msg)
            if not os.path.exists(path):
                print(error_msg)
        return False