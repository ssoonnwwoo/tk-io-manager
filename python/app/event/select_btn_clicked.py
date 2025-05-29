import os
import re
from tank.platform.qt import QtGui
for name, cls in QtGui.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls

def on_select_clicked(main_widget):
    selected_date_path = select_date_directory(main_widget)
    main_widget.file_path_le.setText(selected_date_path)
    latest_xlsx_path = get_latest_xlsx(selected_date_path)
    print(latest_xlsx_path)

def select_date_directory(main_widget):
    current_path = main_widget.DEFAULT_PATH
    if not os.path.isdir(current_path):
        current_path = main_widget.DEFAULT_PATH

    selected_date_path = QFileDialog.getExistingDirectory(
        main_widget, # Parent Widget
        "Select scan data folder",
        current_path, 
        QFileDialog.ShowDirsOnly
    )

    # return date directory path
    # ex) date directory path : /home/rapa/show/my_project/product/scan/250529
    if selected_date_path:
        if not os.path.abspath(selected_date_path).startswith(os.path.abspath(main_widget.DEFAULT_PATH)):
            QMessageBox.warning(
                main_widget, 
                "Warning", 
                f"'{selected_date_path}' Out of boundary : Not Selected Project"
            )
            return main_widget.DEFAULT_PATH
        return selected_date_path
    return main_widget.DEFAULT_PATH
############################################################################
# Model쪽으로 옮기고 다른 xlsx 만드는 코드(파이썬 파일)들하고 통합을 좀 하자
def get_latest_xlsx(date_path):
    date_prefix = os.path.basename(date_path)
    # ex ) date_prefix: 20250529
    
    pattern = re.compile(rf"{date_prefix}_list_v(\d{{3}})\.xlsx$")

    versioned_files = []
    for fname in os.listdir(date_path):
        match = pattern.match(fname)
        if match:
            version = int(match.group(1))
            versioned_files.append((version, fname))
    # ex ) versioned_files: [(1, 20250529_list_v001.xlsx), (2, 20250529_list_v002.xlsx), ..., (14, 20250529_list_v014.xlsx)]

    if not versioned_files:
        print("No versioned file found")
        return None
    
    # return latest version filename
    latest_version_tuple = max(versioned_files)
    return os.path.join(date_path, latest_version_tuple[1])