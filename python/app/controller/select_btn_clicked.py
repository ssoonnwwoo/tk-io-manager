import os
from tank.platform.qt import QtGui
for name, cls in QtGui.__dict__.items():
    if isinstance(cls, type): globals()[name] = cls

def on_select_clicked(main_widget):
    selected_date_path = select_date_directory(main_widget)
    if not selected_date_path :
        return
    main_widget.file_path_le.setText(selected_date_path)
    main_widget.excel_manager.set_date_path(selected_date_path)
    latest_xlsx_path = main_widget.excel_manager.get_latest_xlsx()
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
            return None
        return selected_date_path
    return None

############################################################################
# Model쪽으로 옮기고 다른 xlsx 만드는 코드(파이썬 파일)들하고 통합을 좀 하자
