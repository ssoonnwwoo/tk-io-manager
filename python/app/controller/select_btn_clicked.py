import os
from tank.platform.qt import QtGui

def on_select_clicked(iomanager_ui):
    selected_date_path = select_date_directory(iomanager_ui)

    if not selected_date_path :
        return
    iomanager_ui.file_path_le.setText(selected_date_path)
    iomanager_ui.excel_manager.set_date_path(selected_date_path)
    latest_xlsx_path = iomanager_ui.excel_manager.get_latest_xlsx()
    # update table
    iomanager_ui.update_table(latest_xlsx_path)
    # update label
    iomanager_ui.excel_label.setText(latest_xlsx_path)

def select_date_directory(iomanager_ui):
    current_path = iomanager_ui.DEFAULT_PATH
    if not os.path.isdir(current_path):
        current_path = iomanager_ui.DEFAULT_PATH

    selected_date_path = QtGui.QFileDialog.getExistingDirectory(
        iomanager_ui, # Parent Widget
        "Select scan data folder",
        current_path, 
        QtGui.QFileDialog.ShowDirsOnly
    )

    # return date directory path
    # ex) date directory path : /home/rapa/show/my_project/product/scan/250529
    if selected_date_path:
        if not os.path.abspath(selected_date_path).startswith(os.path.abspath(iomanager_ui.DEFAULT_PATH)):
            QtGui.QMessageBox.warning(
                iomanager_ui, 
                "Warning", 
                f"'{selected_date_path}' Out of boundary : Not Selected Project"
            )
            return None
        return selected_date_path
    return None

