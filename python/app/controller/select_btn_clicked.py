import os
from tank.platform.qt import QtGui

def on_select_clicked(main_widget):
    selected_date_path = select_date_directory(main_widget)
    if not selected_date_path :
        return
    main_widget.file_path_le.setText(selected_date_path)
    main_widget.excel_manager.set_date_path(selected_date_path)
    latest_xlsx_path = main_widget.excel_manager.get_latest_xlsx()
    # update table
    main_widget.update_table(latest_xlsx_path)
    # update label
    main_widget.excel_label.setText(latest_xlsx_path)

def select_date_directory(main_widget):
    current_path = main_widget.DEFAULT_PATH
    if not os.path.isdir(current_path):
        current_path = main_widget.DEFAULT_PATH

    selected_date_path = QtGui.QFileDialog.getExistingDirectory(
        main_widget, # Parent Widget
        "Select scan data folder",
        current_path, 
        QtGui.QFileDialog.ShowDirsOnly
    )

    # return date directory path
    # ex) date directory path : /home/rapa/show/my_project/product/scan/250529
    if selected_date_path:
        if not os.path.abspath(selected_date_path).startswith(os.path.abspath(main_widget.DEFAULT_PATH)):
            QtGui.QMessageBox.warning(
                main_widget, 
                "Warning", 
                f"'{selected_date_path}' Out of boundary : Not Selected Project"
            )
            return None
        return selected_date_path
    return None

