import os
from tank.platform.qt import QtGui

def on_select_xl_clicked(iomanager_instance):
    file_path_le = iomanager_instance.file_path_le
    current_path = file_path_le.text()
    if not os.path.isdir(current_path):
        return

    file_path, _ = QtGui.QFileDialog.getOpenFileName(
        iomanager_instance,
        "Select XLSX file",
        current_path,
        "XLSX files (*.xlsx)"
    )

    if file_path:
        if not os.path.abspath(file_path).startswith(iomanager_instance.DEFAULT_PATH):
            QtGui.QMessageBox.warning(None, "Warning", f"Out of project boundary.\nPlease select in below path.\n{iomanager_instance.DEFAULT_PATH}")
            return
        iomanager_instance.update_table(file_path)
        iomanager_instance.excel_label.setText(file_path)
